from __future__ import annotations

import os
from collections.abc import Callable
from pathlib import Path
from threading import Event, Lock, Thread
from typing import Optional

from docconvert.config import DEFAULT_CONFIG, AppConfig
from docconvert.converters import DocConverter, ExcelConverter, WordConverter
from docconvert.converters.base import BaseConverter
from docconvert.logger import get_logger
from docconvert.models import ProgressEvent
from docconvert.utils import clean_filename, get_excel_sheet_names

ProgressCallback = Callable[[ProgressEvent], None]


class ConversionController:

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or DEFAULT_CONFIG
        self.logger = get_logger()
        self._cancel_event = Event()
        self._done_event = Event()
        self._thread: Optional[Thread] = None
        self._running = False
        self._lock = Lock()
        self.last_results: list[tuple[str, str, Optional[str]]] = []
        self.was_cancelled: bool = False
        self.last_error: Optional[str] = None

    @property
    def is_running(self) -> bool:
        return self._running

    def cancel(self):
        self._cancel_event.set()

    def reset_cancel(self):
        self._cancel_event.clear()

    def set_config(self, config: AppConfig):
        """Replace the active config. Should only be called when idle."""
        if self._running:
            self.logger.warning("无法在转换进行中更新 config,已忽略")
            return False
        self.config = config
        return True

    @staticmethod
    def _batch_parent(files: list[str], output_dir: Optional[str] = None) -> Path:
        """Return the single directory a batch converts into.

        Every output of a batch goes to one directory: the explicit
        ``output_dir`` or, when omitted, the first input file's parent.
        ``convert_files`` and ``check_overwrite_paths`` must agree on this
        value or the overwrite confirmation will not match reality, so both
        resolve it through this helper.
        """
        if output_dir:
            return Path(output_dir)
        return Path(files[0]).resolve().parent if files else Path.cwd()

    @staticmethod
    def _make_unique_stem(base_stem: str, source_path: str, used: set[str]) -> str:
        """Return a batch-unique output stem.

        All files in a batch are written to a single directory (either the
        explicit ``output_dir`` or the first input file's parent), so two
        input files that share a basename would produce the same output
        path and the second would silently overwrite the first. Disambiguate
        subsequent collisions with the source's parent-directory name, then a
        numeric suffix as a last resort. ``used`` is mutated in place.
        """
        if base_stem not in used:
            used.add(base_stem)
            return base_stem
        parent_name = clean_filename(Path(source_path).resolve().parent.name)
        suffix = parent_name if parent_name and parent_name != base_stem else ''
        candidate = f'{base_stem}_{suffix}' if suffix else base_stem
        n = 2
        while candidate in used:
            candidate = f'{base_stem}_{suffix}_{n}' if suffix else f'{base_stem}_{n}'
            n += 1
        used.add(candidate)
        return candidate

    def _get_converter(self, ext: str) -> BaseConverter:
        if ext in ('.xlsx', '.xls'):
            c: BaseConverter = ExcelConverter(self.config)
        elif ext == '.docx':
            c = WordConverter(self.config)
        elif ext == '.doc':
            c = DocConverter(self.config)
        else:
            raise ValueError(f'不支持的文件格式： {ext}')
        c.cancel_check = self._cancel_event.is_set
        return c

    def convert_files(
        self,
        files: list[str],
        output_fmt: str,
        output_dir: Optional[str] = None,
        enhanced_md: bool = False,
        sheets: Optional[list[str]] = None,
        progress_callback: Optional[ProgressCallback] = None,
    ) -> list[tuple[str, str, Optional[str]]]:
        """
        Convert files synchronously.
        Returns list of (filename, output_path, error_or_None).

        Sets ``self.was_cancelled = True`` if the loop was aborted because the
        cancel event was set mid-batch. The caller is expected to surface this
        to the user (e.g., show a partial-results dialog instead of a green
        "completed" badge).
        """
        self.reset_cancel()
        self.was_cancelled = False
        results: list[tuple[str, str, Optional[str]]] = []

        if not files:
            return results

        parent = self._batch_parent(files, output_dir)
        if output_dir and not parent.is_dir():
            try:
                parent.mkdir(parents=True, exist_ok=True)
                self.logger.info("已创建输出目录： %s", parent)
            except OSError as e:
                raise ValueError(
                    f'无法创建输出目录： {parent} - {e}'
                ) from e

        total = len(files)
        used_stems: set[str] = set()
        for idx, input_path in enumerate(files):
            if self._cancel_event.is_set():
                self.was_cancelled = True
                self.logger.info("转换已取消")
                break

            if not os.path.exists(input_path):
                err = '文件不存在'
                fname_missing = Path(input_path).name
                results.append((fname_missing, '', err))
                self._report(progress_callback, ProgressEvent(
                    message=f'文件不存在： {fname_missing}',
                    error=err,
                    progress=(idx + 1) / total,
                ))
                continue

            fname = Path(input_path).name
            self._report(progress_callback, ProgressEvent(
                message=f'处理文件 {idx + 1}/{total}: {fname}',
                progress=(idx + 0.5) / total,
            ))

            ext = Path(input_path).suffix.lower()
            base_stem = clean_filename(Path(input_path).stem)
            unique_stem = self._make_unique_stem(base_stem, input_path, used_stems)
            try:
                converter = self._get_converter(ext)
                # ``sheets`` is Excel-only; passing it to Word/Doc converters
                # raises TypeError (they explicitly reject it), which would
                # otherwise fail every Word/Doc file in a mixed batch.
                convert_kwargs = {
                    'enhanced_md': enhanced_md,
                    'stem_override': unique_stem,
                }
                if ext in ('.xlsx', '.xls'):
                    convert_kwargs['sheets'] = sheets
                file_results, file_errors = converter.convert(
                    input_path, output_fmt, parent, **convert_kwargs
                )
                for name, path in file_results:
                    results.append((name, path, None))
                    self.logger.info("转换成功： %s -> %s", name, path)
                for name, err in file_errors:
                    results.append((name, '', err))
                    self.logger.warning("转换失败 [%s]: %s", name, err)
                    self._report(progress_callback, ProgressEvent(
                        message=f'转换失败： {name}',
                        error=err,
                        progress=(idx + 1) / total,
                    ))
            except Exception as e:
                self.logger.error("转换异常 [%s]: %s", fname, str(e))
                results.append((fname, '', str(e)))
                self._report(progress_callback, ProgressEvent(
                    message=f'转换异常： {fname}',
                    error=str(e),
                    progress=(idx + 1) / total,
                ))

            self._report(progress_callback, ProgressEvent(
                message=f'完成 {idx + 1}/{total}',
                progress=(idx + 1) / total,
            ))
            if self._cancel_event.is_set():
                self.was_cancelled = True

        if self.was_cancelled:
            self._report(progress_callback, ProgressEvent(
                message='转换已取消',
                done=True,
            ))
        else:
            self._report(progress_callback, ProgressEvent(
                message='转换完成',
                progress=1.0,
                done=True,
            ))
        return results

    def convert_files_async(
        self,
        files: list[str],
        output_fmt: str,
        output_dir: Optional[str] = None,
        enhanced_md: bool = False,
        sheets: Optional[list[str]] = None,
        progress_callback: Optional[ProgressCallback] = None,
    ):
        """Start conversion in background thread."""
        with self._lock:
            if self._running:
                self.logger.warning("转换任务已在运行中")
                return False
            self._running = True
            self._cancel_event.clear()
            self._done_event.clear()
            self.was_cancelled = False
            self.last_error = None

        def _run():
            try:
                self.last_results = self.convert_files(
                    files=files,
                    output_fmt=output_fmt,
                    output_dir=output_dir,
                    enhanced_md=enhanced_md,
                    sheets=sheets,
                    progress_callback=progress_callback,
                )
            except Exception as e:
                self.logger.error("转换任务异常： %s", str(e))
                self.last_results = []
                self.last_error = str(e)
            finally:
                with self._lock:
                    self._running = False
                self._done_event.set()

        self._thread = Thread(target=_run, daemon=True)
        self._thread.start()
        return True

    def wait_for_completion(self, timeout: Optional[float] = None) -> bool:
        """Wait for the background thread to complete. Returns True if completed."""
        if self._thread is None:
            return True
        self._thread.join(timeout=timeout)
        return not self._running

    def wait_done(self, timeout: Optional[float] = None) -> bool:
        """Event-based wait for completion. Returns True if done within timeout."""
        if self._thread is None:
            return True
        return self._done_event.wait(timeout=timeout)

    def _report(
        self,
        callback: Optional[ProgressCallback],
        event: ProgressEvent,
    ):
        if callback:
            try:
                callback(event)
            except Exception as e:
                self.logger.debug("Progress callback error: %s", e)

    def check_overwrite_paths(
        self,
        files: list[str],
        output_fmt: str,
        output_dir: Optional[str] = None,
        sheets: Optional[list[str]] = None,
    ) -> list[str]:
        """Pre-compute output paths and return existing ones.

        Uses the same batch directory resolution as ``convert_files``
        (``_batch_parent``), so the reported paths are exactly the ones a
        conversion run would write — including when no ``output_dir`` is
        given and a multi-directory batch all lands in the first file's
        parent.
        """
        existing: list[str] = []
        used_stems: set[str] = set()
        parent = self._batch_parent(files, output_dir)
        for input_path in files:
            if not os.path.exists(input_path):
                continue
            base_stem = clean_filename(Path(input_path).stem)
            unique_stem = self._make_unique_stem(base_stem, input_path, used_stems)
            try:
                paths = self._compute_output_paths(
                    input_path, output_fmt, str(parent), sheets,
                    stem_override=unique_stem,
                )
            except Exception as e:
                self.logger.warning(
                    "无法计算输出路径 [%s]: %s （转换时将报错）", input_path, e
                )
                continue
            for p in paths:
                if os.path.exists(p):
                    existing.append(p)
        return existing

    def _compute_output_paths(
        self,
        input_path: str,
        output_fmt: str,
        output_dir: Optional[str] = None,
        sheets: Optional[list[str]] = None,
        stem_override: Optional[str] = None,
    ) -> list[str]:
        """Return the list of output paths that ``input_path`` would produce.

        Order:
        - .docx → [stem_doc.<fmt>]
        - .doc  → [stem.<fmt>]
        - .xlsx / .xls → [stem_<sheet>.<fmt>] per requested sheet, or
          [stem_<sheet>.<fmt>] for every sheet in the workbook when
          ``sheets`` is None.
        - other → [] (unsupported extension)
        """
        from docconvert.utils import clean_filename

        if not os.path.exists(input_path):
            return []

        if output_dir:
            parent = Path(output_dir)
        else:
            parent = Path(input_path).parent

        stem = Path(input_path).stem
        stem_clean = stem_override or clean_filename(stem)
        ext = Path(input_path).suffix.lower()

        if ext == '.docx':
            return [str(parent / f'{stem_clean}_{clean_filename("doc")}.{output_fmt}')]
        if ext == '.doc':
            return [str(parent / f'{stem_clean}.{output_fmt}')]
        if ext in ('.xlsx', '.xls'):
            if not sheets:
                sheets = get_excel_sheet_names(input_path, ext)
            # Mirror ExcelConverter's per-sheet suffix dedup so the
            # reported paths are the ones a conversion run actually writes.
            from docconvert.utils import unique_cleaned_suffixes
            clean_names = unique_cleaned_suffixes(sheets)
            return [
                str(parent / f'{stem_clean}_{cn}.{output_fmt}')
                for cn in clean_names
            ]
        return []
