from __future__ import annotations

import html as html_mod
from pathlib import Path
from typing import Optional

from docconvert.config import AppConfig
from docconvert.converters.base import BaseConverter
from docconvert.utils import clean_filename, decode_text, html_to_md


class DocConverter(BaseConverter):

    def __init__(self, config: Optional[AppConfig] = None):
        super().__init__(config)

    def convert(
        self,
        input_path: str,
        output_fmt: str,
        parent: Path,
        **kwargs,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        enhanced_md = kwargs.get('enhanced_md', False)
        stem_override = kwargs.get('stem_override')
        if kwargs.get('sheets') is not None:
            raise TypeError('DocConverter does not accept "sheets" parameter')
        results: list[tuple[str, str]] = []
        errors: list[tuple[str, str]] = []
        if self.cancel_check and self.cancel_check():
            return results, errors
        try:
            name, path = self._convert_doc(
                input_path, output_fmt, parent, enhanced_md, stem_override
            )
            results.append((name, path))
        except Exception as e:
            self.logger.error("Doc转换失败： %s", str(e))
            errors.append((Path(input_path).name, str(e)))
        return results, errors

    def _convert_doc(
        self, input_path: str, output_fmt: str, parent: Path,
        enhanced_md: bool = False, stem_override: Optional[str] = None,
    ) -> tuple[str, str]:
        import textract

        self.logger.info("提取文本： %s", input_path)
        raw_bytes = textract.process(input_path)
        text = decode_text(raw_bytes)
        stem = stem_override or clean_filename(Path(input_path).stem)
        source_name = Path(input_path).name

        content: object = ""
        if output_fmt == 'html':
            content = (
                f'<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n'
                f'    <meta charset="UTF-8">\n'
                f'    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
                f'    <title>{html_mod.escape(Path(input_path).stem)}</title>\n'
                f'    <style>\n'
                f'        body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 20px; line-height: 1.6; }}\n'
                f'        pre {{ white-space: pre-wrap; word-wrap: break-word; }}\n'
                f'    </style>\n</head>\n<body>\n<pre>\n{html_mod.escape(text)}\n</pre>\n</body>\n</html>'
            )
            output_name = f'{stem}.html'
        elif output_fmt == 'md':
            if enhanced_md:
                wrapped = f'<pre>\n{html_mod.escape(text)}\n</pre>'
                content = html_to_md(wrapped)
            else:
                content = text
            content = f'<!-- source: {source_name} | format: doc -->\n\n{content}'
            output_name = f'{stem}.md'
        elif output_fmt == 'json':
            content = {'source': source_name, 'content': text}
            output_name = f'{stem}.json'
        else:
            raise ValueError(f'不支持的输出格式： {output_fmt}')

        output_path = str(parent / output_name)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self._export(content, output_fmt))
        return output_name, output_path
