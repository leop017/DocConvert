from __future__ import annotations

import html as html_mod
import io
from pathlib import Path
from typing import Optional

import mammoth
from docx import Document as DocxDocument

from docconvert.config import AppConfig
from docconvert.converters.base import BaseConverter
from docconvert.cleaners import WordMdCleaner
from docconvert.utils import clean_filename, html_to_md


class WordConverter(BaseConverter):

    def __init__(self, config: Optional[AppConfig] = None):
        super().__init__(config)
        self.md_cleaner = WordMdCleaner(self.config)

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
            raise TypeError('WordConverter does not accept "sheets" parameter')
        results: list[tuple[str, str]] = []
        errors: list[tuple[str, str]] = []
        if self.cancel_check and self.cancel_check():
            return results, errors
        try:
            name, path = self._convert_word(
                input_path, output_fmt, parent, enhanced_md, stem_override
            )
            results.append((name, path))
        except Exception as e:
            self.logger.error("Word转换失败： %s", str(e))
            errors.append((Path(input_path).name, str(e)))
        return results, errors

    def _convert_word(
        self, input_path: str, output_fmt: str, parent: Path,
        enhanced_md: bool = False, stem_override: Optional[str] = None,
    ) -> tuple[str, str]:
        self.logger.info("读取文档： %s", input_path)
        doc = DocxDocument(input_path)
        stem = stem_override or clean_filename(Path(input_path).stem)

        self._strip_headers_footers(doc)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        content: object = ""
        paragraph_count = len([p for p in doc.paragraphs if p.text.strip()])

        if output_fmt == 'html':
            content = self._to_html(buf, input_path)
        elif output_fmt == 'md':
            content = self._to_md(buf, input_path, paragraph_count, enhanced_md)
        elif output_fmt == 'json':
            content = self._to_json(buf, input_path, paragraph_count)
        else:
            raise ValueError(f'不支持的输出格式： {output_fmt}')

        output_name = f'{stem}_{clean_filename("doc")}.{output_fmt}'
        output_path_full = str(parent / output_name)
        with open(output_path_full, 'w', encoding='utf-8') as f:
            f.write(self._export(content, output_fmt))
        return output_name, output_path_full

    @staticmethod
    def _strip_headers_footers(doc: object):
        body_xml = doc.element.body  # type: ignore[attr-defined]
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        for sect_pr in body_xml.findall(f'.//{ns}sectPr'):
            for child in list(sect_pr):
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag in ('headerReference', 'footerReference'):
                    sect_pr.remove(child)

    def _to_html(self, buf: io.BytesIO, input_path: str) -> str:
        result = mammoth.convert_to_html(buf)
        content = result.value
        title = Path(input_path).stem
        return (
            f'<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n'
            f'    <meta charset="UTF-8">\n'
            f'    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
            f'    <title>{html_mod.escape(title)}</title>\n'
            f'    <style>\n'
            f'        body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 20px; line-height: 1.6; }}\n'
            f'        table {{ border-collapse: collapse; width: 100%; }}\n'
            f'        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; white-space: pre-wrap; }}\n'
            f'        img {{ max-width: 100%; }}\n'
            f'    </style>\n</head>\n<body>\n{content}\n</body>\n</html>'
        )

    def _to_md(self, buf: io.BytesIO, input_path: str, paragraph_count: int, enhanced_md: bool = False) -> str:
        if enhanced_md:
            result = mammoth.convert_to_html(buf)
            content = html_to_md(result.value)
        else:
            result = mammoth.convert_to_markdown(buf)
            content = result.value
        content = self.md_cleaner.clean(content)
        header = (
            f'<!-- source: {Path(input_path).name}'
            f' | paragraphs: {paragraph_count} -->\n\n'
        )
        return header + content

    def _to_json(self, buf: io.BytesIO, input_path: str, paragraph_count: int) -> dict:
        result = mammoth.convert_to_html(buf)
        return {
            'metadata': {
                'source': Path(input_path).name,
                'format': 'docx',
                'paragraphs': paragraph_count,
            },
            'content': result.value,
        }
