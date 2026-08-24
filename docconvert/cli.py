from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docconvert.config import DEFAULT_CONFIG
from docconvert.controller import ConversionController
from docconvert.logger import setup_logging
from docconvert.models import ProgressEvent


def _cli_progress(event: ProgressEvent):
    if event.message:
        print(f'\r[{int(event.progress * 100):3d}%] {event.message:<50s}', file=sys.stderr, end='')
    if event.done:
        print(file=sys.stderr)


class _HelpFormatter(argparse.HelpFormatter):
    """Preserves newlines in description, epilog, and all help strings."""

    def _fill_text(self, text, width, indent):
        if text:
            return ''.join(indent + line + '\n' for line in text.splitlines())
        return ''

    def _split_lines(self, text, width):
        return text.splitlines() if text else []


def main_cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description='DocConvert - 文档转换工具 (CLI)',
        formatter_class=_HelpFormatter,
        epilog=(
            '示例:\n'
            '  python main.py convert input.xlsx --format md\n'
            '  python main.py convert input.docx --format html -o ./output\n'
            '  python main.py convert file1.xlsx file2.xlsx --format json\n'
        ),
    )

    subparsers = parser.add_subparsers(dest='command', help='子命令')

    convert_parser = subparsers.add_parser('convert', help='转换文件')
    convert_parser.add_argument('files', nargs='+', help='输入文件路径')
    convert_parser.add_argument(
        '--format', '-f',
        choices=['html', 'md', 'json'],
        default='html',
        help='输出格式 （默认： html)',
    )
    convert_parser.add_argument(
        '--output', '-o',
        default=None,
        help='输出目录 （默认： 输入文件所在目录）',
    )
    convert_parser.add_argument(
        '--enhanced', '-e',
        action='store_true',
        help='启用增强 Markdown 输出',
    )
    convert_parser.add_argument(
        '--sheet', '-s',
        action='append',
        help='指定 Excel 工作表 （可重复使用， 默认： 全部）',
    )
    convert_parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='详细日志输出',
    )

    args = parser.parse_args(argv)

    if args.command != 'convert':
        parser.print_help()
        return 1

    setup_logging(level='DEBUG' if getattr(args, 'verbose', False) else 'INFO')

    controller = ConversionController(DEFAULT_CONFIG)

    all_results: list[tuple[str, str, str | None]] = []
    failed = 0

    for filepath in args.files:
        p = Path(filepath)
        if not p.exists():
            print(f'错误： 文件不存在 - {filepath}', file=sys.stderr)
            all_results.append((p.name, '', '文件不存在'))
            failed += 1
            continue

        ext = p.suffix.lower()
        if ext not in {'.xlsx', '.xls', '.docx', '.doc'}:
            print(f'错误： 不支持的文件格式 - {filepath}', file=sys.stderr)
            all_results.append((p.name, '', '不支持的文件格式'))
            failed += 1
            continue

        try:
            convert_results = controller.convert_files(
                files=[filepath],
                output_fmt=args.format,
                output_dir=args.output,
                enhanced_md=args.enhanced,
                sheets=args.sheet,
                progress_callback=_cli_progress,
            )
            for name, path, err in convert_results:
                all_results.append((name, path, err))
                if err:
                    failed += 1
                    print(f'失败： {name} - {err}', file=sys.stderr)
                else:
                    print(f'成功： {name} -> {path}')
        except Exception as e:
            print(f'错误： {filepath} - {e}', file=sys.stderr)
            all_results.append((p.name, '', str(e)))
            failed += 1

    if failed:
        print(f'\n处理完成： {len(all_results)} 个文件， {failed} 个失败')
        return 1

    print(f'\n处理完成： {len(all_results)} 个文件， 全部成功')
    return 0


if __name__ == '__main__':
    sys.exit(main_cli())
