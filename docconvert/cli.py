from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docconvert.__version__ import __version__
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
        description='DocConvert — convert Excel & Word documents to clean Markdown, HTML, or JSON.',
        formatter_class=_HelpFormatter,
        epilog=(
            'Examples:\n'
            '  python main.py convert input.xlsx --format md\n'
            '  python main.py convert input.docx --format html -o ./output\n'
            '  python main.py convert file1.xlsx file2.xlsx --format json\n'
        ),
    )

    subparsers = parser.add_subparsers(dest='command', help='Command to run')

    convert_parser = subparsers.add_parser('convert', help='Convert one or more files')
    convert_parser.add_argument('files', nargs='+', help='Input file paths')
    convert_parser.add_argument(
        '--format', '-f',
        choices=['html', 'md', 'json'],
        default='html',
        help='Output format (default: html)',
    )
    convert_parser.add_argument(
        '--output', '-o',
        default=None,
        help='Output directory (default: same as input file)',
    )
    convert_parser.add_argument(
        '--enhanced', '-e',
        action='store_true',
        help='Enable enhanced Markdown cleaning (remove page numbers, duplicate headers, etc.)',
    )
    convert_parser.add_argument(
        '--sheet', '-s',
        action='append',
        help='Select Excel sheet(s) (repeatable; default: all sheets)',
    )
    parser.add_argument(
        '--version',
        action='version',
        version=f'DocConvert {__version__}',
    )
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose debug logging',
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
            print(f'Error: file not found — {filepath}', file=sys.stderr)
            all_results.append((p.name, '', 'File not found'))
            failed += 1
            continue

        ext = p.suffix.lower()
        if ext not in {'.xlsx', '.xls', '.docx', '.doc'}:
            print(f'Error: unsupported format — {filepath}', file=sys.stderr)
            all_results.append((p.name, '', 'Unsupported format'))
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
                    print(f'Failed: {name} — {err}', file=sys.stderr)
                else:
                    print(f'OK: {name} -> {path}')
        except Exception as e:
            print(f'Error: {filepath} — {e}', file=sys.stderr)
            all_results.append((p.name, '', str(e)))
            failed += 1

    if failed:
        print(f'\nDone: {len(all_results)} file(s), {failed} failed')
        return 1

    print(f'\nDone: {len(all_results)} file(s), all succeeded')
    return 0


if __name__ == '__main__':
    sys.exit(main_cli())
