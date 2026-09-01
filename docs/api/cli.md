# `docconvert.cli` — 命令行接口

> 源码: [`docconvert/cli.py`](../../docconvert/cli.py)
> 入口点(由 `pyproject.toml` 注册): `docconvert = "docconvert.cli:main_cli"`

CLI 是 `ConversionController` 的最薄壳层 — 仅负责参数解析、进度打印和退出码。

---

## `main_cli(argv=None)`

```python
def main_cli(argv: list[str] | None = None) -> int: ...
```

| 参数 | 说明 |
|------|------|
| `argv` | 命令行参数列表;`None` 时用 `sys.argv[1:]` |
| **返回** | 退出码:`0` = 全部成功,`1` = 有失败或用法错误 |

### 子命令:`convert`

```text
docconvert convert <files...> [--format {html,md,json}]
                              [--output DIR] [--enhanced]
                              [--sheet NAME]...
                              [--verbose] [--version]
```

| 参数 | 说明 |
|------|------|
| `files` (位置) | 一个或多个输入文件路径;支持 `.xlsx` / `.xls` / `.docx` / `.doc` |
| `--format` / `-f` | 输出格式,默认 `html` |
| `--output` / `-o` | 输出目录;默认与每个输入文件同目录(实际由 `ConversionController._batch_parent` 解析) |
| `--enhanced` / `-e` | 对 Word/Doc 启用 enhanced MD 路径(HTML 中转 + 清洗) |
| `--sheet` / `-s` | Excel 工作表过滤,**可重复**;不指定则全部 |
| `--verbose` / `-v` | DEBUG 日志 |
| `--version` | 打印 `DocConvert <version>` 后退出 |

### 进度显示

通过 `_cli_progress` 回调把 `ProgressEvent` 实时打印到 **stderr**:

```text
[ 17%] 处理文件 1/3: report.xlsx
[ 33%] 完成 1/3
[ 50%] 处理文件 2/3: manual.docx
[100%] 转换完成
```

### 用法示例

```bash
# 转单个 Excel
docconvert convert report.xlsx --format md

# 转 Word 到自定义目录
docconvert convert manual.docx --format html -o ./output

# 批量转多个文件
docconvert convert a.xlsx b.docx c.xls --format json -o ./out

# 只转某些 sheet
docconvert convert finance.xlsx --format md -s Q1 -s Q2

# 详细日志
docconvert convert big.xlsx --format md --verbose
```

### 退出码

| 情况 | 退出码 |
|------|--------|
| 全部成功 | `0` |
| 至少一个文件失败 / 找不到 / 不支持的扩展名 | `1` |
| 用法错误(无 `convert` 子命令) | `1`(打印 help) |

### 入口点

`pyproject.toml` 中:

```toml
[project.scripts]
docconvert = "docconvert.cli:main_cli"
```

直接 `pip install docconvert-local` 后,`docconvert` 命令即可在任意 shell 中使用。

---

## `main.py` 中的派发

> 源码: [`main.py`](../../main.py)

```python
import sys
from docconvert.cli import main_cli

if __name__ == '__main__':
    if len(sys.argv) == 1:
        from docconvert.gui.app import DocConvertApp
        DocConvertApp().run()
    elif sys.argv[1] == 'convert':
        sys.exit(main_cli(sys.argv[2:]))
    else:
        # 透传给 CLI 解析
        sys.exit(main_cli(sys.argv[1:]))
```

- `python main.py` → 启动 GUI
- `python main.py convert ...` → 走 CLI
- `python main.py --help` → CLI help

---

## `_cli_progress(event)`

```python
def _cli_progress(event: ProgressEvent) -> None: ...
```

CLI 默认的进度回调。把 `event.message` 同行覆盖式打印到 stderr,`event.done=True` 时换行。

> 不直接面向用户,通常无需覆盖。
