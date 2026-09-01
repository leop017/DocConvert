# DocConvert

> 从 Excel & Word 中提取干净的 Markdown 和结构化数据 —— 专为 RAG 管道和 LLM 工作流设计。
> 100% 离线运行，无需 API 密钥，数据不离开你的机器。

[![PyPI version](https://img.shields.io/pypi/v/docconvert-local.svg?cachebust=1725084000)](https://pypi.org/project/docconvert-local/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![Build](https://github.com/leop017/DocConvert/actions/workflows/ci.yml/badge.svg)](https://github.com/leop017/DocConvert/actions)
[![Stars](https://img.shields.io/github/stars/leop017/DocConvert?style=social)](https://github.com/leop017/DocConvert/stargazers)

**English README**: [README.md](README.md)

## 为什么选择 DocConvert

大多数文档转 Markdown 的工具只做转换——并不**清洗**。原始输出中充斥着页码、重复标题和多余空白，会浪费 context window 并稀释检索质量。

DocConvert 专为需要将文档喂给 LLM 的用户而设计，确保每一个 token 都有价值。

| 功能                              |       DocConvert      | MarkItDown | Pandas + python-docx |
| -------------------------------- | :-------------------: | :--------: | :------------------: |
| 支持旧版 `.doc` 文件              |           ✅           |      ❌     |           ❌          |
| Excel 合并单元格（rowspan/colspan）|           ✅           |     ⚠️     |        手动           |
| 可配置清洗管线                    | ✅ 4 条规则，任意开关 |      ❌     |           ❌          |
| 批量转换 + 指定工作表              |           ✅           |      ✅     |           ❌          |
| 桌面 GUI（无需终端）               |           ✅           |      ❌     |           ❌          |
| PDF / PPT / 音频支持               |           ❌           |      ✅     |           ❌          |
| MCP Server / Claude 集成           |           ❌           |      ✅     |           ❌          |
| 100% 离线，无云服务依赖            |           ✅           |      ✅     |           ✅          |

**选择 DocConvert 如果：** 你在组织内部处理 Excel/Word 文档，需要旧版 `.doc` 支持，或希望在将文档喂入 RAG 系统前运行可配置的清洗管线。

**选择 MarkItDown 如果：** 你需要 PDF、PPT、图片或音频转换，或想要开箱即用的 MCP/Claude Desktop 集成。

## 使用场景

* **RAG 数据摄入** —— 将 Excel 财务报表和 Word 合同转换为干净 Markdown，直接用于 embedding

* **LLM 上下文预处理** —— 去除页码、重复标题和噪声后再进行分块

* **离线合规** —— 转换敏感文档，无需上传到任何云服务

* **批量自动化** —— 将整个文件夹的报告批量转换为结构化目录

## 安装

```bash
pip install docconvert-local
```

可选扩展：

```bash
# 旧版 .doc 支持（仅 Linux / macOS）
pip install docconvert-local[doc]

# 完整功能集（含构建工具）
pip install docconvert-local[all]
```

## 快速开始

### GUI（图形界面）

```bash
python main.py
```

### CLI（批量 / 脚本）

```bash
# 单个文件 → 干净 Markdown
python main.py convert input.xlsx --format md

# 批量转换 + 增强清洗（推荐用于 RAG）
python main.py convert input.docx --format md --enhanced

# 多个文件 → HTML 输出到指定目录
python main.py convert file1.xlsx file2.docx --format html -o ./output

# 指定 Excel 工作表 → JSON
python main.py convert input.xlsx --format json --sheet "Sheet1" --sheet "Sheet2"
```

### Python API

```python
from docconvert.controller import ConversionController
from docconvert.config import DEFAULT_CONFIG

controller = ConversionController(DEFAULT_CONFIG)
results = controller.convert_files(
    files=["input.xlsx", "report.docx"],
    output_fmt="md",
    output_dir="./out",
    enhanced_md=True,
)

for name, path, error in results:
    if error:
        print(f"失败: {name} – {error}")
    else:
        print(f"成功: {name} → {path}")
```

完整的 API 参考和扩展指南见 [API 文档](docs/api/index.md)：

* [模块总览](docs/api/index.md#模块总览) ·
* [CLI 参考](docs/api/cli.md) ·
* [ConversionController](docs/api/controller.md) ·
* [转换器 / 导出器 / 清洗器](docs/api/converters.md) ·
* [Config & Models](docs/api/config.md) ·
* [GUI (Tkinter)](docs/api/gui.md)

### RAG 管道集成

```python
from docconvert.controller import ConversionController
from docconvert.config import DEFAULT_CONFIG
from langchain.text_splitter import RecursiveCharacterTextSplitter

controller = ConversionController(DEFAULT_CONFIG)
docs = []
for name, path, error in controller.convert_files(
    files=["contracts/*.docx"],
    output_fmt="md",
    output_dir="./out",
    enhanced_md=True,
):
    if not error:
        with open(path) as f:
            docs.append(f.read())

# 分块 + embedding —— 噪声已去除
splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=100)
chunks = splitter.split_text("\n".join(docs))
```

## 输出预览

**输入** —— 一个含合并单元格的 Excel 表格：

|   区域  |  Q1 |  Q2 |
| :-----: | :-: | :-: |
| **北**  | 120 | 150 |
| **南**  |  90 | 200 |

→ **Markdown 输出**（自动清洗）：

```markdown
## 区域    Q1    Q2
北         120   150
南          90   200
```

→ **JSON 输出**：

```json
{
  "区域": ["北", "南"],
  "Q1": [120, 90],
  "Q2": [150, 200]
}
```

## 智能清洗管线

`--enhanced` 标志会运行一个可配置的清洗阶段，在输出前移除常见的文档噪声。每条规则可独立开关：

```python
from docconvert.config import AppConfig

config = AppConfig(
    cleaning_rules={
        "remove_page_numbers": True,       # 移除页码，如 1, 2, 3… 和"第N页"
        "remove_duplicate_headers": True,  # 去重重复的章节标题
        "remove_empty_lines": True,        # 压缩多余空行
        "normalize_spaces": True,          # 合并多余空白，保留表格结构
    }
)
```

开启 `--enhanced` 时四条规则默认全部启用。将任意规则设为 `False` 可保留原始输出。

## 功能特性

* **Excel** —— 工作表选择、合并单元格（rowspan/colspan）、HTML / Markdown / JSON 输出

* **Word** —— `.docx` 通过 `python-docx` + `mammoth`，旧版 `.doc` 通过 `textract`

* **智能 Markdown** —— 移除页码、重复标题、压缩空行；所有规则可配置

* **GUI** —— Tkinter 桌面应用，支持文件列表、预览、进度条、覆盖保护

* **CLI** —— 通过 argparse 一行命令批量转换

* **Python API** —— 程序化控制，完整类型提示

* **独立可执行文件** —— 下载 Windows / macOS / Linux 的 `.exe`，无需安装 Python

## 发布版本（无需 Python）

每次 tag push 后会自动构建 Windows、macOS 和 Linux 的独立可执行文件。从 [Releases](https://github.com/leop017/DocConvert/releases) 下载。

## 项目结构

```
docconvert/
  converters/     # Excel / Word / .doc 读取器
  cleaners/       # Markdown 清洗管线
  exporters/      # HTML / Markdown / JSON 输出
  controller/     # 调度、异步、覆盖检测
  gui/            # Tkinter 桌面应用
  parsers/, chunkers/  # 扩展点
tests/
main.py           # GUI / CLI 入口
```

## 贡献

参见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 许可证

MIT
