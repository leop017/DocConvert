# DocConvert Code Wiki

> **项目**：DocConvert v2.0.0
> **描述**：Excel（`.xlsx`/`.xls`）和 Word（`.docx`/`.doc`）文档转 HTML / Markdown / JSON 的多格式转换工具
> **语言**：Python ≥ 3.10
> **许可证**：MIT

---

## 一、项目概览

DocConvert 是一个文档格式转换工具，支持将 Excel 和 Word 文件转换为 HTML、Markdown 和 JSON 三种格式。项目采用**插件化架构**，通过抽象基类（ABC）定义转换器、清洗器、导出器的接口，便于扩展新的输入格式或输出格式。

### 核心能力

| 输入格式 | HTML | Markdown | JSON |
|----------|------|----------|------|
| `.xlsx`  | ✅   | ✅       | ✅   |
| `.xls`   | ✅   | ✅       | ✅   |
| `.docx`  | ✅   | ✅       | ✅   |
| `.doc`   | ✅   | ✅       | ✅   |

**特色功能**：
- Excel 合并单元格（rowspan/colspan）完整支持
- Word Markdown 输出多级清洗管道（页码移除、重复页眉去重、空行折叠、空白归一化）
- GUI（Tkinter）+ CLI 双入口
- 异步后台转换 + 窗口关闭自动取消

---

## 二、整体架构

```
main.py                          ← 程序入口
    │
    ├── GUI 模式                   ← DocConvertApp (tkinter)
    │       └── ConversionController   (调度转换任务)
    │               ├── ExcelConverter  (read Excel)
    │               ├── WordConverter   (read Word)
    │               ├── DocConverter    (read legacy .doc)
    │               └── BaseConverter   (抽象基类)
    │
    └── CLI 模式                   ← main_cli (argparse)
            └── ConversionController
                    ├── converters/     ← 各格式读取 & 内容构建
                    ├── cleaners/       ← Word→MD 清洗管道
                    ├── exporters/      ← 输出格式适配器
                    ├── utils/          ← 工具函数
                    ├── models/         ← 数据模型
                    ├── config/         ← 配置
                    ├── gui/            ← Tkinter 界面
                    ├── parsers/        ← 预留扩展点（当前仅 ABC）
                    └── chunkers/       ← 预留扩展点（当前仅 ABC）
```

### 数据流

```
输入文件
  │
  ▼
┌─────────────────┐
│  Converter      │  读取源文件，生成原始内容
│  (Excel/Word/Doc)│
└────────┬────────┘
         │
         ▼
┌─────────────────┐     ┌─────────────────┐
│  Cleaner        │────▶│  Exporter       │  最终写入文件
│  (仅 Word→MD)   │     │  (HTML/MD/JSON) │
└─────────────────┘     └─────────────────┘
```

---

## 三、目录结构

```
github-codex/
├── main.py                      # 入口：GUI 或 CLI 模式选择
├── pyproject.toml               # 项目元数据 & 依赖声明
├── gen_doc.py                   # 自动生成文档脚本
├── build_scripts/
│   └── build_exe.py             # PyInstaller 打包脚本
├── docconvert/
│   ├── __init__.py              # 包初始化（懒加载 DocConvertApp）
│   ├── cli.py                   # CLI 入口：argparse 命令解析
│   ├── config.py                # AppConfig 数据类 & 默认配置
│   ├── logger.py                # 日志初始化（单例 LOGGER_NAME="docconvert"）
│   ├── controllers/
│   │   ├── __init__.py
│   │   └── conversion_controller.py  # 核心调度器：异步/取消/覆写检查
│   ├── converters/
│   │   ├── __init__.py
│   │   ├── base.py              # BaseConverter ABC
│   │   ├── excel.py             # ExcelConverter：xlsx/xls → HTML/MD/JSON
│   │   ├── word.py              # WordConverter：docx → HTML/MD/JSON
│   │   └── doc.py               # DocConverter：legacy .doc → HTML/MD/JSON
│   ├── cleaners/
│   │   ├── __init__.py
│   │   ├── base.py              # BaseCleaner ABC
│   │   └── word_md.py           # WordMdCleaner：4 级清洗规则
│   ├── exporters/
│   │   ├── __init__.py          # get_exporter() 工厂函数
│   │   ├── base.py              # BaseExporter ABC
│   │   ├── html.py              # HtmlExporter（透传）
│   │   ├── markdown.py          # MarkdownExporter（透传）
│   │   └── json_exporter.py     # JsonExporter（json.dumps）
│   ├── models/
│   │   ├── __init__.py
│   │   └── models.py            # MergeInfo, ProgressEvent
│   ├── utils/
│   │   ├── __init__.py
│   │   └── utils.py             # 工具函数集
│   ├── gui/
│   │   ├── __init__.py
│   │   └── app.py               # DocConvertApp（Tkinter GUI）
│   ├── parsers/                 # 预留扩展点（仅 BaseParser ABC）
│   └── chunkers/                # 预留扩展点（仅 BaseChunker ABC）
├── tests/                       # unittest 测试套件
│   ├── test_controller.py
│   ├── test_cleaners.py
│   ├── test_excel_converter.py
│   ├── test_exporters.py
│   ├── test_logger.py
│   ├── test_models.py
│   ├── test_utils.py
│   └── test_word_doc_sheets.py
├── .github/workflows/
│   ├── ci.yml                   # CI 流水线
│   └── release.yml              # PyInstaller 发布流水线
└── README.md / CONTRIBUTING.md / LICENSE
```

---

## 四、关键类与函数详解

### 4.1 入口模块

#### `main.py`

程序唯一入口，根据命令行参数决定运行模式。

```python
def main():
    setup_logging()
    if len(sys.argv) > 1 and sys.argv[1] == 'convert':
        from docconvert.cli import main_cli
        sys.exit(main_cli(sys.argv[1:]))
    else:
        # 启动 GUI
        import tkinter as tk
        from docconvert.gui import DocConvertApp
        root = tk.Tk()
        DocConvertApp(root)
        root.mainloop()
```

> **设计要点**：GUI 模块在 `docconvert/__init__.py` 中以 `__getattr__` 懒加载，使 CLI / 库调用不强制依赖 Tkinter。

#### `docconvert/cli.py` — `main_cli(argv)`

CLI 子命令入口，负责参数解析与批量转换循环。

| 参数 | 说明 |
|------|------|
| `files` | 一个或多个输入文件路径 |
| `--format / -f` | `html`（默认）、`md`、`json` |
| `--output / -o` | 输出目录（默认：首文件所在目录） |
| `--enhanced / -e` | 启用增强 Markdown 输出 |
| `--sheet / -s` | 指定 Excel 工作表（可重复） |
| `--verbose / -v` | 开启 DEBUG 级别日志 |

返回码：`0` 全部成功，`1` 有失败项。

---

### 4.2 配置与日志

#### `docconvert/config.py` — `AppConfig`

```python
@dataclass
class AppConfig:
    chunk_size: int = 1000
    max_rows: int = 0
    markdown_style: str = "github"
    preview_chars: int = 5000
    preview_lines: int = 150
    large_file_size: int = 20 * 1024 * 1024  # 20MB

    cleaning_rules: dict[str, bool] = field(default_factory=lambda: {
        "remove_page_numbers": True,
        "remove_duplicate_headers": True,
        "remove_empty_lines": True,
        "normalize_spaces": True,
    })
```

全局单例 `DEFAULT_CONFIG = AppConfig()` 被 Controller 和 GUI 共享使用。

#### `docconvert/logger.py`

```python
LOGGER_NAME = "docconvert"

def setup_logging(level: Union[int, str] = logging.INFO) -> logging.Logger
def get_logger() -> logging.Logger
```

使用 `StreamHandler` 输出到 stdout，格式化字符串：
```
[HH:MM:SS] LEVEL    docconvert - message
```

> **注意**：`setup_logging` 第二次调用时会同步更新已有 handler 的 level，避免日志被旧的 level 过滤掉。

---

### 4.3 数据模型

#### `docconvert/models/models.py`

```python
@dataclass
class MergeInfo:
    """描述 Excel 合并单元格信息"""
    rowspan: int = 1
    colspan: int = 1
    is_master: bool = False   # 合并区域左上角（真正持有数据的单元格）
    is_merged: bool = False
    min_row: int = 0
    min_col: int = 0
    max_row: int = 0
    max_col: int = 0

@dataclass
class ProgressEvent:
    """GUI/CLI 进度回调事件"""
    message: str = ""
    progress: float = 0.0     # 0.0 ~ 1.0
    done: bool = False
    error: Optional[str] = None
```

---

### 4.4 转换器层（Converters）

#### `docconvert/converters/base.py` — `BaseConverter`

所有转换器的抽象基类：

```python
class BaseConverter(ABC):
    def __init__(self, config: Optional[AppConfig] = None)
    @abstractmethod
    def convert(self, input_path: str, output_fmt: str, parent: Path, **kwargs) \
        -> tuple[list[tuple[str, str]], list[tuple[str, str]]]
        # 返回 (成功列表, 错误列表)，每项为 (output_name, output_path)

    def _export(self, content: object, output_fmt: str) -> str
        # 委托给 get_exporter(output_fmt).export(content)
```

每个子类实现 `convert()`，负责读取文件、构建内容、写入磁盘。

---

#### `docconvert/converters/excel.py` — `ExcelConverter`

**负责**：`.xlsx` 和 `.xls` 文件的读取与转换。

##### 核心方法

| 方法 | 职责 |
|------|------|
| `convert()` | 主入口：加载所有 sheet，遍历转换，处理取消信号 |
| `_load_merged_cache()` | 加载合并单元格信息（xlsx → openpyxl，xls → xlrd） |
| `_load_merged_cache_xls()` | 专门处理 `.xls` 格式的合并单元格 |
| `_build_merged_map()` | 将合并范围展开为 `(row, col) → MergeInfo` 映射 |
| `_df_to_rows()` | DataFrame → 二维字符串列表（含表头） |
| `_filter_merges_for_dropped_rows()` | dropna 后过滤越界合并范围 |
| `_remap_merged_rows()` | 将 workbook 行号重映射为 dropna 后的渲染行号 |
| `_convert_sheet()` | 单个 sheet 的完整转换逻辑 |
| `_write_html()` | 构建带合并单元格支持的 HTML 表格 |
| `_generate_md_via_html()` | 增强 MD：通过 HTML → markdownify 路径 |
| `_generate_md_standard()` | 标准 MD：pandas `to_html()` + `html_to_md()` |
| `_write_json()` | 生成包含 metadata + data + merged_cells 的 JSON |
| `_cell_attrs()` | 生成 `data-row/data-col/rowspan/colspan` 属性字符串 |

##### Excel HTML 表格的关键处理

- `<thead>` 中支持 `colspan`，但 `rowspan` 不能跨越 thead/tbody 边界，因此**垂直合并的表头会被展平**（仅保留 colspan）。
- 被表头垂直合并覆盖的 tbody 单元格输出为 `&nbsp;` 占位符以保持列对齐。
- JSON 输出中非主合并单元格（slave leg）会记录 `merged=True, skipped=True`，值引用 master cell。

##### 合并单元格空行处理（高级）

`dropna(how='all')` 会删除全空行，导致 workbook 行号与实际渲染行号不一致。解决方案分两步：
1. `_filter_merges_for_dropped_rows()` — 删除跨被删行的合并范围
2. `_remap_merged_rows()` — 将剩余合并范围的 workbook 坐标重映射为渲染坐标

---

#### `docconvert/converters/word.py` — `WordConverter`

**负责**：`.docx` 文件的读取与转换。

##### 核心方法

| 方法 | 职责 |
|------|------|
| `convert()` | 主入口：调 `_convert_word()`，异常捕获后返回 |
| `_convert_word()` | 加载 DOCX → 可选清洗 → 按需转换为目标格式 |
| `_strip_headers_footers()` | 从 DOCX XML 中删除 `headerReference`/`footerReference` |
| `_to_html()` | mammoth → 包装为完整 HTML 文档 |
| `_to_md()` | mammoth（HTML 或 Markdown 路径）→ WordMdCleaner 清洗 |
| `_to_json()` | mammoth HTML → 结构化 JSON（metadata + content） |

##### Word→MD 流程（enhanced mode）

```
docx
  → mammoth.convert_to_html()
  → html_to_md()          (BeautifulSoup + markdownify)
  → WordMdCleaner.clean() (4 级 regex 规则)
  → 写入 .md 文件
```

---

#### `docconvert/converters/doc.py` — `DocConverter`

**负责**：`.doc`（旧版 Word 97-2003 格式）转换，依赖 `textract`。

> 仅在 Linux/macOS 上可用（Windows 不支持 textract）。

核心方法 `_convert_doc()`：
- `textract.process(input_path)` 提取纯文本
- `decode_text()` 自动检测编码（UTF-8 → GB18030 → Latin-1）
- 按 output_fmt 包装为对应格式

---

### 4.5 清洗器层（Cleaners）

#### `docconvert/cleaners/base.py` — `BaseCleaner`

```python
class BaseCleaner(ABC):
    @abstractmethod
    def clean(self, content: str, **kwargs) -> str
```

#### `docconvert/cleaners/word_md.py` — `WordMdCleaner`

四级清洗管道，**只在 Word→MD 增强模式下生效**：

| 阶段 | 规则 | 正则表达式 |
|------|------|-----------|
| 1 | 移除页码行 | `^\[\d+\]$`、`^第\s*\d+\s*页$`、`^-\s*\d+\s*-$`、`^page\s*\d+\s*(?:of\|/)\s*\d+\s*$`、`^(?:Page\|Pág\.\|P\.)\s*\d+\s*$` |
| 2 | 折叠连续空行 | 保留首个空行，后续跳过 |
| 3 | 规范化空格 | U+3000 → 普通空格，tab → 空格，折叠连续水平空白 |
| 4 | 移除连续重复行 | 仅对相邻非空行去重 |

**重要**：所有规则均跳过代码块（fenced ` ``` ` / `~~~` 和缩进 ≥4 空格）。

---

### 4.6 导出器层（Exporters）

#### `docconvert/exporters/__init__.py` — `get_exporter()`

```python
def get_exporter(fmt: str) -> BaseExporter:
    if fmt == 'html':    return HtmlExporter()
    if fmt == 'md':      return MarkdownExporter()
    if fmt == 'json':    return JsonExporter()
    raise ValueError(...)
```

| 类 | `export()` 行为 |
|----|----------------|
| `HtmlExporter` | 直接返回 str；dict 取 `.get('content')` |
| `MarkdownExporter` | 同上 |
| `JsonExporter` | 直接返回 str；dict 用 `json.dumps(..., ensure_ascii=False, indent=2)` |

---

### 4.7 控制器（Controller）

#### `docconvert/controller/conversion_controller.py` — `ConversionController`

整个系统的**核心调度器**，协调多文件批量转换、取消、进度推送。

##### 关键属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `config` | `AppConfig` | 当前转换配置 |
| `_cancel_event` | `threading.Event` | 取消信号 |
| `_done_event` | `threading.Event` | 完成信号 |
| `_thread` | `Thread` | 后台工作线程 |
| `_running` | `bool` | 是否正在转换 |
| `last_results` | `list[(name, path, err)]` | 最后一次转换结果 |
| `was_cancelled` | `bool` | 是否被用户取消 |
| `last_error` | `Optional[str]` | 最后一次异常信息 |

##### 关键方法

| 方法 | 签名 | 说明 |
|------|------|------|
| `convert_files()` | `(files, output_fmt, output_dir, enhanced_md, sheets, progress_callback)` | **同步**批量转换，逐文件处理 |
| `convert_files_async()` | 同上 | **异步**启动后台线程 |
| `cancel()` | `()` | 设置取消事件 |
| `reset_cancel()` | `()` | 清除取消事件 |
| `wait_for_completion(timeout)` | `Optional[float]` | 阻塞等待线程结束 |
| `wait_done(timeout)` | `Optional[float]` | 基于 Event 等待 |
| `check_overwrite_paths()` | `(files, output_fmt, output_dir, sheets)` | 预计算输出路径并返回已存在的路径列表 |
| `set_config(config)` | `AppConfig` | 更新配置（转换期间忽略） |

##### 批量输出文件名去重策略

`_make_unique_stem(base_stem, source_path, used)` 处理同名文件冲突：
```
"report"              → "report"
"report" (再次出现)    → "report_parentDir"
"report" (第三次)      → "report_parentDir_2"
...
```
所有输出文件写入同一目录（由 `_batch_parent()` 决定：显式 output_dir 或首文件的父目录）。

---

### 4.8 工具函数

#### `docconvert/utils/utils.py`

| 函数 | 说明 |
|------|------|
| `get_excel_sheet_names(filepath, ext)` | 读取 Excel 工作表名（xlrd / openpyxl） |
| `safe_str(value)` | 安全地将任意值转为字符串（处理 NaN、None、换行符规范化） |
| `clean_filename(name)` | 清理文件系统非法字符，限制长度 ≤180，防止 Windows 保留名（CON、PRN 等） |
| `unique_cleaned_suffixes(names)` | 对 sheet 名进行 clean_filename 后确保唯一性（追加 `_2`/`_3` 后缀） |
| `decode_text(raw)` | 自动尝试 UTF-8 → GB18030 → Latin-1 解码 |
| `escape_md_cell(text)` | 转义 Markdown 表格中的 `\|` 和换行符（换行 → `<br>`） |
| `html_to_md(html_content)` | BeautifulSoup 移除 script/style → markdownify 转 Markdown |

---

### 4.9 GUI 模块

#### `docconvert/gui/app.py` — `DocConvertApp`

基于 Tkinter 的现代化 GUI，主要组件：

```
┌─────────────────────────────────────────────────┐
│  Title Bar: "DocConvert   文档转换工具  v2.0"    │
├─────────────────────────────────────────────────┤
│  [文件选择卡]                                     │
│    输入文件: [____________] [浏览]               │
│    工作表:   [___________]                        │
│    输出目录: [____________] [选择][默认]          │
│    ☑ 转换所有工作表                              │
├─────────────────────────────────────────────────┤
│  [文件列表卡]                                     │
│    [添加文件] [移除选中] [清空]                   │
│    ┌──────────────────────────────────┐         │
│    │ file1.xlsx                        │         │
│    │ file2.docx                        │         │
│    └──────────────────────────────────┘         │
├─────────────────────────────────────────────────┤
│  [输出格式卡]                                     │
│    ○ HTML  ○ Markdown  ○ JSON                   │
│    ☐ 增强 Markdown 输出                          │
│    [Markdown 清洗 (Word→MD)]                      │
│      ☐ 移除页码  ☐ 移除重复页眉                  │
│      ☐ 移除多余空行  ☐ 合并多余空白              │
├─────────────────────────────────────────────────┤
│  [预览卡]                                         │
│    ┌──────────────────────────────────┐         │
│    │ 文件内容预览...                   │         │
│    └──────────────────────────────────┘         │
├─────────────────────────────────────────────────┤
│  状态: 就绪                           [进度条] [开始转换] │
└─────────────────────────────────────────────────┘
```

**线程安全设计**：
- 后台线程通过 `queue.Queue` 将回调分发到主线程（Tkinter 非线程安全）
- `_poll_pending()` 每 50ms 通过 `root.after()` 在 UI 线程中消费队列
- 关闭窗口时触发 `_on_close()` → 设置取消事件 → 等待线程结束后销毁

**覆写确认逻辑**：转换前调用 `check_overwrite_paths()` 预计算所有输出路径，若存在则弹出确认对话框。

---

### 4.10 预留扩展点

| 模块 | 基类 | 当前实现 | 用途 |
|------|------|----------|------|
| `parsers/` | `BaseParser` | 无具体实现 | 语义解析（预留） |
| `chunkers/` | `BaseChunker` | 无具体实现 | 内容分块（预留） |

---

## 五、依赖关系图

```
                         ┌─────────────────────┐
                         │  main.py / cli.py   │
                         └──────────┬──────────┘
                                    │
                         ┌──────────▼──────────┐
                         │  ConversionController│
                         └────┬─────────┬──────┘
                              │         │
              ┌───────────────┘         └───────────────┐
              ▼                                         ▼
    ┌─────────────────┐                       ┌─────────────────┐
    │  ExcelConverter │                       │  WordConverter  │
    │  DocConverter   │                       │  (optional)     │
    └───────┬─────────┘                       └────────┬────────┘
            │                                          │
    ┌───────▼─────────┐                       ┌───────▼─────────┐
    │  openpyxl       │                       │  python-docx    │
    │  pandas         │                       │  mammoth        │
    │  xlrd (optional)│                       └───────┬─────────┘
    └───────┬─────────┘                               │
            │                                  ┌─────▼──────┐
    ┌───────▼─────────┐                       │  WordMdCleaner│
    │  exporters/     │                       └──────────────┘
    │  (html/md/json) │
    └───────┬─────────┘
            │
    ┌───────▼─────────┐
    │  bs4 (optional) │
    │  markdownify    │
    │  (MD exporter)  │
    └─────────────────┘
```

### 依赖清单（来自 pyproject.toml）

**核心依赖**：
- `python-docx>=1.1.0` — DOCX 读写
- `openpyxl>=3.1.0` — XLSX 读写
- `pandas>=2.0.0` — Excel 数据读取
- `mammoth>=1.6.0` — DOCX → HTML/Markdown
- `beautifulsoup4>=4.12.0` — HTML 解析（MD 转换辅助）
- `markdownify>=0.11.0` — HTML → Markdown
- `xlrd>=2.0.0` — XLS 格式读取

**可选依赖**：
- `[test]` — `mypy`, `xlwt`
- `[doc]` — `textract>=1.6.0`（仅 Linux/macOS，用于 `.doc` 支持）
- `[build]` — `pyinstaller>=6.0`

---

## 六、项目运行方式

### 6.1 安装

```bash
# 基础安装
pip install docconvert-local

# 支持 .doc 格式（Linux/macOS）
pip install docconvert-local[doc]

# 全部功能（含测试、文档、打包）
pip install docconvert-local[all]

# 开发模式安装
pip install -e ".[all]"
```

### 6.2 运行模式

#### GUI 模式（推荐普通用户）

```bash
python main.py
```

#### CLI 模式（脚本/批量）

```bash
# 单文件转换
python main.py convert input.xlsx --format md

# 多文件批量转换
python main.py convert file1.xlsx file2.docx --format html -o ./output

# 增强 Markdown（Word 清洗管道）
python main.py convert input.docx --format md --enhanced

# 指定工作表
python main.py convert input.xlsx --format json --sheet "Sheet1" --sheet "Sheet2"

# 详细日志
python main.py convert input.xlsx --format html --verbose
```

#### 编程调用（作为库）

```python
from docconvert.controller import ConversionController

controller = ConversionController()
results = controller.convert_files(
    files=["input.xlsx"],
    output_fmt="html",
    enhanced_md=True,
)
for name, path, error in results:
    if error:
        print(f"失败: {name} - {error}")
    else:
        print(f"成功: {name} -> {path}")
```

#### 后台异步调用

```python
controller.convert_files_async(
    files=["file1.xlsx", "file2.docx"],
    output_fmt="md",
    progress_callback=my_callback,
)
controller.wait_for_completion(timeout=60)
print(controller.last_results)
print(controller.was_cancelled)
```

### 6.3 测试

```bash
pip install -e ".[all]"
python -m unittest discover -s tests -v
```

测试文件与源代码模块一一对应：

| 测试文件 | 测试目标 |
|----------|----------|
| `test_controller.py` | 批量转换、取消、覆写检查 |
| `test_cleaners.py` | 页码移除、空行折叠、去重、空格归一化 |
| `test_excel_converter.py` | 合并单元格处理、多 sheet 转换 |
| `test_exporters.py` | 三种导出器的数据路由 |
| `test_logger.py` | 日志级别更新行为 |
| `test_models.py` | MergeInfo / ProgressEvent |
| `test_utils.py` | clean_filename、escape_md_cell、html_to_md |
| `test_word_doc_sheets.py` | Word 段落/表格读取、Doc 文件处理 |

### 6.4 打包为可执行文件

```bash
pip install -e ".[all,build]"
python build_scripts/build_exe.py --clean
```

输出路径：`dist/DocConvert-<platform>-<arch>/`

---

## 七、关键设计模式

### 7.1 策略模式（Converter）

`BaseConverter` 定义了统一的 `convert()` 接口，三种具体转换器（`ExcelConverter`、`WordConverter`、`DocConverter`）分别实现不同文件格式的读取逻辑，Controller 通过文件扩展名动态选择对应策略。

### 7.2 模板方法模式（Exporter）

`BaseExporter.export()` 是模板方法，子类决定具体序列化方式（字符串透传 / JSON 序列化）。

### 7.3 工厂模式

`get_exporter(fmt)` 工厂函数根据格式字符串返回对应的 `BaseExporter` 实例。

### 7.4 责任链（Cleaner）

`WordMdCleaner` 实现四级清洗管道，每级规则独立可配置，按顺序依次执行（页码 → 空行 → 空格 → 去重），代码块（fenced/缩进）全程豁免。

### 7.5 生产者-消费者（GUI 线程安全）

后台线程通过 `queue.Queue` 提交 UI 回调，主线程每 50ms 通过 `root.after()` 轮询消费，确保所有 Tkinter 操作都在主线程执行。

---

## 八、扩展指南

### 添加新的输出格式

1. 在 `docconvert/exporters/` 下新建 `myformat.py`，实现 `BaseExporter`：
```python
class MyFormatExporter(BaseExporter):
    def export(self, data: Any, **kwargs) -> str:
        ...
```
2. 在 `docconvert/exporters/__init__.py` 的 `get_exporter()` 中添加分支。

### 添加新的输入格式

1. 在 `docconvert/converters/` 下新建 `myformat.py`，实现 `BaseConverter.convert()`。
2. 在 `docconvert/converters/__init__.py` 中导出新类。
3. 在 `ConversionController._get_converter()` 中添加扩展名分支。

### 修改清洗规则

`AppConfig.cleaning_rules` 字典控制各规则开关，直接修改后传入新 `AppConfig` 即可，无需修改 `WordMdCleaner` 代码。
