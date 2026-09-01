# `docconvert.converters` — 输入文件解析器

> 源码: [`docconvert/converters/`](../../docconvert/converters/)
> `__all__`: `["BaseConverter", "ExcelConverter", "WordConverter", "DocConverter"]`

每个具体转换器负责把一种**输入文件**解析为**中间表示**,再通过 `BaseExporter` 序列化为最终输出。绝大多数业务场景下,你不需要直接 new 这些类 — 让 `ConversionController` 帮你选择即可。

---

## `class BaseConverter`(ABC)

```python
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Callable, Optional

from docconvert.config import AppConfig, DEFAULT_CONFIG

class BaseConverter(ABC):
    def __init__(self, config: Optional[AppConfig] = None) -> None: ...
```

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `config` | `AppConfig` | 当前使用的配置 |
| `logger` | `logging.Logger` | 命名 `docconvert` 的 logger |
| `cancel_check` | `Optional[Callable[[], bool]]` | 控制器注入的取消检查函数,返回 `True` 时应尽快退出 |

### 抽象方法

```python
@abstractmethod
def convert(
    self,
    input_path: str,
    output_fmt: str,
    parent: Path,
    **kwargs,
) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]: ...
```

| 参数 | 说明 |
|------|------|
| `input_path` | 输入文件绝对 / 相对路径 |
| `output_fmt` | `'html'` / `'md'` / `'json'` |
| `parent` | 输出目录 |
| `**kwargs` | 转换器私有参数(如 `enhanced_md` / `sheets` / `stem_override`) |

**返回**: `(results, errors)`

- `results: list[tuple[name, output_path]]` — 成功项
- `errors: list[tuple[name, error_msg]]` — 失败项

### 受保护方法

```python
def _export(self, content: object, output_fmt: str) -> str: ...
```

`content` 交给 `get_exporter(output_fmt)` 序列化为最终字符串。子类直接 `self._export(content, fmt)` 即可。

---

## `class ExcelConverter`

> 源码: [`docconvert/converters/excel.py`](../../docconvert/converters/excel.py)

支持 `.xlsx`(openpyxl 引擎)和 `.xls`(xlrd 引擎)。

### 构造

```python
def __init__(self, config: Optional[AppConfig] = None) -> None: ...
```

无额外属性。`config.max_rows` 限制每个 sheet 读入的最大行数(0 = 不限制)。

### `convert` 关键字参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `sheets` | `Optional[list[str]]` | 仅转换列出的 sheet;`None` = 全部。**不存在的 sheet 名**会被记录为 `('sheet', '工作表不存在')` 错误并跳过。 |
| `enhanced_md` | `bool` | 当 `output_fmt='md'` 时,走 HTML 中转路径(mammoth-style 不适用 Excel,这里实际为 HTML→markdownify 路径)。 |
| `stem_override` | `Optional[str]` | 强制覆盖默认 stem,用于批内 basename 去重。 |

### 关键能力

1. **合并单元格处理**
   - 行/列 `rowspan` / `colspan` 正确渲染为 HTML 的 `rowspan` / `colspan` 属性
   - `dropna` 删除空行后,合并范围会被**重映射**(`_remap_merged_rows`)、并**过滤掉跨越被删行的合并**(`_filter_merges_for_dropped_rows`)
   - **表头锚定的垂直合并**会被"展平"为每行重复值(HTML 表格规范中,`<thead>` 与 `<tbody>` 之间的 `rowspan` 行为不一致)
2. **数值保真**
   - `0` / `False` 不会被替换为 `&nbsp;`
   - `NaN` / `None` → 空字符串
   - 浮点 `NaN` 通过 `safe_str` 判空
3. **Markdown 转义**
   - 单元格内的 `|` 转为 `\|`(避免破坏表格布局)
   - 单元格内的换行转为 `<br>`(经 markdownify 透传)
4. **JSON 输出**包含 `metadata` / `headers` / `rows` / `merged_cells` 四段,便于下游重建表格。

### 输出文件名

- 多 sheet: `<stem>_<sheet>.html|md|json`
- 单文件多 sheet 撞名(如 `Q"1` 和 `Q<1` 都清洗为 `Q_1`):自动加 `_2`、`_3` 后缀

### 示例

```python
from pathlib import Path
from docconvert.converters import ExcelConverter
from docconvert.config import DEFAULT_CONFIG

conv = ExcelConverter(DEFAULT_CONFIG)
results, errors = conv.convert(
    input_path='finance.xlsx',
    output_fmt='md',
    parent=Path('./out'),
    sheets=['Q1', 'Q2'],
)
for name, path in results:
    print('OK', name, '->', path)
for name, err in errors:
    print('FAIL', name, err)
```

### 异常

| 场景 | 行为 |
|------|------|
| `.xls` 但未装 `xlrd` | 抛 `RuntimeError('需安装 xlrd 库以处理 .xls 文件')` |
| 工作簿无法打开 | pandas 抛 `ValueError` / `openpyxl` 抛 `InvalidFileException` |
| 所有请求的 sheet 都不存在 | 直接返回 `([], [...错误])` |

---

## `class WordConverter`

> 源码: [`docconvert/converters/word.py`](../../docconvert/converters/word.py)

支持 `.docx`(基于 `python-docx` + `mammoth`)。**注意:这是唯一会构造 `WordMdCleaner` 的转换器**。

### 构造

```python
def __init__(self, config: Optional[AppConfig] = None) -> None:
    super().__init__(config)
    self.md_cleaner = WordMdCleaner(self.config)
```

`self.md_cleaner` 在 `output_fmt='md'` 时自动使用。

### `convert` 关键字参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `enhanced_md` | `bool` | 走 `mammoth → HTML → markdownify → cleaner`(更稳)或 `mammoth → markdown → cleaner`(更快,质量略差) |
| `stem_override` | `Optional[str]` | 强制覆盖 stem |
| `sheets` | — | **显式抛** `TypeError('WordConverter does not accept "sheets" parameter')` |

### 关键能力

1. **页眉页脚剥离**:在送入 mammoth 之前,通过修改 `sectPr` XML 把 `headerReference` / `footerReference` 直接删掉,避免 mammoth 把页眉页脚重复展平到正文里。
2. **HTML / Markdown / JSON 三种格式**
   - HTML:套用 `Microsoft YaHei` 字体 + 表格边框样式
   - MD:开头追加 `<!-- source: ... | paragraphs: N -->` 注释
   - JSON:返回 `{"metadata": {...}, "content": "<html string>"}`
3. **空文档防御**:`paragraph_count` 统计非空段落,用于 metadata 字段。

### 输出文件名

`<stem>_doc.<fmt>`(`doc` 后缀是给批内多文件区分用,即使只有一个 Word 文件也会带)。

---

## `class DocConverter`

> 源码: [`docconvert/converters/doc.py`](../../docconvert/converters/doc.py)

支持**旧版二进制 `.doc`**(基于 `textract`,**仅 Linux / macOS**;Windows 安装 `textract` 需要额外的 antiword,且 CI 中通过可选依赖隔离)。

### `convert` 关键字参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `enhanced_md` | `bool` | 把纯文本包到 `<pre>` 后再走 `html_to_md`(会有 HTML 实体转义) |
| `stem_override` | `Optional[str]` | 强制覆盖 stem |
| `sheets` | — | 显式抛 `TypeError` |

### 关键能力

1. **多编码容错**:`decode_text` 按 UTF-8 → GB18030 → Latin-1 顺序回退,任意 .doc 都能解出可读文本。
2. **HTML 包壳**:`<pre>` 标签 + 字体,保持等宽。

### 输出文件名

`<stem>.<fmt>`(没有 `_doc` 后缀)。

### 平台限制

| OS | 是否可用 |
|----|---------|
| Linux | ✅(需 `textract` + 底层 antiword) |
| macOS | ✅ |
| Windows | ⚠️ 需手动装 `antiword.exe` 并加入 PATH。CI 通过 `sys.platform != 'win32'` 守卫。 |

---

## 扩展指南:新增一种文件格式

```python
# docconvert/converters/csv.py
from pathlib import Path
from typing import Optional
from docconvert.config import AppConfig
from docconvert.converters.base import BaseConverter

class CsvConverter(BaseConverter):
    def convert(
        self,
        input_path: str,
        output_fmt: str,
        parent: Path,
        **kwargs,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        results, errors = [], []
        try:
            import csv
            with open(input_path, newline='', encoding='utf-8') as f:
                rows = list(csv.reader(f))
            content = self._to_md(rows) if output_fmt == 'md' else str(rows)
            out = parent / f'{Path(input_path).stem}.{output_fmt}'
            out.write_text(self._export(content, output_fmt), encoding='utf-8')
            results.append((out.name, str(out)))
        except Exception as e:
            errors.append((Path(input_path).name, str(e)))
        return results, errors

    @staticmethod
    def _to_md(rows: list[list[str]]) -> str:
        if not rows:
            return ''
        head, *body = rows
        lines = ['| ' + ' | '.join(head) + ' |',
                 '| ' + ' | '.join('---' for _ in head) + ' |']
        lines += ['| ' + ' | '.join(r) + ' |' for r in body]
        return '\n'.join(lines)
```

并在 `docconvert/converters/__init__.py` 与 `docconvert/controller/conversion_controller.py._get_converter` 中分别导出 / 分发。

---

## 设计要点

| 关注点 | 处理方式 |
|--------|---------|
| **取消语义** | 每次循环开头检查 `self.cancel_check()`,返回 `True` 时立即 break |
| **错误隔离** | 单个 sheet / 文件失败不阻断后续,通过 `errors` 列表上报 |
| **路径去重** | 完全由 `ConversionController` 处理,转换器只接受 `stem_override` |
| **格式互不耦合** | 转换器只关心"中间表示",序列化为 `str` 全部委托给 `BaseExporter` |
| **平台差异** | `.doc` 通过 `textract` 平台守卫;`.xls` 通过按需 `import xlrd` |
