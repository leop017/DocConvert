# `docconvert.utils` — 通用工具

> 源码: [`docconvert/utils/utils.py`](../../docconvert/utils/utils.py)
> `__all__`: `["INVALID_NAMES", "safe_str", "clean_filename", "escape_md_cell", "html_to_md", "get_excel_sheet_names", "decode_text", "unique_cleaned_suffixes"]`

一组**纯函数**工具,被转换器、控制器、CLI 共用。

---

## `INVALID_NAMES`

```python
INVALID_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
```

Windows 保留的设备名(大写)。`clean_filename` 在生成 stem 时,若清洗后的 stem 命中此集合,会**前缀下划线**防止在 Windows 上写文件失败。

---

## `get_excel_sheet_names(filepath, ext)`

```python
def get_excel_sheet_names(filepath: str, ext: str) -> list[str]: ...
```

读取工作簿中**所有 sheet 名**。

| 参数 | 说明 |
|------|------|
| `filepath` | 工作簿绝对 / 相对路径 |
| `ext` | 已是小写的扩展名(`.xls` / `.xlsx`);`.xls` 走 `xlrd`,`.xlsx` 走 `openpyxl`(`read_only=True`) |

**返回**:`list[str]`,顺序与工作簿中 sheet 的物理顺序一致。

**抛出**:文件不存在 / 损坏时由底层库(`xlrd` / `openpyxl`)抛出。

```python
from docconvert.utils import get_excel_sheet_names
get_excel_sheet_names('finance.xlsx', '.xlsx')  # ['Q1', 'Q2', '汇总']
```

---

## `safe_str(value)`

```python
def safe_str(value: Any) -> str: ...
```

把任意值规范化为可写入 Markdown 单元格的字符串。

| 输入 | 输出 |
|------|------|
| `None` | `""` |
| `float('nan')` / `pandas.NA` / `numpy.nan` | `""` |
| `"a\r\nb"` | `"a\nb"`(统一换行) |
| 其他 | `str(value)`(再统一换行) |

> 通过 `try/except ImportError` 守卫 `pandas.isna`,在没有装 pandas 的环境下也能工作(仅 `None` 和 `float('nan')` 被识别为空)。

---

## `clean_filename(name)`

```python
def clean_filename(name: str) -> str: ...
```

把任意字符串清洗为**安全的文件名 stem**:

1. 去除控制字符(`\x00 – \x1f`)
2. 把 `/\:?"<>|` 替换为 `_`
3. 空白 / 全空 → `"untitled"`
4. **Windows 保留名**(`CON` / `PRN` / ...):前缀 `_`
5. 长度 > 180:按 UTF-8 截断(避免某些文件系统 255 字节限制)

```python
clean_filename('Q"1')              # 'Q_1'
clean_filename('报告/2026')         # '报告_2026'
clean_filename('CON')              # '_CON'
clean_filename('')                 # 'untitled'
clean_filename('中文' * 200)        # 截断到 180 字节
```

---

## `unique_cleaned_suffixes(names)`

```python
def unique_cleaned_suffixes(names: list[str]) -> list[str]: ...
```

`clean_filename` 之后,再做**批内去重**。撞名时附加 `_2` / `_3`...

```python
unique_cleaned_suffixes(['Q"1', 'Q<1', 'Q/1'])
# ['Q_1', 'Q_1_2', 'Q_1_3']
```

被 `ExcelConverter` 和 `ConversionController._compute_output_paths` 共用,保证"算出的路径"与"实际写入的路径"完全一致。

---

## `decode_text(raw)`

```python
def decode_text(raw: bytes) -> str: ...
```

按 UTF-8 → GB18030 → Latin-1 顺序回退解码字节流。**永不抛 `UnicodeDecodeError`**(Latin-1 一定成功)。

| 输入 | 期望输出 |
|------|---------|
| UTF-8 字节 | 原 UTF-8 字符串 |
| GBK / GB18030 字节(常见中文 .doc) | 正确中文 |
| 含 `\ufffd`(U+FFFD 替换符) | 自动尝试下一编码 |
| 其他字节 | Latin-1 兜底 |

> 设计动机:`textract` / `antiword` 的输出编码依赖系统 locale,在 zh-CN Windows 上常是 GBK,在 macOS / Linux 上常是 UTF-8。

---

## `escape_md_cell(text)`

```python
def escape_md_cell(text: str) -> str: ...
```

把 Markdown 表格**单元格内**的 `|` 转义为 `\|`,换行转为 `<br>`。

- 仅用于 `<th>` / `<td>` 的内容,**不要**对 HTML 标签本身调用
- 表格行格式: `| col1 | col2 |`,若 `col1` 含 `|` 会破坏列数
- `<br>` 会被 markdownify 透传,渲染为真实换行

```python
escape_md_cell('a|b\nc')  # 'a\\|b<br>c'
```

---

## `html_to_md(html_content)`

```python
def html_to_md(html_content: str) -> str: ...
```

把 HTML 片段转换为 Markdown。

- 用 `BeautifulSoup` 解析(`html.parser`,无外部依赖)
- **先剥离** `<script>` / `<style>` 标签
- `markdownify(heading_style=ATX)` — 标题用 `#` 前缀
- 末尾 `strip()` 去尾随空白

```python
html_to_md('<h1>Title</h1><p>Body</p>')
# '# Title\n\nBody'
```

> 与 `DocConverter` / `WordConverter` 的"enhanced_md"路径共用。

---

## 设计要点

| 关注点 | 实现 |
|--------|------|
| **无 IO** | 所有函数纯计算;`get_excel_sheet_names` 是唯一例外(读工作簿) |
| **平台兼容** | `clean_filename` 显式处理 Windows 保留名 |
| **中文友好** | `decode_text` 显式尝试 GB18030(GBK 超集) |
| **批一致** | `clean_filename` → `unique_cleaned_suffixes` 流水线,确保"算路径"与"写路径"一致 |
