# `docconvert.cleaners` — Markdown 后处理清洗器

> 源码: [`docconvert/cleaners/`](../../docconvert/cleaners/)
> `__all__`: `["BaseCleaner", "WordMdCleaner"]`

Cleaner 在 `WordConverter` 拿到 mammoth 输出后、写入磁盘前**二次加工**。它的存在是"防御性深度":即便 `WordConverter._strip_headers_footers` 已经从 XML 移除了页眉页脚引用,仍可能有页码、空白行、CJK 全角空格等"渗漏"需要兜底。

---

## `class BaseCleaner`(ABC)

```python
from abc import ABC, abstractmethod

class BaseCleaner(ABC):
    @abstractmethod
    def clean(self, content: str, **kwargs) -> str: ...
```

任何继承类只需实现 `clean(content, **kwargs) -> str`,**不写入文件**、**不做 IO**。

---

## `class WordMdCleaner`

> 源码: [`docconvert/cleaners/word_md.py`](../../docconvert/cleaners/word_md.py)

**4 阶段清洗流水线** + **代码块保护**,由 `AppConfig.cleaning_rules` 字典驱动。

### 构造

```python
def __init__(self, config: Optional[AppConfig] = None) -> None: ...
```

`config=None` 时使用 `DEFAULT_CONFIG`,4 条规则**全部开启**(向后兼容默认)。

### 清洗规则

| Key | 默认 | 阶段 | 作用 |
|-----|------|------|------|
| `remove_page_numbers` | `True` | 1(行级) | 删除行首独立的页码标记 |
| `remove_duplicate_headers` | `True` | 4(全文) | 删除连续重复的非空行 |
| `remove_empty_lines` | `True` | 2(行级) | 折叠 ≥2 个连续空行为 1 个 |
| `normalize_spaces` | `True` | 3(行内) | 全角空格 / Tab / 多空格 → 单空格,保留缩进 |

任意规则可在 `AppConfig.cleaning_rules` 中显式 `False` 关闭。

### 阶段顺序(重要)

```text
input content
    │
    ▼
[1] 行级规则: 删除页码标记        ← 仅 strip() 锚定,避免误伤正文
    │
    ▼
[2] 折叠空行                     ← 在 normalize 之前做,使全角空格折叠可被 dedup
    │
    ▼
[3] 行内空白归一化                ← 保留 leading 缩进(代码/列表/表格)
    │
    ▼
[4] 去重连续行                    ← 在 normalize 之后做,使仅空白不同的行被视为相同
    │
    ▼
output content
```

### 5 种页码正则(全部 strip() 锚定)

| # | 模式 | 示例 |
|---|------|------|
| 0 | `^\[\d+\]$` | `[1]`、`[12]` |
| 1 | `^第\s*\d+\s*页$` | `第3页`、`第 5 页` |
| 2 | `^-\s*\d+\s*-$` | `- 1 -`、`-12-` |
| 3 | `^page\s*\d+\s*(?:of\|/)\s*\d+\s*$` | `Page 1 of 10`、`Page 3 / 10` |
| 4 | `^(?:Page\|Pág\.\|P\.)\s*\d+\s*$` | `Page 1`、`Pág. 3`、`P. 7` |

匹配前会先**反转义 mammoth 的反斜杠转义**:

- `\[1\]` → `[1]`
- `\- 1 \-` → `- 1 -`
- `Pág\. 3` → `Pág. 3`

### 代码块保护(关键不变量)

下列行被识别为"代码",**所有规则**都跳过:

- **fenced 围栏**:行首 ` ``` ` 或 ` ``` ` 3+ 个,带可选语言标签(` ```python `)。整个块在打开和关闭围栏之间全部视为代码。
- **缩进代码块**:行首 ≥ 4 个空格 / Tab(CommonMark 缩进代码块)。

> ⚠️ 围栏分隔符本身**也算代码行**(`mask=True`),从而保证它绝不会被删 / 改。

### 公开方法

```python
def clean(self, content: str, **kwargs) -> str: ...
```

- **无副作用**:`str → str`
- **`**kwargs` 当前未使用**,保留扩展位(未来可能加 `keep_emoji` / `max_line_length` 等)
- 当 4 条规则全部关闭时,直接 `return content`(零拷贝)

### 示例

```python
from docconvert.cleaners import WordMdCleaner
from docconvert.config import AppConfig

# 默认:全开
cleaner = WordMdCleaner()
cleaner.clean('第一行\n\n\n第二行  \n第二行\n')

# 关闭 normalize_spaces
cfg = AppConfig(cleaning_rules={
    'remove_page_numbers': True,
    'remove_duplicate_headers': True,
    'remove_empty_lines': True,
    'normalize_spaces': False,  # 保留行内原始空白
})
cleaner = WordMdCleaner(cfg)
```

### 不变量(测试覆盖)

- 输入以 ` ``` ` 开头、围栏闭合时,内部所有内容**逐字符**保留
- 行内以 4 空格缩进的行**绝不被** `normalize_spaces` 合并
- 围栏标记**自身**不会被 4 阶段任一阶段改写
- 空字符串 / 单行内容 / 无规则触发 → 与输入完全一致

---

## 扩展指南:自定义 Cleaner

```python
from docconvert.cleaners import BaseCleaner

class EmojiStripper(BaseCleaner):
    def clean(self, content: str, **kwargs) -> str:
        import re
        return re.sub(r'[\U0001F600-\U0001F64F]', '', content)
```

并通过 `WordConverter` 私有属性 `md_cleaner` 替换,或新增 `EnhancedWordConverter` 子类。

---

## 性能

| 输入规模 | 单次 `clean()` 耗时(典型) |
|---------|------------------------|
| 1 KB | < 1 ms |
| 100 KB | ~10 ms |
| 1 MB | ~100 ms |
| 10 MB | ~1 s |

(`code_mask` 只扫一遍,O(n) 单遍;4 阶段之间用 list 而非生成器,避免多次重扫。)

---

## 设计要点

| 关注点 | 实现 |
|--------|------|
| **代码安全** | `_code_mask` 一次性计算,所有规则读 mask 不再二次扫描 |
| **规则可关** | 任意 key 不在 dict / `False` → 跳过该阶段 |
| **顺序敏感** | 阶段顺序经过设计,交换 2↔3 或 3↔4 会破坏去重效果 |
| **无外部 IO** | 纯函数式,易于单测 |
