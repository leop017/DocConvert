# `docconvert.config` — 全局配置

> 源码: [`docconvert/config.py`](../../docconvert/config.py)
> 公开符号: `AppConfig`, `DEFAULT_CONFIG`

DocConvert 的所有可调参数集中在 `AppConfig`,通过依赖注入传给转换器、清洗器、导出器。**没有全局可变状态**。

---

## `class AppConfig`

```python
@dataclass
class AppConfig:
    chunk_size: int = 1000
    max_rows: int = 0
    markdown_style: str = "github"
    preview_chars: int = 5000
    preview_lines: int = 150
    large_file_size: int = 20 * 1024 * 1024

    cleaning_rules: dict[str, bool] = field(default_factory=lambda: {
        "remove_page_numbers": True,
        "remove_duplicate_headers": True,
        "remove_empty_lines": True,
        "normalize_spaces": True,
    })
```

### 字段

| 字段 | 类型 | 默认 | 用途 |
|------|------|------|------|
| `chunk_size` | `int` | `1000` | 预留:分块大小(目前未在转换器中实现,用于未来 RAG chunking) |
| `max_rows` | `int` | `0` | Excel 单 sheet 最大读入行数;`0` = 不限制 |
| `markdown_style` | `str` | `"github"` | 预留:Markdown 风格(目前仅 `github` 风格) |
| `preview_chars` | `int` | `5000` | GUI 预览窗口最多显示的字符数 |
| `preview_lines` | `int` | `150` | GUI 预览窗口最多显示的行数 |
| `large_file_size` | `int` | `20 MB` | GUI 中"大文件"阈值(> 此值给更激进提示) |
| `cleaning_rules` | `dict[str, bool]` | 4 条全开 | 控制 `WordMdCleaner` 各阶段是否启用 |

### `cleaning_rules` 详解

| Key | 控制的清洗阶段 | 详见 |
|-----|---------------|------|
| `remove_page_numbers` | 5 种页码正则删除 | [cleaners.md](cleaners.md) |
| `remove_duplicate_headers` | 连续重复行删除 | [cleaners.md](cleaners.md) |
| `remove_empty_lines` | 连续空行折叠 | [cleaners.md](cleaners.md) |
| `normalize_spaces` | 行内空白归一化 | [cleaners.md](cleaners.md) |

不识别的 key 会被 `WordMdCleaner._resolve_line_rules` 静默忽略。

### 修改默认值

```python
from docconvert.config import AppConfig
from docconvert.cleaners import WordMdCleaner
from docconvert.converters import WordConverter

# 只关闭全角空格归一化
cfg = AppConfig(cleaning_rules={
    'remove_page_numbers': True,
    'remove_duplicate_headers': True,
    'remove_empty_lines': True,
    'normalize_spaces': False,  # 唯一改动
})

cleaner = WordMdCleaner(cfg)
converter = WordConverter(cfg)
```

### `field(default_factory=...)` 的意义

`cleaning_rules` 使用 `default_factory` 而不是 `default`,确保**每个 `AppConfig` 实例**有独立 dict,避免"实例 A 修改 → 实例 B 被影响"的共享可变状态陷阱。

---

## `DEFAULT_CONFIG` — 单例

```python
DEFAULT_CONFIG = AppConfig()
```

- 模块导入时实例化一次
- 所有"未传 `config` 参数"的组件(转换器、清洗器、导出器、控制器)**默认**使用它
- 业务代码**不应** `DEFAULT_CONFIG.cleaning_rules['normalize_spaces'] = False` 这样直接 mutate(会污染其他使用者)。需要修改时,显式 `AppConfig(cleaning_rules=...)` 创建新实例。

---

## 与 `ConversionController.set_config` 配合

```python
from docconvert.controller import ConversionController
from docconvert.config import AppConfig

ctrl = ConversionController()  # 使用 DEFAULT_CONFIG
ctrl.convert_files_async(['a.docx'], 'md')  # 默认清洗规则
ctrl.wait_done()

# 新建一个 config
strict_cfg = AppConfig(cleaning_rules={
    'remove_page_numbers': True,
    'remove_duplicate_headers': False,
    'remove_empty_lines': True,
    'normalize_spaces': True,
})
ctrl.set_config(strict_cfg)  # 必须在 is_running=False 时
ctrl.convert_files_async(['b.docx'], 'md')
ctrl.wait_done()
```

> 详见 [controller.md](controller.md) 关于 `set_config` 的并发约束。
