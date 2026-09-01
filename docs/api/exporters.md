# `docconvert.exporters` — 输出格式化器

> 源码: [`docconvert/exporters/`](../../docconvert/exporters/)
> `__all__`: `["BaseExporter", "HtmlExporter", "MarkdownExporter", "JsonExporter", "get_exporter"]`

Exporter 把转换器返回的**中间表示**序列化为**最终字符串**。每个 exporter 都是无状态、纯函数式(只读 `self.config`)。

---

## `class BaseExporter`(ABC)

```python
from abc import ABC, abstractmethod
from typing import Any, Optional
from docconvert.config import AppConfig, DEFAULT_CONFIG

class BaseExporter(ABC):
    def __init__(self, config: Optional[AppConfig] = None) -> None: ...
    @abstractmethod
    def export(self, data: Any, **kwargs) -> str: ...
```

| 方法 | 说明 |
|------|------|
| `__init__(config)` | 持有配置(子类可读 `self.config`) |
| `export(data, **kwargs)` | 接收转换器传入的中间值,返回字符串 |

`data` 实际类型由转换器约定:

| 转换器 | data 类型 |
|--------|----------|
| `ExcelConverter` | `dict`(`metadata` / `headers` / `rows` / `merged_cells`) |
| `WordConverter`(html) | `str`(已含完整 HTML 文档) |
| `WordConverter`(md) | `str`(已含 YAML 注释 + cleaned markdown) |
| `WordConverter`(json) | `dict`(`metadata` + `content`) |
| `DocConverter`(html) | `str` |
| `DocConverter`(md) | `str` |
| `DocConverter`(json) | `dict`(`source` + `content`) |

---

## 内置 Exporter

### `class HtmlExporter`

```python
class HtmlExporter(BaseExporter):
    def export(self, data: Any, **kwargs) -> str: ...
```

行为:

- `data` 已经是 `str` → 原样返回
- `data` 是 `dict` → 返回 `data.get('content', str(data))`
- 其他 → `str(data)`

> 当前 `WordConverter` / `DocConverter` 已经把完整 HTML 文档壳构造好,exporter 只做 pass-through。如果转换器改了返回 `dict`,exporter 会自动从 `content` 字段抽取。

### `class MarkdownExporter`

```python
class MarkdownExporter(BaseExporter):
    def export(self, data: Any, **kwargs) -> str: ...
```

行为与 `HtmlExporter` 相同 — pass-through。命名分开只是为了和 `get_exporter('md')` 对应。

### `class JsonExporter`

```python
class JsonExporter(BaseExporter):
    def export(self, data: Any, **kwargs) -> str: ...
```

- `data` 是 `str` → 原样返回
- `data` 是 `dict` → `json.dumps(data, ensure_ascii=False, indent=2)`(中文不被转义)
- 其他 → `str(data)`

> 注:Excel 转换器在调 `self._export()` 之前**自己** `json.dumps` 一次,所以落到 `JsonExporter.export` 的就是字符串。

---

## `get_exporter(fmt)` — 工厂函数

```python
def get_exporter(fmt: str) -> BaseExporter: ...
```

| `fmt` | 返回 |
|-------|------|
| `'html'` | `HtmlExporter()` |
| `'md'` | `MarkdownExporter()` |
| `'json'` | `JsonExporter()` |
| 其他 | 抛 `ValueError('不支持的输出格式: {fmt}')` |

> 错误信息为中文,沿用 DocConvert 整体风格。

### 示例

```python
from docconvert.exporters import get_exporter

exp = get_exporter('json')
text = exp.export({'metadata': {'source': 'a.xlsx'}, 'rows': []})
# '{"metadata": {"source": "a.xlsx"}, "rows": []}'
```

---

## 扩展指南:自定义 Exporter

```python
from typing import Any
from docconvert.exporters import BaseExporter

class YamlExporter(BaseExporter):
    def export(self, data: Any, **kwargs) -> str:
        import yaml
        if isinstance(data, str):
            return data
        if isinstance(data, dict):
            return yaml.safe_dump(data, allow_unicode=True, sort_keys=False)
        return str(data)
```

在 `docconvert/exporters/__init__.py` 注册:

```python
from docconvert.exporters.yaml_exporter import YamlExporter

def get_exporter(fmt: str) -> BaseExporter:
    if fmt == 'yaml':
        return YamlExporter()
    ...
```

并在 `ConversionController` 之外的所有调用方支持 `output_fmt='yaml'`。

---

## 设计要点

| 关注点 | 实现 |
|--------|------|
| **无副作用** | `export` 只做 `data → str` 转换,不写文件 |
| **可单测** | 每个 exporter 接受任意 `data`,无需文件系统 |
| **命名一致性** | 工厂 / 转换器 / CLI / GUI 共用 `html` / `md` / `json` 三个 magic string |
| **错误统一** | 未知 `fmt` 在工厂层抛 `ValueError`,转换器层不重复校验 |
