# `docconvert.models` — 数据模型

> 源码: [`docconvert/models/models.py`](../../docconvert/models/models.py)
> `__all__`: `["MergeInfo", "ProgressEvent"]`

两个轻量 `dataclass`,承载模块间的数据传递。无业务逻辑,无外部依赖。

---

## `class MergeInfo`

```python
@dataclass
class MergeInfo:
    rowspan: int = 1
    colspan: int = 1
    is_master: bool = False
    is_merged: bool = False
    min_row: int = 0
    min_col: int = 0
    max_row: int = 0
    max_col: int = 0
```

描述 Excel 工作表中**单个单元格**的合并信息(由 `ExcelConverter` 内部使用,通常不直接面向终端用户)。

### 字段

| 字段 | 类型 | 默认 | 含义 |
|------|------|------|------|
| `rowspan` | `int` | `1` | 垂直跨度(行数) |
| `colspan` | `int` | `1` | 水平跨度(列数) |
| `is_master` | `bool` | `False` | 是否为合并区域的**左上角主单元格** |
| `is_merged` | `bool` | `False` | 是否为合并区域的**从单元格** |
| `min_row` | `int` | `0` | 1-based 起始行 |
| `min_col` | `int` | `0` | 1-based 起始列 |
| `max_row` | `int` | `0` | 1-based 结束行 |
| `max_col` | `int` | `0` | 1-based 结束列 |

> 主单元格 = `is_master=True`,从单元格 = `is_master=False and is_merged=True`。
> `rowspan == 1 and colspan == 1` 的普通单元格:`is_master=False, is_merged=False`。

### 用法

- 通常**不直接构造**,由 `ExcelConverter._build_merged_map` 根据 openpyxl / xlrd 合并区域填充
- JSON 导出时,合并信息进入 `merged_cells` 字段(供下游 RAG 重建表格布局)

---

## `class ProgressEvent`

```python
@dataclass
class ProgressEvent:
    message: str = ""
    progress: float = 0.0
    done: bool = False
    error: Optional[str] = None
```

`ConversionController` 通过 `progress_callback` 推送的事件载体。

### 字段

| 字段 | 类型 | 默认 | 含义 |
|------|------|------|------|
| `message` | `str` | `""` | 人类可读的进度消息(中文) |
| `progress` | `float` | `0.0` | 0.0 – 1.0 的进度比例;`done=True` 时通常为 1.0 |
| `done` | `bool` | `False` | 整批是否已结束(包括成功 / 取消 / 错误) |
| `error` | `Optional[str]` | `None` | 若本事件携带错误,这里填错误描述 |

### 典型事件流

```text
[1] message="处理文件 1/3: a.xlsx"   progress=0.17
[2] message="完成 1/3"              progress=0.33
[3] message="处理文件 2/3: b.docx"   progress=0.50
[4] message="转换失败: b.docx"        progress=0.67  error="..."
[5] message="完成 2/3"              progress=0.67
[6] message="处理文件 3/3: c.xls"    progress=0.83
[7] message="完成 3/3"              progress=1.00
[8] message="转换完成"              progress=1.00  done=True
```

### 取消事件流

```text
[...]
[N]   message="处理文件 2/3: b.docx" progress=0.50
[N+1] message="完成 2/3"             progress=0.67
[N+2] message="转换已取消"           progress=0.67  done=True
```

`controller.was_cancelled == True`。

### 回调最佳实践

```python
def on_progress(event: ProgressEvent) -> None:
    if event.error:
        logger.warning('%s: %s', event.message, event.error)
    if event.message:
        status_bar.set(event.message)
    if event.progress:
        progress_bar.set(event.progress)
    if event.done:
        status_bar.set('就绪')
```

> **不要**在回调里抛异常。`ConversionController._report` 内部会 try/except 兜底,但异常仅记 DEBUG 日志。

---

## 扩展:自定义事件

`ProgressEvent` 故意保持扁平。如果未来需要携带更复杂的状态(如 `cancelled: bool` 或 `file_index: int`),可:

- 升级为 `class ProgressEvent(BaseModel)`(增加字段)
- 或新增 `ProgressDetail(ProgressEvent)` 派生 dataclass

接口兼容原则:`progress_callback` 应只依赖 `message` / `progress` / `done` / `error` 这 4 个现有字段。
