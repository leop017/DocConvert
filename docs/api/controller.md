# `docconvert.controller` — 转换调度核心

> 源码: [`docconvert/controller/conversion_controller.py`](../../docconvert/controller/conversion_controller.py)
> `__all__`: `["ConversionController"]`

`ConversionController` 是 DocConvert 的**唯一推荐入口**。它负责:

- 决定每个输入文件使用哪个 `BaseConverter`
- 协调同步 / 异步执行
- 支持中途取消
- 预先计算覆盖路径,方便业务侧弹确认框
- 保证批内多个文件的输出路径互不冲突(同 basename 自动加 `_parent` 或 `_2`、`_3` 后缀)

---

## `class ConversionController`

```python
class ConversionController:
    def __init__(self, config: Optional[AppConfig] = None) -> None: ...
```

### 构造参数

| 参数 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `config` | `Optional[AppConfig]` | `None` | 全局配置。为 `None` 时使用 `DEFAULT_CONFIG`。 |

### 实例属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `config` | `AppConfig` | 当前生效的配置。 |
| `is_running` | `bool`(property) | 后台线程是否仍在运行。 |
| `was_cancelled` | `bool` | 最近一次批处理是否被用户取消(部分文件可能已成功)。 |
| `last_results` | `list[tuple[str, str, Optional[str]]]` | 最近一次 `convert_files` 的完整结果。 |
| `last_error` | `Optional[str]` | 最近一次 `convert_files_async` 顶层异常信息。 |

### 类型别名

```python
ProgressCallback = Callable[[ProgressEvent], None]
```

---

## 公开方法

### `convert_files` — 同步批量转换

```python
def convert_files(
    self,
    files: list[str],
    output_fmt: str,
    output_dir: Optional[str] = None,
    enhanced_md: bool = False,
    sheets: Optional[list[str]] = None,
    progress_callback: Optional[ProgressCallback] = None,
) -> list[tuple[str, str, Optional[str]]]:
```

同步执行整批转换,返回每条结果。

| 参数 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `files` | `list[str]` | 必填 | 输入文件绝对 / 相对路径列表。 |
| `output_fmt` | `str` | 必填 | `'html'` / `'md'` / `'json'`。 |
| `output_dir` | `Optional[str]` | `None` | 输出目录。`None` 时落到"批内第一个文件的父目录"(见下方 _批目录规则_)。 |
| `enhanced_md` | `bool` | `False` | 是否对 Word/Doc 走增强 Markdown 路径(HTML 中转 + 清洗)。 |
| `sheets` | `Optional[list[str]]` | `None` | Excel 工作表过滤;`None` 表示全部。Word/Doc 显式传参会抛 `TypeError`。 |
| `progress_callback` | `Optional[ProgressCallback]` | `None` | 进度回调。 |

**返回**: `list[tuple[filename, output_path, error_or_None]]`

- 成功 → `(name, '/abs/out/x.md', None)`
- 失败 → `(name, '', '错误描述')`
- 文件不存在 → `(filename, '', '文件不存在')`

**取消语义**: 任意文件处理前/后都会检查 `cancel_event`。若被取消,`self.was_cancelled` 置为 `True`,已生成的文件保留,未开始的跳过,函数正常返回。

**抛出**:
- `ValueError` — 输出目录无法创建(底层 `OSError` 被包装)
- 转换器自身抛出的异常会被捕获并写入结果项,不会向上冒泡

#### 示例

```python
results = controller.convert_files(
    files=['a.xlsx', 'b.docx'],
    output_fmt='md',
    output_dir='./out',
    enhanced_md=True,
    progress_callback=lambda e: print(e.message),
)
ok = [r for r in results if not r[2]]
print(f'{len(ok)}/{len(results)} succeeded')
```

---

### `convert_files_async` — 异步批量转换

```python
def convert_files_async(
    self,
    files: list[str],
    output_fmt: str,
    output_dir: Optional[str] = None,
    enhanced_md: bool = False,
    sheets: Optional[list[str]] = None,
    progress_callback: Optional[ProgressCallback] = None,
) -> bool:
```

在后台 `Thread` 中执行 `convert_files`,立即返回。

- 已有任务在跑 → 返回 `False` 并记日志
- 成功启动 → 返回 `True`
- 任务结束后通过 `last_results` / `last_error` 取结果

> 后台线程为 `daemon=True`,主进程退出时会被强制终止,**不要依赖它做关键落盘**之外的副作用。

---

### `cancel` / `reset_cancel`

```python
def cancel(self) -> None: ...
def reset_cancel(self) -> None: ...
```

- `cancel()` — 置位 cancel event。下一次循环检查点会停止后续文件。
- `reset_cancel()` — 清空 cancel event。`convert_files` 入口处会自动调用。

---

### `wait_for_completion` / `wait_done`

```python
def wait_for_completion(self, timeout: Optional[float] = None) -> bool: ...
def wait_done(self, timeout: Optional[float] = None) -> bool: ...
```

| 方法 | 机制 | 推荐场景 |
|------|------|---------|
| `wait_for_completion` | `Thread.join(timeout)` | 想复用线程句柄 |
| `wait_done` | `Event.wait(timeout)` | 想被多路复用 / 跨线程通知 |

两者都返回 `True` 表示在超时前完成。

---

### `set_config`

```python
def set_config(self, config: AppConfig) -> bool: ...
```

运行时切换配置。**仅在空闲时**(非 `is_running`)生效,否则忽略并返回 `False`、记 WARNING 日志。

---

### `check_overwrite_paths` — 预先计算将被覆盖的文件

```python
def check_overwrite_paths(
    self,
    files: list[str],
    output_fmt: str,
    output_dir: Optional[str] = None,
    sheets: Optional[list[str]] = None,
) -> list[str]:
```

在真正写入前,先算出每个文件将要生成的输出路径,并返回**磁盘上已存在**的那些路径。

- 与 `convert_files` 使用**完全相同**的目录解析和 basename 去重逻辑(共享 `_batch_parent` / `_make_unique_stem` / `_compute_output_paths`)
- 典型用法:弹"将覆盖以下 N 个文件,是否继续?"对话框

```python
existing = controller.check_overwrite_paths(files, 'md', './out')
if existing and not confirm_overwrite(existing):
    raise SystemExit(0)
controller.convert_files(files, 'md', './out')
```

---

### `is_running` (property)

```python
@property
def is_running(self) -> bool: ...
```

后台线程是否仍在运行。注意 `cancel()` 之后线程会很快退出,但中间有几毫秒的窗口 `is_running` 仍为 `True`。

---

## 内部辅助(了解即可,非公开 API)

| 方法 | 说明 |
|------|------|
| `_batch_parent(files, output_dir)` | 解析批目录:`output_dir` 优先,否则取 `files[0]` 的父目录。 |
| `_make_unique_stem(base, source, used)` | 在批内生成唯一 stem:先尝试原 stem,撞了则加 `_parent_dir_name`,再撞则 `_n`。 |
| `_get_converter(ext)` | 按扩展名返回 `ExcelConverter` / `WordConverter` / `DocConverter`,未知扩展名抛 `ValueError`。 |
| `_compute_output_paths(input, fmt, dir, sheets, stem_override)` | 列出该输入会产出的所有输出文件路径(用于 `check_overwrite_paths`)。 |
| `_report(callback, event)` | 包装一层 try/except 防止回调异常打断主流程。 |

---

## 内部状态机

```
                 +-----------+
   convert_files |  IDLE     |  cancel()
   /            |  _running | <----+
   v            |  =False   |      |
+----+----+      +-----+-----+      |
|         |            |            |
|         v            v            |
|  +------------+  +-----------+    |
|  | RUNNING    |  |  CANCEL   |----+
|  | _running=T |  | _running=F|
|  | _done_event|  | _done=T   |
|  |     clear  |  | results   |
|  +-----+------+  | preserved |
        |          +-----------+
        | cancel()
        +----------------------+
                               |
                               v
                       _cancel_event.set
                       _done_event.set
                       _running=False
```

---

## 异常与边界

| 场景 | 行为 |
|------|------|
| `files=[]` | 立即返回 `[]`,不发任何 `ProgressEvent` |
| `output_dir` 不存在 | 尝试 `mkdir(parents=True, exist_ok=True)`,失败抛 `ValueError` |
| 单个文件抛异常 | 捕获并写入 `results`,不影响其他文件 |
| 未知扩展名 | `_get_converter` 抛 `ValueError`,该文件 result 为 `(filename, '', str(e))` |
| 同时调用 `convert_files_async` | 第二次返回 `False`,**不会**排队 |
| 后台线程内异常 | 写入 `self.last_error`,线程退出,`wait_done()` 仍能解除 |

---

## 完整示例(GUI 中的典型用法)

```python
from docconvert.controller import ConversionController
from docconvert.config import DEFAULT_CONFIG
from docconvert.models import ProgressEvent

class Worker:
    def __init__(self):
        self.controller = ConversionController(DEFAULT_CONFIG)
        self.on_done = lambda results: None
        self.on_progress = lambda ev: None

    def start(self, files, fmt, outdir):
        def cb(ev: ProgressEvent):
            self.on_progress(ev)

        self.controller.convert_files_async(
            files=files,
            output_fmt=fmt,
            output_dir=outdir,
            enhanced_md=True,
            progress_callback=cb,
        )
        # ... 通过 wait_done 轮询 / Event 通知主线程
```
