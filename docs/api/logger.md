# `docconvert.logger` — 日志

> 源码: [`docconvert/logger.py`](../../docconvert/logger.py)
> 公开符号: `LOGGER_NAME`, `setup_logging`, `get_logger`

DocConvert 使用 Python 标准库 `logging`,**所有组件共享同一个命名 logger** `docconvert`。

---

## `LOGGER_NAME`

```python
LOGGER_NAME = "docconvert"
```

> 业务代码如需获取 logger:`logging.getLogger("docconvert")` 或 `docconvert.logger.get_logger()`。

---

## `setup_logging(level)`

```python
def setup_logging(level: Union[int, str] = logging.INFO) -> logging.Logger: ...
```

初始化 / 更新命名 logger。

| 参数 | 默认 | 说明 |
|------|------|------|
| `level` | `logging.INFO` | `int` (如 `logging.DEBUG`)或 `str` (如 `"DEBUG"`) |

**行为**:
- 第一次调用:创建 `StreamHandler(sys.stdout)`,format 为 `[HH:MM:SS] LEVEL     docconvert - message`
- 后续调用:**同时**更新 logger 和**所有已注册 handler** 的 level(否则 handler 会按旧 level 过滤,导致看似设置 DEBUG 实际无变化)

**返回**:`logging.Logger` 实例(便于链式调用)。

### 关键修复(代码注释明示)

```python
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(level)
    ...
    logger.addHandler(handler)
else:
    # A later call (e.g. main.py -> main_cli with --verbose) must
    # raise the already-registered handler's level too; setting only
    # the logger's level is not enough, the handler would keep
    # filtering the messages out.
    for h in logger.handlers:
        h.setLevel(level)
```

> 这是 v2.0.x 修复的真实 bug:`--verbose` 只改 logger 级别 → handler 仍按 INFO 过滤 → DEBUG 日志不显示。

### 用法

```python
from docconvert.logger import setup_logging

setup_logging()              # INFO
setup_logging('DEBUG')       # 切换为 DEBUG
setup_logging(logging.WARNING)
```

CLI 中通过 `--verbose` 自动调用 `setup_logging('DEBUG')`(`cli.py:87`)。

---

## `get_logger()`

```python
def get_logger() -> logging.Logger: ...
```

获取命名 logger,不会自动初始化 handler。如果业务代码在 `setup_logging` 之前调用,handler 为空、INFO+ 日志会被吞。

**推荐**:
- 应用入口(`main.py` / `cli.main_cli`)先调 `setup_logging`
- 业务模块直接 `get_logger()`

```python
from docconvert.logger import get_logger
logger = get_logger()
logger.info('开始处理 %s', filename)
```

---

## 日志格式

```text
[14:23:01] INFO     docconvert - 开始处理 report.xlsx
[14:23:02] DEBUG    docconvert - Sheet 'Q1' 共 150 行
[14:23:02] WARNING  docconvert - 转换失败 [b.docx]: 文件不存在
[14:23:03] ERROR    docconvert - 转换异常 [c.xls]: ...
```

> 输出到 **stdout**(不是 stderr),便于管道(`| tee log.txt`)。

---

## 与第三方日志库集成

如果使用 `loguru` / `structlog`,在 `setup_logging` 之后接入:

```python
from docconvert.logger import setup_logging
import logging

setup_logging('INFO')
logging.getLogger('docconvert').info = loguru_logger.info  # 不推荐,会破坏 level 处理
```

> 推荐做法:**保留** `docconvert` 命名 logger,在应用层用 `logging.config.dictConfig` 统一接管。
