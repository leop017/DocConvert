# DocConvert API 参考文档

本目录包含 DocConvert 的 Python API 完整参考。DocConvert 提供 **GUI**、**CLI**、**Python API** 三种使用方式,本页主要面向希望通过 Python 代码集成 DocConvert 的开发者。

***

## 快速开始

```python
from docconvert.controller import ConversionController
from docconvert.config import DEFAULT_CONFIG
from docconvert.models import ProgressEvent

def on_progress(event: ProgressEvent) -> None:
    print(f'[{int(event.progress * 100):3d}%] {event.message}')

controller = ConversionController(DEFAULT_CONFIG)
results = controller.convert_files(
    files=['report.xlsx', 'manual.docx'],
    output_fmt='md',
    output_dir='./out',
    enhanced_md=True,
    progress_callback=on_progress,
)
for name, path, err in results:
    if err:
        print(f'FAIL {name}: {err}')
    else:
        print(f'OK   {name} -> {path}')
```

***

## 模块总览

| 模块                      | 作用                                   | 文档                             |
| ----------------------- | ------------------------------------ | ------------------------------ |
| `docconvert.controller` | 转换调度核心(同步/异步/取消/覆盖检测)                | [controller.md](controller.md) |
| `docconvert.converters` | 各类输入文件的解析器(Excel/Word/Doc)           | [converters.md](converters.md) |
| `docconvert.exporters`  | 输出格式化器(HTML / Markdown / JSON)       | [exporters.md](exporters.md)   |
| `docconvert.cleaners`   | Markdown 后处理清洗器                      | [cleaners.md](cleaners.md)     |
| `docconvert.models`     | 数据模型(`MergeInfo` / `ProgressEvent`)  | [models.md](models.md)         |
| `docconvert.config`     | 全局配置(`AppConfig` / `DEFAULT_CONFIG`) | [config.md](config.md)         |
| `docconvert.utils`      | 通用工具函数                               | [utils.md](utils.md)           |
| `docconvert.logger`     | 日志                                   | [logger.md](logger.md)         |
| `docconvert.cli`        | 命令行入口                                | [cli.md](cli.md)               |
| `docconvert.gui`        | Tkinter 图形界面(懒加载)                    | [gui.md](gui.md)               |

***

## 公开 API 一览(从 `__all__` 自动汇总)

```text
docconvert
└── DocConvertApp                    # 懒加载

docconvert.config
├── AppConfig
└── DEFAULT_CONFIG

docconvert.controller
└── ConversionController

docconvert.converters
├── BaseConverter
├── ExcelConverter
├── WordConverter
└── DocConverter

docconvert.exporters
├── BaseExporter
├── HtmlExporter
├── MarkdownExporter
├── JsonExporter
└── get_exporter

docconvert.cleaners
├── BaseCleaner
└── WordMdCleaner

docconvert.models
├── MergeInfo
└── ProgressEvent

docconvert.utils
├── INVALID_NAMES
├── safe_str
├── clean_filename
├── escape_md_cell
├── html_to_md
├── get_excel_sheet_names
├── decode_text
└── unique_cleaned_suffixes

docconvert.logger
├── LOGGER_NAME
├── setup_logging
└── get_logger
```

***

## 典型使用模式

### 1. 直接使用某个具体转换器

```python
from pathlib import Path
from docconvert.converters import ExcelConverter
from docconvert.config import DEFAULT_CONFIG

conv = ExcelConverter(DEFAULT_CONFIG)
results, errors = conv.convert(
    input_path='finance.xlsx',
    output_fmt='md',
    parent=Path('./out'),
    enhanced_md=True,
    sheets=['Q1', 'Q2'],  # 仅转换这些 sheet
)
```

### 2. 通过 `ConversionController` 批量调度

```python
controller = ConversionController()
existing = controller.check_overwrite_paths(
    files=['a.xlsx', 'b.docx'],
    output_fmt='md',
    output_dir='./out',
)
if existing:
    confirm = input(f'将覆盖 {len(existing)} 个文件,是否继续? (y/N) ')
    if confirm.lower() != 'y':
        raise SystemExit(0)

results = controller.convert_files(
    files=['a.xlsx', 'b.docx'],
    output_fmt='md',
    output_dir='./out',
    enhanced_md=True,
)
```

### 3. 在后台线程运行

```python
import threading
from docconvert.controller import ConversionController

controller = ConversionController()
controller.convert_files_async(
    files=['big.xlsx'],
    output_fmt='json',
    progress_callback=lambda e: print(e.message) if e.message else None,
)

# ... 业务侧可随时取消
# controller.cancel()

# 阻塞等待完成
if controller.wait_done(timeout=600):
    print('全部完成', controller.last_results)
else:
    print('超时')
```

### 4. 自定义清洗规则

```python
from docconvert.config import AppConfig
from docconvert.cleaners import WordMdCleaner

cfg = AppConfig(cleaning_rules={
    'remove_page_numbers': True,
    'remove_duplicate_headers': False,  # 关闭去重
    'remove_empty_lines': True,
    'normalize_spaces': True,
})
cleaner = WordMdCleaner(cfg)
print(cleaner.clean(messy_md))
```

### 5. 自定义导出器(扩展点)

```python
from docconvert.exporters import BaseExporter
from docconvert.exporters import get_exporter

class CsvExporter(BaseExporter):
    def export(self, data, **kwargs) -> str:
        # 自定义实现
        ...

# 不修改源码的情况下注册: 直接 new 一个并使用
exporter = CsvExporter()
text = exporter.export({'content': 'hello'})
```

***

## 设计原则

1. **零 Tkinter 依赖**: 不使用 GUI 时不会触发 `tkinter` 导入。`DocConvertApp` 通过 `__getattr__` 懒加载。
2. **线程安全**: `ConversionController` 内部使用 `threading.Event` / `threading.Lock` 保证取消与状态一致性。
3. **可扩展**: 所有转换器、导出器、清洗器均继承自 ABC,只需实现对应方法即可插入新格式。
4. **配置驱动**: 通过 `AppConfig` 集中管理清洗规则、块大小等可调参数,无全局可变状态。
5. **失败隔离**: 单个文件失败不会影响批内其他文件,错误会出现在返回值/回调中。

***

## 下一步

- 新接入一种文件格式?参考 [converters.md](converters.md) 的"扩展指南"。

- 想了解控制器内部状态机?阅读 [controller.md](controller.md)。

- 想了解 Markdown 清洗的 4 阶段流水线?阅读 [cleaners.md](cleaners.md)。

