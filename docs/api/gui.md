# `docconvert.gui` — Tkinter 图形界面

> 源码: [`docconvert/gui/app.py`](../../docconvert/gui/app.py)
> 公开符号: `DocConvertApp`(在包顶层通过 `__getattr__` 懒加载,见 `docconvert/__init__.py`)

GUI 是 DocConvert 的**可选**入口。设计上做了**懒加载** — 不调 `DocConvertApp` 就不会触发 `tkinter` 导入,因此 CLI / Python API 用户无需安装 GUI 依赖。

---

## 懒加载机制

```python
# docconvert/__init__.py
__all__ = ["DocConvertApp"]

def __getattr__(name: str):
    if name == "DocConvertApp":
        from docconvert.gui.app import DocConvertApp
        return DocConvertApp
    raise AttributeError(...)
```

这意味着:

```python
import docconvert                          # 不会触发 tkinter
docconvert.DocConvertApp                   # 此时才 import tkinter
```

> 在**无 GUI**(headless / Docker / CI)环境也能正常 import 包。

---

## `class DocConvertApp`

```python
class DocConvertApp:
    def __init__(self) -> None: ...
    def run(self) -> None: ...
```

### 构造

无参数。构造时:

1. 创建 Tk 根窗口
2. 应用自定义配色(`#bg` / `#card_bg` / `#accent` / `#success` / `#error` 等 8 种色)
3. 构造顶部栏、文件选择卡、文件列表卡、格式卡、预览卡、底部状态栏
4. 注册 4 个清洗规则 checkbox(**仅当**选中 Word + Markdown 时启用)

### `run()`

进入 Tk 主事件循环,直到窗口关闭。

---

## 窗口结构

```text
┌────────────────────────────────────────────────────────┐
│  DocConvert                              v2.0.x        │  ← title bar
├────────────────────────────────────────────────────────┤
│  [ 选择文件 ]  [ 清空 ]                                │  ← file select card
├────────────────────────────────────────────────────────┤
│  ☐ a.xlsx                                              │
│  ☐ b.docx                                              │  ← file list card
│  ☐ c.xls                                               │
├────────────────────────────────────────────────────────┤
│  输出格式:  ◉ HTML  ◯ Markdown  ◯ JSON                  │  ← format card
│  清洗规则(仅 Word+MD):                                 │
│   ☑ 删除页码  ☑ 去重表头  ☑ 折叠空行  ☑ 归一化空格    │
├────────────────────────────────────────────────────────┤
│  ┌────────────────────────────────────────────────┐    │
│  │ <预览区: 选中文本后显示>                        │    │  ← preview card
│  └────────────────────────────────────────────────┘    │
├────────────────────────────────────────────────────────┤
│  [ 开始转换 ]  [ 取消 ]      状态:就绪   进度: ▱▱▱    │  ← bottom bar
└────────────────────────────────────────────────────────┘
```

---

## 线程模型

GUI 主线程跑 Tk 事件循环,转换跑在 `ConversionController` 启动的后台 `Thread` 中。两者通过 `queue.Queue` 通信:

```text
worker thread                  GUI main thread
─────────────                  ────────────────
ConversionController
   └─ progress_callback(e) ─┐
                            │  enqueue
                            ▼
                       _call_queue ── root.after(50, _poll_pending) ──> UI 更新
```

- `root.after(50, _poll_pending)` 每 50 ms 排空队列
- 这样 Tk widget 只在主线程被 touch,天然线程安全

---

## 覆盖确认

点击"开始转换"时:

1. `ConversionController.check_overwrite_paths(...)` 算出会覆盖的现有文件
2. 若非空,弹 `messagebox.askyesno("将覆盖以下 N 个文件,是否继续?")`
3. 用户点"否"→ 取消;点"是"→ 进入 `convert_files_async`

---

## 自动窗口尺寸

启动时根据卡片内容自动 `winfo_reqwidth` / `winfo_reqheight`,避免小屏 / 高 DPI 下出现滚动条。窗口可手动拉伸。

---

## 入口

```python
# main.py
if __name__ == '__main__':
    if len(sys.argv) == 1:
        from docconvert.gui.app import DocConvertApp
        DocConvertApp().run()
```

或从命令行:

```bash
python main.py            # 启动 GUI
docconvert                # 已 pip install 时,空参数会触发 GUI(若注册了默认子命令)
```

> 当前 `pyproject.toml` 仅注册 `docconvert = "docconvert.cli:main_cli"`,所以**已安装版本**下,`docconvert` 不会自动启动 GUI;需要 `python -m docconvert.gui` 或 `python main.py`。

---

## 可访问性 / 国际化

- 当前界面文案为**中文**
- 字体优先 `Microsoft YaHei`(Windows / 国内 Linux)
- 颜色对比度按 WCAG AA 校准
- 无 RTL 支持

如需 i18n,需修改 `app.py` 中硬编码字符串并接入 `gettext`。

---

## 设计要点

| 关注点 | 实现 |
|--------|------|
| **Tkinter 隔离** | 通过 `__getattr__` 懒加载,无 GUI 场景不付出 import 成本 |
| **线程安全** | `queue.Queue` + `root.after` 轮询,避免跨线程直接改 widget |
| **覆盖确认** | `check_overwrite_paths` 复用同一目录解析逻辑,与实际写入路径严格一致 |
| **响应式布局** | 自动 `winfo_reqwidth/height`,不强制最小尺寸 |
| **取消** | GUI"取消"按钮 → `controller.cancel()`;后台线程在下一次检查点退出 |
