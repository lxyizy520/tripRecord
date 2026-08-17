---
kind: error_handling
name: 基于 tkinter messagebox 的轻量级错误处理：输入校验、静默降级与顶层异常兜底
category: error_handling
scope:
    - '**'
source_files:
    - 差旅记录.pyw
---

## 1. 整体方案
该仓库是一个单文件 Python + tkinter + SQLite 桌面应用（`差旅记录.pyw`），没有独立的错误模块、自定义异常类或日志框架。错误处理采用“前端输入即时校验 + 关键路径 try/except 兜底 + 静默降级”的轻量模式，所有用户可见的错误通过 `tkinter.messagebox` 弹窗反馈。

## 2. 关键位置与模式
- **顶层异常兜底**：`main()` 函数用 `try/except Exception as e` 包裹整个程序生命周期（`db_init()` → `MainApp().mainloop()`），捕获后调用 `messagebox.showerror('程序错误', str(e))`；若弹窗本身失败则重新抛出。这是唯一的全局异常处理器。
- **输入校验（对话框层）**：所有新增/编辑对话框（`EditDialog._save`、`LodgingDialog._save`、`MealDialog._save`、`TransportDialog._save`）在写入数据库前进行参数校验：日期格式使用 `parse_date` 检查，必填字段为空时调用 `messagebox.showwarning('提示', ...)` 并 `return` 中止保存；金额字段用 `float()` 转换并捕获 `ValueError` 提示“金额请输入数字”。
- **查询过滤校验**：主窗口 `refresh` / `refresh_lodging` 对日期筛选框也调用 `parse_date`，非法时弹出警告并终止刷新。
- **截图预览容错**：`_preview_shot` 中 `Image.open` 被 `try/except Exception` 包裹，失败时仅将标签文本改为“无法预览该图片”，不中断 UI。
- **文件系统操作静默降级**：删除截图的 `remove_screenshot_file` 和 `delete_trip` 中的 `os.remove` 均使用 `try/except OSError: pass`，即“找不到文件或权限不足”等 IO 错误被忽略，保证删除流程不因残留文件而中断。
- **导出后打开文件**：`os.startfile(path)` 被 `try/except OSError: pass` 包裹，防止系统无法打开文件导致崩溃。
- **数据库连接管理**：每个 `get_conn()` 返回的连接都在 `try/finally` 中关闭，避免连接泄漏；SQL 执行本身不捕获异常，交由上层或顶层 `main` 处理。

## 3. 架构约定
- **无自定义异常类型**：代码未定义任何业务异常类，也不向上抛出结构化错误码；UI 层直接消费返回值（如 `export_*_to_excel` 返回 `(None, 0)` 表示无可导出数据）并通过 `messagebox` 告知用户。
- **无日志系统**：未见 `logging` 模块导入或日志文件输出，调试依赖控制台或打包后的 exe 运行环境。
- **无中间件/装饰器**：错误处理以内联 `try/except` 和前置 `if` 校验为主，没有统一的拦截层。
- **用户提示文案风格统一**：警告标题固定为 `'提示'`，错误标题为 `'程序错误'`，导出成功标题为 `'导出成功'`，便于用户识别。

## 4. 约束与规则（从实现中可观察到的约定）
- 所有用户可触发的输入错误必须在保存/提交前拦截并以 `messagebox.showwarning` 提示，不得让脏数据进入数据库。
- 非致命 I/O 错误（截图删除、打开文件）一律 `except OSError: pass`，保持业务流程继续。
- 不可恢复的运行时异常由 `main()` 顶层 `except Exception` 捕获并弹窗，不再向上传播。
- 数据库层函数不对外抛出异常，调用方只关心返回值（如 `query_*` 返回列表、`add_*` 返回 lastrowid）。
- 截图预览失败属于“弱错误”，仅改变 UI 文本，不阻断用户继续操作。

## 5. 适用性说明
该类别适用于本仓库——虽然错误处理非常朴素，但确实存在明确的约定：输入校验 + 静默降级 + 顶层兜底，全部通过 `tkinter.messagebox` 呈现给用户。