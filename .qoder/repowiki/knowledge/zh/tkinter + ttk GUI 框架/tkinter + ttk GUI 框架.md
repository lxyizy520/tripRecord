---
kind: external_dependency
name: tkinter + ttk GUI 框架
slug: tkinter-ttk
category: external_dependency
category_hints:
    - framework_behavior
scope:
    - '**'
source_files:
    - 差旅记录.pyw
---

桌面端 UI 基于 Python 内置 tkinter/ttk 构建，采用单文件架构：主窗口用 ttk.Notebook 组织三个标签页（行程记录、住宿记录、餐饮/交通），新增/编辑操作通过继承 tk.Toplevel 的独立 Dialog 类实现（EditDialog、LodgingDialog、MealDialog、TransportDialog）。字体统一使用 Microsoft YaHei UI。打包为 exe 后仍依赖系统安装的 Python 运行时或 PyInstaller 内嵌环境。