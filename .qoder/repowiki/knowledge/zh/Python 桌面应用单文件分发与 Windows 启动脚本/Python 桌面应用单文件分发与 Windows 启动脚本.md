---
kind: build_system
name: Python 桌面应用单文件分发与 Windows 启动脚本
category: build_system
scope:
    - '**'
source_files:
    - 差旅记录.pyw
    - 启动差旅记录.bat
    - 差旅记录.exe
---

## 1. 使用的系统/方法

该项目是一个基于 Python + tkinter + SQLite 的本地桌面程序，**没有使用传统的构建系统（Makefile、Dockerfile、CI 流水线等）**。其“构建”与“发布”方式非常轻量：
- 源代码为单个 Python 脚本 `差旅记录.pyw`。
- 通过第三方打包工具（从根目录存在 `差旅记录.exe` 可推断使用了 PyInstaller 或类似工具）将 `.pyw` 打包成独立可执行文件 `差旅记录.exe`。
- 用户通过 Windows 批处理脚本 `启动差旅记录.bat` 双击启动。

## 2. 关键文件

- `差旅记录.pyw`：唯一的应用源码入口，包含全部业务逻辑（数据库初始化、CRUD、截图管理、Excel 导出、tkinter UI）。代码中通过 `sys.frozen` 判断是否处于打包后的 exe 环境，从而动态确定数据目录位置（见第 21–29 行）。
- `启动差旅记录.bat`：Windows 启动脚本，设置 UTF-8 编码 (`chcp 65001`)，切换到脚本所在目录后调用 `差旅记录.exe`。
- `差旅记录.exe`：已打包好的可执行文件，随仓库一起分发。
- `data/`：运行时数据目录，包含 SQLite 数据库 `差旅记录.db` 和按日期组织的截图子目录 `screenshots/`。
- `导出文件/`：运行时生成的 Excel 导出目录。
- `reasonix.toml`：Reasonix 元数据配置文件，与构建无关。

## 3. 架构与约定

- **单文件应用**：所有功能集中在一个 `.pyw` 文件中，无模块拆分、无 `setup.py` / `pyproject.toml` / `requirements.txt` 等依赖声明文件。
- **运行时路径自适应**：应用通过 `getattr(sys, 'frozen', False)` 区分源码运行与打包后的 exe 运行，统一以 `BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))` 作为根目录，确保数据文件始终写入 exe 同级目录而非临时解压目录。
- **数据持久化约定**：SQLite 数据库 `差旅记录.db` 位于 `data/` 子目录；截图按日期归档到 `data/screenshots/YYYY-MM-DD/`；导出文件写入根级 `导出文件/` 目录。
- **打包产物即发行物**：仓库直接提交编译后的 `差旅记录.exe`，使用者无需安装 Python 环境即可运行。

## 4. 约定与约束

- 应用必须与 `启动差旅记录.bat` 放在同一目录下，因为批处理脚本使用相对路径 `start "" "差旅记录.exe"` 启动。
- 数据目录 `data/` 必须在应用可写位置存在（首次运行会自动创建），且需对 `screenshots/` 有写入权限。
- 由于未提供 `requirements.txt` 或依赖清单，打包时已将所需库（Pillow、openpyxl 等）静态嵌入 `差旅记录.exe`，因此该 exe 是平台绑定的 Windows 二进制。
- 不存在跨平台构建、容器化、自动化测试或 CI 流程；版本管理仅体现为人工维护的 `差旅记录.exe` 替换。
- 数据库表结构在 `db_init()` 中以 `CREATE TABLE IF NOT EXISTS` 形式定义，新增字段时需手动升级已有数据库（当前还包含针对旧 trips 表的迁移逻辑 `migrate_legacy_trips`）。

总结：这是一个极简的桌面应用分发方案——源码单文件 + 预打包 exe + 批处理启动脚本，没有任何自动化构建、依赖管理或持续集成配置。