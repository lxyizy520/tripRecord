---
kind: external_dependency
name: PyInstaller 打包分发
slug: pyinstaller
category: external_dependency
category_hints:
    - migration_status
scope:
    - '**'
source_files:
    - 差旅记录.exe
    - 启动差旅记录.bat
    - 差旅记录.pyw
---

应用通过 PyInstaller 将 Python 源码打包为独立的 `差旅记录.exe` 可执行文件。打包后程序通过 `sys.executable` 定位自身所在目录作为 BASE_DIR，确保数据文件（data/）、截图目录、导出目录始终位于 exe 同级位置，避免写入临时解压目录。提供 `启动差旅记录.bat` 作为一键启动入口。