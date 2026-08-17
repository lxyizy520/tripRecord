---
kind: external_dependency
name: Pillow 图片处理库
slug: pillow-pil
category: external_dependency
category_hints:
    - sdk_real_api
scope:
    - '**'
source_files:
    - 差旅记录.pyw
---

用于打开用户选择的图片文件并在弹窗中生成缩略图进行预览（Image.open + Image.thumbnail + ImageTk.PhotoImage）。支持 .png/.jpg/.jpeg/.bmp/.gif/.webp 格式，不支持的图片会捕获异常并提示。截图文件通过 shutil.copy2 复制到 data/screenshots/日期/ 目录下，文件名格式为 tag_HHMMSSfff.ext。