---
kind: external_dependency
name: openpyxl Excel 导出库
slug: openpyxl
category: external_dependency
category_hints:
    - sdk_real_api
scope:
    - '**'
source_files:
    - 差旅记录.pyw
---

用于将四类记录（行程、住宿、餐饮、交通）导出为带格式的 .xlsx 文件到 `导出文件/` 目录。每个导出函数创建一个 Workbook，写入表头、数据行和合计行，并应用统一的样式（蓝色表头填充 D9E1F2、橙色合计行 FCE4D6、细边框 BFBFBF、冻结首行）。文件名按当前时间戳命名，如 `差旅费用记录_YYYYMMDD_HHMMSS.xlsx`。