"""工具分类索引：给侧边栏做二级分组。

工具多起来以后平铺一列找不到东西，这里给每个工具一个二级分类，
前端按分类渲染小标题，并支持搜索过滤。新增工具记得在这里登记。
"""

# 每个大类下的二级分类顺序
ORDER = {
    "common": ["文件整理", "批量改名", "压缩与编码", "图片处理", "格式转换"],
    "excel": ["拆分与合并", "数据整理", "单元格与区域", "文本处理", "数值与计算",
              "格式与样式", "工作表与工作簿", "查找与对比", "图片与条码",
              "批注", "导入与导出"],
    "word": ["内容处理", "格式排版", "提取与转换", "文档结构"],
}

CAT = {
    # ---------- 通用文件 ----------
    "batch_copy_move": "文件整理",
    "batch_sort": "文件整理",
    "flatten_dir": "文件整理",
    "del_empty_dir": "文件整理",
    "mkdir_batch": "文件整理",
    "file_list": "文件整理",
    "dup_file": "文件整理",
    "batch_rename": "批量改名",
    "batch_zip": "压缩与编码",
    "encoding_ops": "压缩与编码",
    "img_convert": "图片处理",
    "ppt2pdf": "格式转换",

    # ---------- 表格 ----------
    "split_by_col": "拆分与合并",
    "split_by_row": "拆分与合并",
    "split_by_sheet": "拆分与合并",
    "split_by_area": "拆分与合并",
    "split_by_keyword": "拆分与合并",
    "merge_all": "拆分与合并",
    "merge_modes": "拆分与合并",
    "merge_folder": "拆分与合并",

    "pivot_flatten": "数据整理",
    "transpose_ops": "数据整理",
    "hierarchy_ops": "数据整理",
    "subtotal_ops": "数据整理",
    "advanced_sort": "数据整理",
    "random_sample": "数据整理",
    "data_clean": "数据整理",
    "data_dup": "数据整理",
    "row_col_ops": "数据整理",

    "cell_join": "单元格与区域",
    "cell_split": "单元格与区域",
    "cell_merge_same": "单元格与区域",
    "cell_unmerge_fill": "单元格与区域",
    "insert_alt": "单元格与区域",
    "sum_ops": "单元格与区域",
    "hyperlink_batch": "单元格与区域",
    "style_ops": "格式与样式",

    "text_ops": "文本处理",
    "text_extract": "文本处理",
    "text_convert": "文本处理",

    "num_format": "数值与计算",
    "round_ops": "数值与计算",
    "formula_ops": "数值与计算",
    "formula_to_value": "数值与计算",
    "name_data_ops": "数值与计算",

    "format_standard": "格式与样式",
    "format_convert": "格式与样式",
    "format_copy_across": "格式与样式",
    "style_ops": "格式与样式",

    "sheet_batch": "工作表与工作簿",
    "sheet_ops": "工作表与工作簿",
    "excel_list_sheet": "工作表与工作簿",
    "excel_rename_sheet": "工作表与工作簿",
    "excel_del_blank_sheet": "工作表与工作簿",
    "book_rename": "工作表与工作簿",

    "workbook_compare": "查找与对比",
    "compare_ops": "查找与对比",
    "super_find": "查找与对比",

    "qrcode_batch": "图片与条码",
    "barcode_gen": "图片与条码",
    "excel_extract_img": "图片与条码",
    "excel_import_img": "图片与条码",
    "img_ops": "图片与条码",

    "comment_ops": "批注",

    "excel_export_pdf": "导入与导出",
    "json_ops": "导入与导出",
    "textfile_ops": "导入与导出",
    "vcf_dbf": "导入与导出",
    "pdf_ops": "导入与导出",

    # ---------- 文字文档 ----------
    "doc_clean": "内容处理",
    "doc_text_replace": "内容处理",
    "doc_revision": "内容处理",
    "doc_metadata": "内容处理",
    "doc_mail_merge": "内容处理",

    "doc_page_style": "格式排版",
    "doc_content_style": "格式排版",
    "doc_header_footer": "格式排版",
    "doc_img_layout": "格式排版",
    "doc_table_convert": "格式排版",
    "doc_watermark": "格式排版",

    "doc_extract_text": "提取与转换",
    "doc_info_extract": "提取与转换",
    "doc_extract_img": "提取与转换",
    "doc_table_extract": "提取与转换",
    "doc_export_pdf": "提取与转换",
    "doc_compress": "提取与转换",
    "doc_stat": "提取与转换",

    "doc_merge": "文档结构",
    "doc_split": "文档结构",
    "doc_page_op": "文档结构",
    "doc_protect": "文档结构",
    "doc_compare": "文档结构",
    "doc_img_clean": "文档结构",
    "doc_toc": "文档结构",
}


def cat_of(tool_id, group=""):
    return CAT.get(tool_id, "其他")


def sort_key(item):
    """按 大类 → 二级分类顺序 → 工具名 排序。"""
    g = item["group"]
    order = ORDER.get(g, [])
    c = item.get("cat") or "其他"
    idx = order.index(c) if c in order else len(order)
    return (list(ORDER.keys()).index(g) if g in ORDER else 9, idx, item.get("name", ""))
