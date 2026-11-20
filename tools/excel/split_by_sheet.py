"""按工作表拆成独立文件，文件名自动带表名。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import batch_out_dir, out_path, safe_name

META = {
    "id": "split_by_sheet",
    "name": "按工作表拆分",
    "desc": "一个工作簿里的每张工作表拆成独立文件，文件名带表名。",
    "fields": [
        {"key": "keep_values", "label": "公式转数值", "type": "checkbox", "default": False},
    ],
}


def _split_one(path, output_dir, keep_values, ctx):
    wb = xu.load(path, data_only=keep_values)
    base_dir = batch_out_dir(path, output_dir, "_按表拆分")
    outs = []
    for ws in wb.worksheets:
        wb2 = openpyxl.Workbook()
        w2 = wb2.active
        w2.title = ws.title
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                src = ws.cell(row=r, column=c)
                dst = w2.cell(row=r, column=c, value=src.value)
                if src.has_style:
                    dst.font = src.font.copy()
                    dst.fill = src.fill.copy()
                    dst.border = src.border.copy()
                    dst.alignment = src.alignment.copy()
                    dst.number_format = src.number_format
        xu.copy_dimensions(ws, w2)
        dst = out_path(path, base_dir, suffix=f"_{safe_name(ws.title)}", ext=".xlsx")
        xu.save(wb2, dst)
        outs.append(dst)
        ctx.info(f"{os.path.basename(path)} [{ws.title}] -> {os.path.basename(dst)}")
    return outs


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    keep_values = bool(params.get("keep_values"))
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            outs.extend(_split_one(f, output_dir, keep_values, ctx))
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
