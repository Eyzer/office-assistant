"""按列字段拆表：同一列的不同值拆成多个工作表或独立文件。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import batch_out_dir, out_path

META = {
    "id": "split_by_col",
    "name": "按列拆分表格",
    "desc": "按指定列的取值拆分，可拆成多个工作表或独立文件。",
    "fields": [
        {"key": "col", "label": "拆分依据列", "type": "text", "default": "A",
         "help": "填列字母（A、B）或表头文字"},
        {"key": "mode", "label": "拆分到", "type": "select", "default": "sheets",
         "options": [{"value": "sheets", "label": "多个工作表"}, {"value": "files", "label": "多个独立文件"}]},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def _col_index(ws, col, header_row):
    col = str(col).strip()
    if col.isalpha():
        return openpyxl.utils.column_index_from_string(col.upper())
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(row=header_row, column=c).value).strip() == col:
            return c
    return 1


def _split_one(path, output_dir, col, mode, header_row, ctx):
    wb = xu.load(path, data_only=True)
    ws = wb.worksheets[0]
    ci = _col_index(ws, col, header_row)
    header = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]

    groups = {}
    for r in range(header_row + 1, ws.max_row + 1):
        key = ws.cell(row=r, column=ci).value
        key = str(key).strip() if key not in (None, "") else "空值"
        groups.setdefault(key, []).append(
            [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)])

    outs = []
    if mode == "files":
        base_dir = batch_out_dir(path, output_dir, "_拆分")
        for key, rows in groups.items():
            wb2 = openpyxl.Workbook()
            w2 = wb2.active
            w2.title = "数据"
            for c, v in enumerate(header, 1):
                w2.cell(row=1, column=c, value=v)
            for ri, row in enumerate(rows, 2):
                for c, v in enumerate(row, 1):
                    w2.cell(row=ri, column=c, value=v)
            stem, _ = os.path.splitext(os.path.basename(path))
            dst = out_path(path, base_dir, suffix=f"_{key}", ext=".xlsx")
            xu.save(wb2, dst)
            outs.append(dst)
            ctx.info(f"{stem} -> {os.path.basename(dst)}")
    else:
        wb2 = openpyxl.Workbook()
        wb2.remove(wb2.active)
        used = set()
        for key, rows in groups.items():
            title = xu.safe_sheet_title(key, used)
            w2 = wb2.create_sheet(title=title)
            for c, v in enumerate(header, 1):
                w2.cell(row=1, column=c, value=v)
            for ri, row in enumerate(rows, 2):
                for c, v in enumerate(row, 1):
                    w2.cell(row=ri, column=c, value=v)
        dst = out_path(path, output_dir, suffix="_按列拆分")
        xu.save(wb2, dst)
        outs.append(dst)
        ctx.info(f"{os.path.basename(path)} -> {len(groups)} 个工作表")
    return outs


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    col = params.get("col") or "A"
    mode = params.get("mode") or "sheets"
    header_row = int(params.get("header_row") or 1)
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            outs.extend(_split_one(f, output_dir, col, mode, header_row, ctx))
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
