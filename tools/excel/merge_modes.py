"""三种合并模式：结构相同（位置）、字段对齐（按表头名）、横向拼接。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "merge_modes",
    "name": "多模式合并",
    "desc": "结构相同按列位置合并；字段对齐按表头名匹配；横向拼接按列并排。",
    "fields": [
        {"key": "mode", "label": "合并模式", "type": "select", "default": "align",
         "options": [
             {"value": "same", "label": "结构相同（按位置）"},
             {"value": "align", "label": "字段对齐（按表头名）"},
             {"value": "hstack", "label": "横向拼接（并排）"},
         ]},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def _read(path, header_row):
    wb = xu.load(path, data_only=True)
    ws = wb.worksheets[0]
    header = xu.header_values(ws, header_row)
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v not in (None, "") for v in vals):
            rows.append(vals)
    return header, rows


def _merge_same(files, header_row, ctx):
    wb = openpyxl.Workbook()
    ws = wb.active
    row_out = 1
    first = True
    for f in files:
        header, rows = _read(f, header_row)
        if first:
            for c, v in enumerate(header, 1):
                ws.cell(row=1, column=c, value=v)
            row_out = 2
            first = False
        for row in rows:
            for c, v in enumerate(row, 1):
                ws.cell(row=row_out, column=c, value=v)
            row_out += 1
        ctx.info(f"并入 {os.path.basename(f)}")
    return wb


def _merge_align(files, header_row, ctx):
    wb = openpyxl.Workbook()
    ws = wb.active
    all_headers = []
    blocks = []
    for f in files:
        header, rows = _read(f, header_row)
        blocks.append((header, rows))
        for h in header:
            if h not in all_headers and h != "":
                all_headers.append(h)
    for c, h in enumerate(all_headers, 1):
        ws.cell(row=1, column=c, value=h)
    row_out = 2
    for header, rows in blocks:
        idx = {h: i for i, h in enumerate(header)}
        for row in rows:
            for c, h in enumerate(all_headers, 1):
                if h in idx and idx[h] < len(row):
                    ws.cell(row=row_out, column=c, value=row[idx[h]])
            row_out += 1
        ctx.info(f"对齐并入 {len(rows)} 行")
    return wb


def _merge_hstack(files, header_row, ctx):
    wb = openpyxl.Workbook()
    ws = wb.active
    col_off = 1
    max_rows = 0
    for f in files:
        header, rows = _read(f, header_row)
        for c, v in enumerate(header, col_off):
            ws.cell(row=1, column=c, value=v)
        for ri, row in enumerate(rows, 2):
            for c, v in enumerate(row, col_off):
                ws.cell(row=ri, column=c, value=v)
        col_off += max(len(header), 1)
        max_rows = max(max_rows, len(rows))
        ctx.info(f"横向拼接 {os.path.basename(f)}")
    return wb


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    if len(files) < 2:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["至少需要 2 个表格"], "output_files": []}
    mode = params.get("mode") or "align"
    header_row = int(params.get("header_row") or 1)
    files.sort(key=lambda p: os.path.basename(p))

    try:
        if mode == "same":
            wb = _merge_same(files, header_row, ctx)
        elif mode == "hstack":
            wb = _merge_hstack(files, header_row, ctx)
        else:
            wb = _merge_align(files, header_row, ctx)
        dst = out_path(files[0], output_dir, suffix=f"_合并_{mode}")
        xu.save(wb, dst)
        ctx.progress(len(files), len(files), "完成")
        return {"success_count": len(files), "fail_count": 0,
                "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
    except Exception as e:
        ctx.error(str(e))
        return {"success_count": 0, "fail_count": 1,
                "log_messages": ctx.task["log"].lines[-50:], "output_files": []}
