"""二维表与一维表互转。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "pivot_flatten",
    "name": "二维表转一维",
    "desc": "二维交叉表压平成一维明细（行标签、列标签、值），或反向还原。",
    "fields": [
        {"key": "mode", "label": "转换方向", "type": "select", "default": "to_long",
         "options": [{"value": "to_long", "label": "二维 → 一维"}, {"value": "to_wide", "label": "一维 → 二维"}]},
        {"key": "row_label", "label": "行标签表头", "type": "text", "default": ""},
        {"key": "col_header_row", "label": "列标签所在行", "type": "number", "default": 1},
    ],
}


def _to_long(ws, col_header_row):
    headers = [ws.cell(row=col_header_row, column=c).value for c in range(1, ws.max_column + 1)]
    out = openpyxl.Workbook()
    o = out.active
    o.append(["行标签", "列标签", "值"])
    for r in range(col_header_row + 1, ws.max_row + 1):
        row_label = ws.cell(row=r, column=1).value
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v in (None, ""):
                continue
            o.append([row_label, headers[c - 1], v])
    return out


def _to_wide(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return openpyxl.Workbook()
    header = rows[0]
    row_keys, col_keys = [], []
    data = {}
    for r in rows[1:]:
        if len(r) < 3:
            continue
        rk, ck, v = r[0], r[1], r[2]
        if rk not in row_keys:
            row_keys.append(rk)
        if ck not in col_keys:
            col_keys.append(ck)
        data[(rk, ck)] = v
    out = openpyxl.Workbook()
    o = out.active
    o.cell(row=1, column=1, value=header[0] if header else "行标签")
    for ci, ck in enumerate(col_keys, 2):
        o.cell(row=1, column=ci, value=ck)
    for ri, rk in enumerate(row_keys, 2):
        o.cell(row=ri, column=1, value=rk)
        for ci, ck in enumerate(col_keys, 2):
            o.cell(row=ri, column=ci, value=data.get((rk, ck)))
    return out


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    mode = params.get("mode") or "to_long"
    col_header_row = int(params.get("col_header_row") or 1)
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f, data_only=True)
            ws = wb.worksheets[0]
            out = _to_long(ws, col_header_row) if mode == "to_long" else _to_wide(ws)
            dst = out_path(f, output_dir, suffix=f"_{'一维' if mode == 'to_long' else '二维'}")
            xu.save(out, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)} 转换完成")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
