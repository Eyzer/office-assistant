"""按行数拆表：每 N 行一块，拆成多个工作表或独立文件。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import batch_out_dir, out_path

META = {
    "id": "split_by_row",
    "name": "按行数拆分表格",
    "desc": "每 N 行数据拆成一块，可拆成多个工作表或独立文件。",
    "fields": [
        {"key": "rows", "label": "每份行数", "type": "number", "default": 1000},
        {"key": "mode", "label": "拆分到", "type": "select", "default": "sheets",
         "options": [{"value": "sheets", "label": "多个工作表"}, {"value": "files", "label": "多个独立文件"}]},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
        {"key": "repeat_header", "label": "每份重复表头", "type": "checkbox", "default": True},
    ],
}


def _split_one(path, output_dir, n, mode, header_row, repeat, ctx):
    wb = xu.load(path, data_only=True)
    ws = wb.worksheets[0]
    header = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
    data = [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
            for r in range(header_row + 1, ws.max_row + 1)]
    if not data:
        return []

    chunks = [data[i:i + n] for i in range(0, len(data), n)]
    outs = []
    stem, _ = os.path.splitext(os.path.basename(path))

    if mode == "files":
        base_dir = batch_out_dir(path, output_dir, "_按行拆分")
        for idx, chunk in enumerate(chunks, 1):
            wb2 = openpyxl.Workbook()
            w2 = wb2.active
            start = 1
            if repeat:
                for c, v in enumerate(header, 1):
                    w2.cell(row=1, column=c, value=v)
                start = 2
            for ri, row in enumerate(chunk, start):
                for c, v in enumerate(row, 1):
                    w2.cell(row=ri, column=c, value=v)
            dst = out_path(path, base_dir, suffix=f"_{idx}", ext=".xlsx")
            xu.save(wb2, dst)
            outs.append(dst)
        ctx.info(f"{stem} -> {len(chunks)} 个文件")
    else:
        wb2 = openpyxl.Workbook()
        wb2.remove(wb2.active)
        used = set()
        for idx, chunk in enumerate(chunks, 1):
            w2 = wb2.create_sheet(title=xu.safe_sheet_title(f"第{idx}块", used))
            start = 1
            if repeat:
                for c, v in enumerate(header, 1):
                    w2.cell(row=1, column=c, value=v)
                start = 2
            for ri, row in enumerate(chunk, start):
                for c, v in enumerate(row, 1):
                    w2.cell(row=ri, column=c, value=v)
        dst = out_path(path, output_dir, suffix="_按行拆分")
        xu.save(wb2, dst)
        outs.append(dst)
        ctx.info(f"{stem} -> {len(chunks)} 个工作表")
    return outs


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    n = max(1, int(params.get("rows") or 1000))
    mode = params.get("mode") or "sheets"
    header_row = int(params.get("header_row") or 1)
    repeat = bool(params.get("repeat_header", True))
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            outs.extend(_split_one(f, output_dir, n, mode, header_row, repeat, ctx))
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
