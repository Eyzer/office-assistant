"""按指定区域拆分工作表（每个区域一块）。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import batch_out_dir, out_path

META = {
    "id": "split_by_area",
    "name": "按区域拆分",
    "desc": "按给定行区间拆分，如 1-100,101-200；每段输出一份文件。",
    "fields": [
        {"key": "ranges", "label": "行区间", "type": "text", "default": "1-100,101-200",
         "help": "格式 起始行-结束行，逗号分隔"},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def _parse_ranges(raw):
    out = []
    for part in str(raw or "").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.append((int(a), int(b)))
        else:
            n = int(part)
            out.append((n, n))
    return out


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    ranges = _parse_ranges(params.get("ranges"))
    header_row = int(params.get("header_row") or 1)
    if not ranges:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["未提供有效行区间"], "output_files": []}

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f, data_only=True)
            ws = wb.worksheets[0]
            header = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
            sub = batch_out_dir(f, output_dir, "_区域拆分")
            for a, b in ranges:
                out_wb = openpyxl.Workbook()
                o = out_wb.active
                for c, v in enumerate(header, 1):
                    o.cell(row=1, column=c, value=v)
                ri = 2
                for r in range(max(a, header_row + 1), min(b, ws.max_row) + 1):
                    for c in range(1, ws.max_column + 1):
                        o.cell(row=ri, column=c, value=ws.cell(row=r, column=c).value)
                    ri += 1
                dst = out_path(f, sub, suffix=f"_{a}-{b}", ext=".xlsx")
                xu.save(out_wb, dst)
                outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 拆成 {len(ranges)} 份")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
