"""随机抽样：从表格中随机抽取若干行。"""
import os
import random

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "random_sample",
    "name": "随机抽样",
    "desc": "从每张表随机抽取 N 行或 N%，保留表头。",
    "fields": [
        {"key": "count", "label": "抽取行数", "type": "number", "default": 10},
        {"key": "by_percent", "label": "按百分比抽取", "type": "checkbox", "default": False},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    count = max(1, int(params.get("count") or 10))
    by_pct = bool(params.get("by_percent"))
    header_row = int(params.get("header_row") or 1)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f, data_only=True)
            out_wb = None
            for ws in wb.worksheets:
                header = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
                rows = [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                        for r in range(header_row + 1, ws.max_row + 1)]
                if not rows:
                    continue
                n = max(1, int(len(rows) * count / 100)) if by_pct else min(count, len(rows))
                picked = random.sample(rows, n)
                if out_wb is None:
                    import openpyxl
                    out_wb = openpyxl.Workbook()
                    out_wb.remove(out_wb.active)
                o = out_wb.create_sheet(title=xu.safe_sheet_title(ws.title))
                for c, v in enumerate(header, 1):
                    o.cell(row=1, column=c, value=v)
                for ri, row in enumerate(picked, 2):
                    for c, v in enumerate(row, 1):
                        o.cell(row=ri, column=c, value=v)
            if out_wb is None:
                raise ValueError("没有可抽样的数据")
            dst = out_path(f, output_dir, suffix="_抽样")
            xu.save(out_wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 抽样完成")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
