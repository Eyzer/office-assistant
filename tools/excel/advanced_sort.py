"""高级排序：多关键字排序、随机排序。"""
import os
import random

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "advanced_sort",
    "name": "高级排序",
    "desc": "按一列或多列排序（升/降），支持随机打乱。",
    "fields": [
        {"key": "cols", "label": "排序列", "type": "text", "default": "A",
         "help": "多列用逗号分隔，如 A,B"},
        {"key": "desc", "label": "降序", "type": "checkbox", "default": False},
        {"key": "random", "label": "随机排序", "type": "checkbox", "default": False},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def _key(v):
    if v is None:
        return (0, "")
    if isinstance(v, (int, float)):
        return (1, v)
    return (2, str(v))


def run_task(inputs, output_dir, params, ctx):
    from openpyxl.utils import column_index_from_string
    files = collect_files(inputs, EXCEL_EXT)
    cols = [c.strip().upper() for c in (params.get("cols") or "A").split(",") if c.strip()]
    desc = bool(params.get("desc"))
    do_random = bool(params.get("random"))
    header_row = int(params.get("header_row") or 1)
    idxs = [column_index_from_string(c) - 1 for c in cols]

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            for ws in wb.worksheets:
                rows = [[c.value for c in row] for row in ws.iter_rows(min_row=header_row + 1)]
                if do_random:
                    random.shuffle(rows)
                else:
                    rows.sort(key=lambda r: tuple(_key(r[j]) for j in idxs if j < len(r)),
                              reverse=desc)
                for ri, row in enumerate(rows, header_row + 1):
                    for ci, v in enumerate(row, 1):
                        ws.cell(row=ri, column=ci, value=v)
            dst = out_path(f, output_dir, suffix="_排序")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {'随机' if do_random else '按' + ','.join(cols)} 排序")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
