"""数据查重：标出或导出重复项。"""
import os

from openpyxl.styles import PatternFill

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "data_dup",
    "name": "数据查重",
    "desc": "按指定列查重，可标红或只保留重复清单。",
    "fields": [
        {"key": "cols", "label": "查重列", "type": "text", "default": "A"},
        {"key": "mark", "label": "标红重复行", "type": "checkbox", "default": True},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    from openpyxl.utils import column_index_from_string
    files = collect_files(inputs, EXCEL_EXT)
    cols = [c.strip().upper() for c in (params.get("cols") or "A").split(",") if c.strip()]
    mark = bool(params.get("mark", True))
    header_row = int(params.get("header_row") or 1)
    idxs = [column_index_from_string(c) for c in cols]
    red = PatternFill("solid", fgColor="FFC7CE")

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            cnt = 0
            for ws in wb.worksheets:
                seen = {}
                for r in range(header_row + 1, ws.max_row + 1):
                    key = tuple(str(ws.cell(row=r, column=c).value) for c in idxs)
                    if key in seen:
                        cnt += 1
                        if mark:
                            for c in range(1, ws.max_column + 1):
                                ws.cell(row=r, column=c).fill = red
                    else:
                        seen[key] = r
            dst = out_path(f, output_dir, suffix="_查重")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 重复 {cnt} 行")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
