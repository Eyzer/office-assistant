"""取消合并单元格并填充内容。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "cell_unmerge_fill",
    "name": "取消合并并填充",
    "desc": "取消全部合并单元格，把原左上角内容填充到整个区间。",
    "fields": [
        {"key": "start_row", "label": "从第几行开始", "type": "number", "default": 1},
        {"key": "fill_blank_only", "label": "只填空单元格", "type": "checkbox", "default": True},
    ],
}


def _process(path, output_dir, start_row, blank_only, ctx):
    wb = xu.load(path)
    total_merged = 0
    for ws in wb.worksheets:
        if blank_only:
            ranges = list(ws.merged_cells.ranges)
            for rng in ranges:
                top = ws.cell(row=rng.min_row, column=rng.min_col).value
                ws.unmerge_cells(str(rng))
                for r in range(rng.min_row, rng.max_row + 1):
                    if r < start_row:
                        continue
                    for c in range(rng.min_col, rng.max_col + 1):
                        cell = ws.cell(row=r, column=c)
                        if cell.value in (None, ""):
                            cell.value = top
                total_merged += 1
        else:
            total_merged += xu.unmerge_and_fill(ws, start_row)
        ctx.info(f"[{ws.title}] 处理合并区域完成")
    dst = out_path(path, output_dir, suffix="_取消合并")
    xu.save(wb, dst)
    return dst, total_merged


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    start_row = int(params.get("start_row") or 1)
    blank_only = bool(params.get("fill_blank_only", True))
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            dst, n = _process(f, output_dir, start_row, blank_only, ctx)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 取消 {n} 处合并")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
