"""导出所有工作表名称清单为 CSV。"""
import os

from utils import excel_utils as xu
from utils.encoding_utils import write_text
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "excel_list_sheet",
    "name": "导出工作表清单",
    "desc": "把所有文件的工作表名称、行列数导出为一张 CSV 清单。",
    "fields": [],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    lines = ["文件名,工作表,行数,列数"]
    ok, fail = 0, 0
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f, read_only=True)
            for ws in wb.worksheets:
                lines.append(f"{os.path.basename(f)},{ws.title},{ws.max_row},{ws.max_column}")
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    anchor = files[0] if files else "清单.xlsx"
    dst = out_path(anchor, output_dir, suffix="_工作表清单", ext=".csv")
    write_text(dst, "\n".join(lines))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
