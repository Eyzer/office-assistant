"""批量合并同值单元格。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "cell_merge_same",
    "name": "合并相同内容单元格",
    "desc": "把指定列中相邻的同值单元格合并成一个。",
    "fields": [
        {"key": "cols", "label": "作用列", "type": "text", "default": "A",
         "help": "列字母，多个用逗号分隔"},
        {"key": "start_row", "label": "起始行", "type": "number", "default": 2},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    from openpyxl.utils import column_index_from_string
    files = collect_files(inputs, EXCEL_EXT)
    cols = [c.strip().upper() for c in (params.get("cols") or "A").split(",") if c.strip()]
    start_row = int(params.get("start_row") or 2)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            n = 0
            for ws in wb.worksheets:
                for c in cols:
                    n += xu.merge_same_values(ws, column_index_from_string(c), start_row)
            dst = out_path(f, output_dir, suffix="_同值合并")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 合并 {n} 处")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
