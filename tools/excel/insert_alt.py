"""隔行插入行或列、隔行换色。"""
import os

from openpyxl.styles import PatternFill

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "insert_alt",
    "name": "隔行插入/隔行换色",
    "desc": "每隔 N 行插入一行空行，或按奇偶行填充底色。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "color",
         "options": [{"value": "color", "label": "隔行换色"}, {"value": "insert", "label": "隔行插入空行"}]},
        {"key": "step", "label": "间隔行数", "type": "number", "default": 2},
        {"key": "color", "label": "底色", "type": "text", "default": "F2F6FC"},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    mode = params.get("mode") or "color"
    step = max(2, int(params.get("step") or 2))
    color = params.get("color") or "F2F6FC"
    fill = PatternFill("solid", fgColor=color)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            for ws in wb.worksheets:
                if mode == "color":
                    for r in range(1, ws.max_row + 1):
                        if (r - 1) % step == 1:
                            for c in range(1, ws.max_column + 1):
                                ws.cell(row=r, column=c).fill = fill
                else:
                    for r in range(ws.max_row, 1, -step):
                        ws.insert_rows(r)
            dst = out_path(f, output_dir, suffix="_隔行")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {mode} 完成")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
