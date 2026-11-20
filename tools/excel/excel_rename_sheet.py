"""批量重命名工作表。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "excel_rename_sheet",
    "name": "批量重命名工作表",
    "desc": "按统一规则重命名每张工作表，如 表1、部门-数据。",
    "fields": [
        {"key": "rule", "label": "命名规则", "type": "text", "default": "表{index}",
         "help": "可用 {index} 序号、{name} 原名、{file} 文件名"},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    rule = params.get("rule") or "表{index}"
    stem = ""
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            stem = os.path.splitext(os.path.basename(f))[0]
            wb = xu.load(f)
            used = set()
            for idx, ws in enumerate(wb.worksheets, 1):
                title = xu.safe_sheet_title(
                    rule.format(index=idx, name=ws.title, file=stem), used)
                ws.title = title
            dst = out_path(f, output_dir, suffix="_改名")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {len(wb.worksheets)} 张表已改名")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
