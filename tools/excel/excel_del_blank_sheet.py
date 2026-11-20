"""批量删除空白工作表。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "excel_del_blank_sheet",
    "name": "删除空白工作表",
    "desc": "删除没有任何数据的工作表。",
    "fields": [],
}


def _blank(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.value not in (None, ""):
                return False
    return True


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            removed = [ws.title for ws in list(wb.worksheets) if _blank(ws)]
            if len(removed) >= len(wb.worksheets):
                ctx.warn(f"{os.path.basename(f)}: 全部为空，跳过")
                continue
            for title in removed:
                wb.remove(wb[title])
            dst = out_path(f, output_dir, suffix="_删空表")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 删除 {len(removed)} 个空表")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
