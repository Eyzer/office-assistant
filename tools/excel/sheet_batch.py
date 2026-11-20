"""工作表批量操作：重命名、删除空表、保护、排序。"""
import os
import re

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "sheet_batch",
    "name": "工作表批量管理",
    "desc": "批量重命名、删除空白表、保护或取消保护工作表。",
    "fields": [
        {"key": "rename", "label": "重命名规则", "type": "text", "default": "",
         "help": "如 {index}_{name}，留空不改名"},
        {"key": "del_blank", "label": "删除空白工作表", "type": "checkbox", "default": False},
        {"key": "protect", "label": "保护工作表", "type": "checkbox", "default": False},
        {"key": "password", "label": "保护密码", "type": "text", "default": ""},
    ],
}


def _is_blank(ws):
    return ws.max_row <= 1 and ws.max_column <= 1 and all(
        c.value in (None, "") for row in ws.iter_rows() for c in row)


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    rule = (params.get("rename") or "").strip()
    del_blank = bool(params.get("del_blank"))
    protect = bool(params.get("protect"))
    pwd = params.get("password") or ""

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            used = set()
            for idx, ws in enumerate(list(wb.worksheets), 1):
                if del_blank and _is_blank(ws):
                    wb.remove(ws)
                    ctx.info(f"删除空白表 {ws.title}")
                    continue
                if rule:
                    new_title = xu.safe_sheet_title(
                        rule.format(index=idx, name=ws.title), used)
                    ws.title = new_title
                else:
                    used.add(ws.title.lower())
                if protect:
                    ws.protection.sheet = True
                    if pwd:
                        ws.protection.password = pwd
            dst = out_path(f, output_dir, suffix="_工作表")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 工作表处理完成")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
