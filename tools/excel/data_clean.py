"""表格数据清洗：删空行空列、去重、批量替换单元格内容。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "data_clean",
    "name": "表格数据清洗",
    "desc": "删除空行空列、按整行去重、批量替换单元格文本。",
    "fields": [
        {"key": "del_blank_rows", "label": "删除空行", "type": "checkbox", "default": True},
        {"key": "del_blank_cols", "label": "删除空列", "type": "checkbox", "default": True},
        {"key": "dedup", "label": "整行去重", "type": "checkbox", "default": True},
        {"key": "find", "label": "查找内容", "type": "text", "default": ""},
        {"key": "replace", "label": "替换为", "type": "text", "default": ""},
    ],
}


def _del_blank_rows(ws):
    removed = 0
    for r in range(ws.max_row, 0, -1):
        if xu.is_blank_row(ws, r):
            ws.delete_rows(r, 1)
            removed += 1
    return removed


def _del_blank_cols(ws):
    removed = 0
    for c in range(ws.max_column, 0, -1):
        if xu.is_blank_col(ws, c):
            ws.delete_cols(c, 1)
            removed += 1
    return removed


def _dedup(ws):
    seen, remove = set(), []
    for r in range(1, ws.max_row + 1):
        key = tuple(str(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1))
        if key in seen:
            remove.append(r)
        else:
            seen.add(key)
    for r in reversed(remove):
        ws.delete_rows(r, 1)
    return len(remove)


def _replace(ws, find, repl):
    n = 0
    if not find:
        return 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and find in cell.value:
                cell.value = cell.value.replace(find, repl)
                n += 1
    return n


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    do_rows = bool(params.get("del_blank_rows", True))
    do_cols = bool(params.get("del_blank_cols", True))
    do_dedup = bool(params.get("dedup", True))
    find = params.get("find") or ""
    repl = params.get("replace") or ""
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            stats = []
            for ws in wb.worksheets:
                if do_rows:
                    stats.append(f"空行{_del_blank_rows(ws)}")
                if do_cols:
                    stats.append(f"空列{_del_blank_cols(ws)}")
                if do_dedup:
                    stats.append(f"重复{_dedup(ws)}")
                if find:
                    stats.append(f"替换{_replace(ws, find, repl)}")
            dst = out_path(f, output_dir, suffix="_清洗")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {' '.join(stats)}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
