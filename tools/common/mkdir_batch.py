"""按名单批量创建文件夹（名单可直接填，也可来自 txt / xlsx）。"""
import os
import re

from utils.encoding_utils import read_text
from utils.path_utils import ensure_dir, safe_name

META = {
    "id": "mkdir_batch",
    "cat": "文件整理",
    "name": "批量创建文件夹",
    "desc": "按填写的名单（逗号/换行分隔）批量建文件夹，也支持从 txt 每行或表格第一列读取名称。",
    "fields": [
        {"key": "names", "label": "文件夹名", "type": "text", "default": "",
         "help": "逗号或换行分隔，如 一月,二月,三月"},
        {"key": "base", "label": "建在哪个目录", "type": "text", "default": "",
         "help": "留空=输出目录"},
        {"key": "prefix", "label": "统一前缀", "type": "text", "default": ""},
        {"key": "numbered", "label": "自动加序号", "type": "checkbox", "default": False,
         "help": "如 01_销售部"},
    ],
}


def _names_from_text(text):
    return [x.strip() for x in re.split(r"[,，;；\n\r\t]+", str(text or "")) if x.strip()]


def _names_from_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".csv", ".md"):
        return [l.strip() for l in read_text(path).splitlines() if l.strip()]
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        out = []
        for row in ws.iter_rows(values_only=True, max_col=1):
            v = row[0] if row else None
            if v not in (None, ""):
                out.append(str(v).strip())
        wb.close()
        return out
    return []


def run_task(inputs, output_dir, params, ctx):
    names = _names_from_text(params.get("names"))
    files = [p for p in inputs if os.path.isfile(p)]
    for f in files:
        names.extend(_names_from_file(f))
    names = [n for n in dict.fromkeys(names) if n]
    if not names:
        ctx.error("没有可用的文件夹名：请填写名单，或加入一个 txt/表格")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    base = params.get("base") or output_dir
    prefix = params.get("prefix") or ""
    numbered = bool(params.get("numbered"))
    ensure_dir(base)

    ok, fail, outs = 0, 0, []
    total = len(names)
    for i, n in enumerate(names, 1):
        if ctx.cancelled:
            break
        try:
            label = (("%02d_" % i) if numbered else "") + prefix + safe_name(n)
            p = os.path.join(base, label)
            os.makedirs(p, exist_ok=True)
            outs.append(p)
            ok += 1
        except OSError as exc:
            fail += 1
            ctx.error("%s: %s" % (n, exc))
        ctx.progress(i, total, n)
    ctx.info("在 %s 建了 %d 个文件夹" % (base, ok))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
