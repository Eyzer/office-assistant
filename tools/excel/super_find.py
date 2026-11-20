"""超级查找：跨工作表/工作簿查找内容，输出命中清单。"""
import os

from utils import excel_utils as xu
from utils.encoding_utils import write_text
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "super_find",
    "name": "跨表跨簿查找",
    "desc": "在多个文件所有工作表里查找关键字，输出命中位置清单。",
    "fields": [
        {"key": "keyword", "label": "查找内容", "type": "text", "default": ""},
        {"key": "exact", "label": "完全匹配", "type": "checkbox", "default": False},
        {"key": "case_sensitive", "label": "区分大小写", "type": "checkbox", "default": False},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    kw = params.get("keyword") or ""
    exact = bool(params.get("exact"))
    cs = bool(params.get("case_sensitive"))
    if not kw:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["请填写查找内容"], "output_files": []}

    if not cs:
        kw_cmp = kw.lower()
    else:
        kw_cmp = kw

    lines = [f"查找: {kw}", ""]
    ok, fail = 0, 0
    hits = 0
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f, data_only=True)
            base = os.path.basename(f)
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        v = cell.value
                        if v is None:
                            continue
                        text = str(v)
                        cmp_text = text if cs else text.lower()
                        matched = (cmp_text == kw_cmp) if exact else (kw_cmp in cmp_text)
                        if matched:
                            lines.append(f"{base}\t{ws.title}\t{cell.coordinate}\t{text}")
                            hits += 1
            ok += 1
            ctx.info(f"{base}: 累计命中 {hits}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    lines.append("")
    lines.append(f"合计命中: {hits}")
    dst = out_path(files[0] if files else "查找结果.xlsx", output_dir,
                   suffix="_查找结果", ext=".txt")
    write_text(dst, "\n".join(lines))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
