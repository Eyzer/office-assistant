"""按规则把文件分拣到子文件夹。"""
import os
import shutil

from utils.file_utils import collect_files
from utils.path_utils import ensure_dir, safe_name, unique_path

META = {
    "id": "batch_sort",
    "name": "按规则分拣文件",
    "desc": "按扩展名、文件名关键字或修改月份分拣到不同子文件夹。",
    "fields": [
        {"key": "by", "label": "分拣依据", "type": "select", "default": "ext",
         "options": [{"value": "ext", "label": "按扩展名"}, {"value": "keyword", "label": "按关键字"},
                     {"value": "month", "label": "按修改月份"}]},
        {"key": "keywords", "label": "关键字", "type": "text", "default": "",
         "help": "按关键字时使用，逗号分隔，如 合同,报告"},
    ],
}


def _bucket(f, by, keywords):
    if by == "ext":
        return os.path.splitext(f)[1].lstrip(".").lower() or "无扩展名"
    if by == "month":
        import datetime
        return datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m")
    for k in keywords:
        if k and k in os.path.basename(f):
            return k
    return "其他"


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs)
    by = params.get("by") or "ext"
    keywords = [k.strip() for k in (params.get("keywords") or "").split(",") if k.strip()]
    ensure_dir(output_dir)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            bucket = safe_name(_bucket(f, by, keywords))
            sub = ensure_dir(os.path.join(output_dir, bucket))
            dst = unique_path(os.path.join(sub, os.path.basename(f)))
            shutil.copy2(f, dst)
            outs.append(dst)
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
