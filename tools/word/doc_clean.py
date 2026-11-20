"""文档清理：空段落、批注、超链接、多余空行。"""
import os

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_clean",
    "name": "批量清理文档",
    "desc": "删除空白段落、多余空行、批注与超链接。",
    "fields": [
        {"key": "blank_paras", "label": "删除空白段落", "type": "checkbox", "default": True},
        {"key": "max_blank", "label": "允许连续空行数", "type": "number", "default": 0,
         "help": "0=全部删除，1=保留一行"},
        {"key": "comments", "label": "删除批注", "type": "checkbox", "default": True},
        {"key": "hyperlinks", "label": "超链接转文本", "type": "checkbox", "default": True},
    ],
}


def _clean(path, output_dir, params, ctx):
    d = du.load(path)
    stats = []
    if params.get("blank_paras", True):
        n = du.remove_blank_paragraphs(d, int(params.get("max_blank") or 0))
        stats.append(f"空段落 {n}")
    if params.get("comments", True):
        n = du.remove_comments(d)
        stats.append(f"批注 {n}")
    if params.get("hyperlinks", True):
        n = du.remove_hyperlinks(d)
        stats.append(f"超链接 {n}")
    dst = out_path(path, output_dir, suffix="_清理")
    du.save(d, dst)
    ctx.info(f"{os.path.basename(path)}: {' '.join(stats) if stats else '无操作'}")
    return dst


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            outs.append(_clean(f, output_dir, params, ctx))
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
