"""批量查找替换文档文本。"""
import os

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_text_replace",
    "name": "批量查找替换",
    "desc": "多组关键词批量替换，可选区分大小写、整词匹配。",
    "fields": [
        {"key": "pairs", "label": "替换规则", "type": "text", "default": "",
         "help": "每行一条，格式：原文=>新文"},
        {"key": "case_sensitive", "label": "区分大小写", "type": "checkbox", "default": True},
        {"key": "whole_word", "label": "整词匹配", "type": "checkbox", "default": False},
    ],
}


def _parse_pairs(raw):
    pairs = []
    for line in (raw or "").splitlines():
        line = line.strip()
        if not line or "=>" not in line:
            continue
        a, b = line.split("=>", 1)
        if a.strip():
            pairs.append((a.strip(), b.strip()))
    return pairs


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    pairs = _parse_pairs(params.get("pairs"))
    case_sensitive = bool(params.get("case_sensitive", True))
    whole_word = bool(params.get("whole_word", False))
    ok, fail, outs = 0, 0, []
    total = len(files)

    if not pairs:
        ctx.warn("未提供替换规则，应写成 原文=>新文")

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            n = 0
            for old, new in pairs:
                n += du.replace_text(d, old, new, case_sensitive, whole_word)
            dst = out_path(f, output_dir, suffix="_替换")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 替换 {n} 处")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
