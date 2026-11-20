"""批量提取文档正文为 txt。"""
import os

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_extract_text",
    "name": "提取文档文字",
    "desc": "把每个文档的正文导出为同名 txt。",
    "fields": [
        {"key": "with_tables", "label": "包含表格文字", "type": "checkbox", "default": True},
    ],
}


def _text_of(d, with_tables):
    parts = [p.text for p in d.paragraphs]
    if with_tables:
        for t in d.tables:
            for row in t.rows:
                parts.append("\t".join(c.text.strip() for c in row.cells))
    return "\n".join(parts)


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    with_tables = bool(params.get("with_tables", True))
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            text = _text_of(d, with_tables)
            dst = out_path(f, output_dir, suffix="", ext=".txt")
            with open(dst, "w", encoding="utf-8-sig") as fh:
                fh.write(text)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {len(text)} 字")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
