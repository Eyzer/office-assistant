"""删除或替换文档内图片。"""
import os

from docx.oxml.ns import qn

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_img_clean",
    "name": "删除文档图片",
    "desc": "删除文档内全部内嵌图片（保留文字与表格）。",
    "fields": [
        {"key": "keep_first", "label": "保留第一张（如公章）", "type": "checkbox", "default": False},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    keep_first = bool(params.get("keep_first"))
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            drawings = list(d.element.body.iter(qn("w:drawing")))
            removed = 0
            for idx, el in enumerate(drawings):
                if keep_first and idx == 0:
                    continue
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
                    removed += 1
            dst = out_path(f, output_dir, suffix="_去图")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 删除 {removed} 张图片")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
