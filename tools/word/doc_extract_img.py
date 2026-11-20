"""批量提取文档内图片。"""
import os

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import batch_out_dir

META = {
    "id": "doc_extract_img",
    "name": "提取文档图片",
    "desc": "把文档里的图片全部导出到独立文件夹，按文档名分目录。",
    "fields": [
        {"key": "prefix", "label": "文件名前缀", "type": "text", "default": "",
         "help": "留空用文档名"},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    prefix_opt = params.get("prefix") or ""
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            sub = batch_out_dir(f, output_dir, "_图片")
            stem = os.path.splitext(os.path.basename(f))[0]
            prefix = (prefix_opt or stem) + "_"
            saved = du.extract_images(f, sub, prefix)
            outs.extend(saved)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 导出 {len(saved)} 张")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
