"""批量提取表格内的图片（zip 底层，保真）。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import batch_out_dir

META = {
    "id": "excel_extract_img",
    "name": "提取表格图片",
    "desc": "从 xlsx 包内导出全部图片，按文件分目录，不经过 openpyxl 以免丢图。",
    "fields": [
        {"key": "prefix", "label": "文件名前缀", "type": "text", "default": ""},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    prefix_opt = params.get("prefix") or ""
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            if os.path.splitext(f)[1].lower() == ".csv":
                ctx.warn(f"{os.path.basename(f)} 是 CSV，跳过")
                continue
            sub = batch_out_dir(f, output_dir, "_图片")
            stem = os.path.splitext(os.path.basename(f))[0]
            prefix = (prefix_opt or stem) + "_"
            saved = xu.extract_images(f, sub, prefix)
            outs.extend(saved)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 导出 {len(saved)} 张")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
