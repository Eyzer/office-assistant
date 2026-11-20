"""摊平目录：把多层文件夹里的所有文件抽到一层。"""
import os
import shutil

from utils.file_utils import collect_files
from utils.path_utils import ensure_dir, unique_path

META = {
    "id": "flatten_dir",
    "name": "摊平目录",
    "desc": "把文件夹里所有层级的文件抽到同一层，重名自动加序号。",
    "fields": [
        {"key": "keep_prefix", "label": "保留上级目录名", "type": "checkbox", "default": False,
         "help": "勾选后文件名前加所在文件夹名，避免混淆"},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs)
    ensure_dir(output_dir)
    prefix_on = bool(params.get("keep_prefix"))
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, src in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            base = os.path.basename(src)
            if prefix_on:
                parent = os.path.basename(os.path.dirname(src))
                base = f"{parent}_{base}"
            dst = unique_path(os.path.join(output_dir, base))
            shutil.copy2(src, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{src} -> {os.path.basename(dst)}")
        except Exception as e:
            fail += 1
            ctx.error(f"{src}: {e}")
        ctx.progress(i, total, os.path.basename(src))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
