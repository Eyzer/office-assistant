"""按规则批量重命名文件（复制到输出目录）。"""
import os
import re
import shutil

from utils.file_utils import collect_files
from utils.path_utils import ensure_dir, unique_path

META = {
    "id": "book_rename",
    "name": "按规则批量改名",
    "desc": "用正则替换或加前后缀批量改名，复制到输出目录。",
    "fields": [
        {"key": "find", "label": "正则查找", "type": "text", "default": ""},
        {"key": "repl", "label": "替换为", "type": "text", "default": ""},
        {"key": "prefix", "label": "加前缀", "type": "text", "default": ""},
        {"key": "suffix", "label": "加后缀", "type": "text", "default": ""},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs)
    ensure_dir(output_dir)
    find = params.get("find") or ""
    repl = params.get("repl") or ""
    prefix = params.get("prefix") or ""
    suffix = params.get("suffix") or ""

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            stem, ext = os.path.splitext(os.path.basename(f))
            new_stem = re.sub(find, repl, stem) if find else stem
            new_name = f"{prefix}{new_stem}{suffix}{ext}"
            dst = unique_path(os.path.join(output_dir, new_name))
            shutil.copy2(f, dst)
            outs.append(dst)
            ok += 1
            if new_name != os.path.basename(f):
                ctx.info(f"{os.path.basename(f)} -> {os.path.basename(dst)}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
