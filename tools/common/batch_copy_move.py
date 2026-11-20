"""批量复制或移动文件。"""
import os
import shutil

from utils.file_utils import collect_files
from utils.path_utils import ensure_dir, unique_path

META = {
    "id": "batch_copy_move",
    "name": "批量复制/移动",
    "desc": "把选中文件复制或移动到目标目录，重名自动加序号。",
    "fields": [
        {"key": "move", "label": "移动（不勾=复制）", "type": "checkbox", "default": False},
        {"key": "target", "label": "目标目录", "type": "text", "default": ""},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs)
    move = bool(params.get("move"))
    target = (params.get("target") or "").strip() or output_dir
    ensure_dir(target)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            dst = unique_path(os.path.join(target, os.path.basename(f)))
            if move:
                shutil.move(f, dst)
            else:
                shutil.copy2(f, dst)
            outs.append(dst)
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
