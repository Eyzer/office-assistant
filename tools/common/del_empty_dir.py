"""删除空文件夹（只删空目录，不动文件）。"""
import os

from utils.path_utils import ensure_dir

META = {
    "id": "del_empty_dir",
    "name": "删除空文件夹",
    "desc": "从指定目录向下清理空文件夹，含文件的目录一律不动。",
    "fields": [
        {"key": "dry_run", "label": "只统计不删除", "type": "checkbox", "default": True},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    dry = bool(params.get("dry_run", True))
    roots = [p for p in inputs if os.path.isdir(p)]
    if not roots:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["请添加文件夹"], "output_files": []}

    total_removed = 0
    for root in roots:
        if ctx.cancelled:
            break
        # 自底向上，避免父目录先被判断
        for cur, dirs, files in os.walk(root, topdown=False):
            if cur == root:
                continue
            if not dirs and not files:
                total_removed += 1
                if dry:
                    msg = f"[预览] 将删除空目录 {cur}"
                else:
                    try:
                        os.rmdir(cur)
                        msg = f"已删除空目录 {cur}"
                    except OSError as e:
                        msg = f"删除失败 {cur}: {e}"
                level = ctx.info if dry else ctx.ok
                level(msg)
        ctx.progress(1, 1, os.path.basename(root))

    return {"success_count": total_removed, "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": []}
