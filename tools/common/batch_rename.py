"""批量改名：复制到输出目录并按模板重命名，不动源文件。"""
import os
import shutil
from datetime import date

from utils.file_utils import collect_files
from utils.path_utils import ensure_dir, safe_name, unique_path

META = {
    "id": "batch_rename",
    "name": "批量改名",
    "desc": "按模板批量重命名文件，复制到输出目录，不动源文件。",
    "fields": [
        {"key": "pattern", "label": "命名模板", "type": "text", "default": "{index}_{name}",
         "help": "可用 {name} 原名 {index} 序号 {ext} 扩展名 {date} 日期"},
        {"key": "start", "label": "起始序号", "type": "number", "default": 1},
        {"key": "padding", "label": "序号位数", "type": "number", "default": 3},
        {"key": "ext", "label": "新扩展名", "type": "text", "default": "", "help": "留空保持原扩展名"},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs)
    ensure_dir(output_dir)
    pattern = params.get("pattern") or "{index}_{name}"
    start = int(params.get("start") or 1)
    padding = int(params.get("padding") or 3)
    new_ext = (params.get("ext") or "").strip()

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, src in enumerate(files, start):
        if ctx.cancelled:
            ctx.warn("已取消，剩余文件未处理")
            break
        try:
            stem, ext = os.path.splitext(os.path.basename(src))
            ext = new_ext or ext
            name = pattern.format(
                name=stem, index=str(i).zfill(padding),
                ext=ext.lstrip("."), date=date.today().isoformat(),
            )
            dst = unique_path(os.path.join(output_dir, safe_name(name) + ext))
            shutil.copy2(src, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(src)} -> {os.path.basename(dst)}")
        except Exception as e:
            fail += 1
            ctx.error(f"{src}: {e}")
        ctx.progress(i - start + 1, total, os.path.basename(src))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
