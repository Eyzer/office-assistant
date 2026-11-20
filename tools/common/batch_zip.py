"""批量压缩为 zip 或解压。"""
import os
import zipfile

from utils.file_utils import collect_files
from utils.path_utils import ensure_dir, unique_path

META = {
    "id": "batch_zip",
    "name": "批量压缩/解压",
    "desc": "把文件分别压缩成单独 zip；或把 zip 解压到各自目录。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "zip",
         "options": [{"value": "zip", "label": "压缩"}, {"value": "unzip", "label": "解压"}]},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "zip"
    ensure_dir(output_dir)
    ok, fail, outs = 0, 0, []

    if mode == "zip":
        files = collect_files(inputs)
        total = len(files)
        for i, f in enumerate(files, 1):
            if ctx.cancelled:
                ctx.warn("已取消")
                break
            try:
                dst = unique_path(os.path.join(
                    output_dir, os.path.splitext(os.path.basename(f))[0] + ".zip"))
                with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
                    z.write(f, os.path.basename(f))
                outs.append(dst)
                ok += 1
            except Exception as e:
                fail += 1
                ctx.error(f"{f}: {e}")
            ctx.progress(i, total, os.path.basename(f))
    else:
        files = collect_files(inputs, {".zip"})
        total = len(files)
        for i, f in enumerate(files, 1):
            if ctx.cancelled:
                ctx.warn("已取消")
                break
            try:
                sub = ensure_dir(os.path.join(
                    output_dir, os.path.splitext(os.path.basename(f))[0]))
                with zipfile.ZipFile(f) as z:
                    z.extractall(sub)
                outs.append(sub)
                ok += 1
                ctx.info(f"解压 {os.path.basename(f)}")
            except Exception as e:
                fail += 1
                ctx.error(f"{f}: {e}")
            ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
