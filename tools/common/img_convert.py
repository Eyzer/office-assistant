"""图片格式批量转换与压缩。"""
import os

from utils.file_utils import collect_files
from utils.path_utils import ensure_dir, unique_path

META = {
    "id": "img_convert",
    "name": "图片格式转换",
    "desc": "png/jpg/webp 批量互转，可指定质量与最大边长。",
    "fields": [
        {"key": "fmt", "label": "目标格式", "type": "select", "default": "jpg",
         "options": [{"value": "jpg", "label": "JPG"}, {"value": "png", "label": "PNG"},
                     {"value": "webp", "label": "WEBP"}]},
        {"key": "quality", "label": "质量(1-95)", "type": "number", "default": 85},
        {"key": "max_edge", "label": "最大边长(0=不改)", "type": "number", "default": 0},
    ],
}

EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}


def run_task(inputs, output_dir, params, ctx):
    try:
        from PIL import Image
    except ImportError:
        msg = "缺少 Pillow，请先安装：pip install pillow"
        ctx.error(msg)
        return {"success_count": 0, "fail_count": 1,
                "log_messages": [msg], "output_files": []}

    files = collect_files(inputs, EXTS)
    ensure_dir(output_dir)
    fmt = (params.get("fmt") or "jpg").lower()
    quality = int(params.get("quality") or 85)
    max_edge = int(params.get("max_edge") or 0)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            img = Image.open(f)
            if max_edge and max(img.size) > max_edge:
                img.thumbnail((max_edge, max_edge))
            if fmt in ("jpg", "jpeg") and img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            dst = unique_path(os.path.join(
                output_dir, os.path.splitext(os.path.basename(f))[0] + "." + fmt))
            save_kw = {"quality": quality} if fmt in ("jpg", "jpeg", "webp") else {}
            img.save(dst, **save_kw)
            outs.append(dst)
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
