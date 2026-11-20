"""文档优化压缩：重打包、去缩略图、可选压缩图片。"""
import os
import shutil
import zipfile

from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_compress",
    "name": "文档优化压缩",
    "desc": "重新打包并去掉缩略图等冗余部件；装了 Pillow 时自动压图。",
    "fields": [
        {"key": "drop_thumb", "label": "删除缩略图", "type": "checkbox", "default": True},
        {"key": "image_quality", "label": "图片质量(1-95，0=不动)", "type": "number", "default": 0},
    ],
}

SKIP = ("docProps/thumbnail",)


def _has_pillow():
    try:
        import PIL  # noqa: F401
        return True
    except ImportError:
        return False


def _compress(path, dst, drop_thumb, quality):
    tmp_dir = dst + ".tmp"
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        with zipfile.ZipFile(path) as zin:
            for item in zin.infolist():
                if drop_thumb and any(item.filename.startswith(s) for s in SKIP):
                    continue
                data = zin.read(item.filename)
                if quality and _has_pillow() and item.filename.startswith("word/media/"):
                    data = _shrink(data, quality)
                target = os.path.join(tmp_dir, item.filename)
                os.makedirs(os.path.dirname(target) or tmp_dir, exist_ok=True)
                with open(target, "wb") as f:
                    f.write(data)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for root, _, files in os.walk(tmp_dir):
                for name in files:
                    full = os.path.join(root, name)
                    rel = os.path.relpath(full, tmp_dir).replace(os.sep, "/")
                    zout.write(full, rel)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return dst


def _shrink(data, quality):
    try:
        import io

        from PIL import Image
        img = Image.open(io.BytesIO(data))
        if img.width * img.height < 50000:
            return data
        buf = io.BytesIO()
        fmt = img.format or "JPEG"
        if fmt.upper() not in ("JPEG", "PNG"):
            return data
        if fmt.upper() == "PNG":
            img.save(buf, format="PNG", optimize=True)
        else:
            img.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
        out = buf.getvalue()
        return out if len(out) < len(data) else data
    except Exception:
        return data


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    drop_thumb = bool(params.get("drop_thumb", True))
    quality = int(params.get("image_quality") or 0)
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            dst = out_path(f, output_dir, suffix="_压缩")
            _compress(f, dst, drop_thumb, quality)
            before, after = os.path.getsize(f), os.path.getsize(dst)
            pct = (1 - after / before) * 100 if before else 0
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {before // 1024}KB -> {after // 1024}KB ({pct:.0f}%)")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
