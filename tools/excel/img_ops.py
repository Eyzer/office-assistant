"""图片批处理：缩放、文字水印、打码、拼接、转 PDF、EXIF 与主色调导出。"""
import csv
import os

from utils import batch, deps
from utils.file_utils import collect_files
from utils.path_utils import batch_out_dir, safe_name

META = {
    "id": "img_ops",
    "cat": "图片与条码",
    "name": "图片批处理",
    "desc": "批量缩放、加文字水印、打码、纵向/横向拼接长图、图片转 PDF、导出 EXIF 与主色调。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "resize",
         "options": [
             {"value": "resize", "label": "批量调整大小"},
             {"value": "watermark", "label": "加文字水印"},
             {"value": "mosaic", "label": "打码（像素化）"},
             {"value": "compose", "label": "拼接成长图"},
             {"value": "to_pdf", "label": "图片转 PDF"},
             {"value": "exif", "label": "导出 EXIF 信息"},
             {"value": "dominant", "label": "导出主色调"},
         ]},
        {"key": "width", "label": "宽(px)", "type": "number", "default": 1200,
         "help": "缩放用；高留 0 表示按比例"},
        {"key": "height", "label": "高(px)", "type": "number", "default": 0},
        {"key": "text", "label": "水印文字", "type": "text", "default": "内部资料"},
        {"key": "blocks", "label": "打码块大小", "type": "number", "default": 12},
        {"key": "direction", "label": "拼接方向", "type": "select", "default": "vertical",
         "options": [{"value": "vertical", "label": "纵向"}, {"value": "horizontal", "label": "横向"}]},
        {"key": "quality", "label": "JPG 质量", "type": "number", "default": 90},
    ],
}

IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
_FONTS = [
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simhei.ttf",
]


def _font(size):
    from PIL import ImageFont
    for p in _FONTS:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


def _open(f):
    from PIL import Image, ImageOps
    im = Image.open(f)
    im = ImageOps.exif_transpose(im)
    return im.convert("RGB")


def do_resize(files, output_dir, w, h, quality, ctx):
    from PIL import Image
    outs = []
    for f in files:
        if ctx.cancelled:
            break
        try:
            im = _open(f)
            if w and not h:
                hh = max(1, round(im.height * w / im.width))
                im = im.resize((int(w), hh), Image.LANCZOS)
            elif h and not w:
                ww = max(1, round(im.width * h / im.height))
                im = im.resize((ww, int(h)), Image.LANCZOS)
            elif w and h:
                im = im.resize((int(w), int(h)), Image.LANCZOS)
            dst = os.path.join(output_dir, safe_name(os.path.splitext(os.path.basename(f))[0]) + "_缩放.jpg")
            i = 1
            while os.path.exists(dst):
                dst = os.path.join(output_dir, safe_name(os.path.splitext(os.path.basename(f))[0]) + "_缩放_%d.jpg" % i)
                i += 1
            im.save(dst, quality=int(quality or 90))
            outs.append(dst)
        except Exception as exc:
            ctx.error("%s: %s" % (os.path.basename(f), exc))
    ctx.info("缩放 %d 张" % len(outs))
    return outs


def do_watermark(files, output_dir, text, quality, ctx):
    from PIL import ImageDraw
    outs = []
    for f in files:
        if ctx.cancelled:
            break
        try:
            im = _open(f)
            draw = ImageDraw.Draw(im, "RGBA")
            size = max(18, im.width // 22)
            font = _font(size)
            pad = size // 2
            draw.text((im.width - len(text) * size - pad, im.height - size * 2),
                      text, font=font, fill=(255, 255, 255, 150))
            dst = os.path.join(output_dir, safe_name(os.path.splitext(os.path.basename(f))[0]) + "_水印.jpg")
            im.save(dst, quality=int(quality or 90))
            outs.append(dst)
        except Exception as exc:
            ctx.error("%s: %s" % (os.path.basename(f), exc))
    ctx.info("加水印 %d 张" % len(outs))
    return outs


def do_mosaic(files, output_dir, blocks, quality, ctx):
    outs = []
    b = max(2, int(blocks or 12))
    for f in files:
        if ctx.cancelled:
            break
        try:
            im = _open(f)
            small = im.resize((max(1, im.width // b), max(1, im.height // b)))
            im = small.resize(im.size)
            dst = os.path.join(output_dir, safe_name(os.path.splitext(os.path.basename(f))[0]) + "_打码.jpg")
            im.save(dst, quality=int(quality or 90))
            outs.append(dst)
        except Exception as exc:
            ctx.error("%s: %s" % (os.path.basename(f), exc))
    ctx.info("打码 %d 张" % len(outs))
    return outs


def do_compose(files, output_dir, direction, ctx):
    from PIL import Image
    if not files:
        return []
    imgs = []
    for f in files:
        try:
            imgs.append(_open(f))
        except Exception as exc:
            ctx.error("%s: %s" % (os.path.basename(f), exc))
    if not imgs:
        return []
    if direction == "horizontal":
        h = max(i.height for i in imgs)
        imgs = [i.resize((round(i.width * h / i.height), h)) for i in imgs]
        canvas = Image.new("RGB", (sum(i.width for i in imgs), h), "white")
        x = 0
        for i in imgs:
            canvas.paste(i, (x, 0))
            x += i.width
    else:
        w = max(i.width for i in imgs)
        imgs = [i.resize((w, round(i.height * w / i.width))) for i in imgs]
        canvas = Image.new("RGB", (w, sum(i.height for i in imgs)), "white")
        y = 0
        for i in imgs:
            canvas.paste(i, (0, y))
            y += i.height
    dst = os.path.join(output_dir, "拼接长图.jpg")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "拼接长图_%d.jpg" % i)
        i += 1
    canvas.save(dst, quality=90)
    ctx.info("%d 张图拼成一张" % len(imgs))
    return [dst]


def do_to_pdf(files, output_dir, ctx):
    if not files:
        return []
    imgs = []
    for f in files:
        try:
            imgs.append(_open(f))
        except Exception as exc:
            ctx.error("%s: %s" % (os.path.basename(f), exc))
    if not imgs:
        return []
    dst = os.path.join(output_dir, "图片合集.pdf")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "图片合集_%d.pdf" % i)
        i += 1
    imgs[0].save(dst, save_all=True, append_images=imgs[1:])
    ctx.info("%d 张图片 → PDF" % len(imgs))
    return [dst]


def do_exif(files, output_dir, ctx):
    from PIL import Image, ExifTags
    dst = os.path.join(output_dir, "图片EXIF.csv")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "图片EXIF_%d.csv" % i)
        i += 1
    with open(dst, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["文件", "拍摄时间", "相机", "镜头", "光圈", "快门", "ISO", "尺寸", "大小KB"])
        for f in files:
            try:
                im = Image.open(f)
                ex = {}
                for k, v in (im.getexif() or {}).items():
                    ex[ExifTags.TAGS.get(k, k)] = v
                w.writerow([os.path.basename(f), ex.get("DateTimeOriginal", ""),
                            "%s %s" % (ex.get("Make", ""), ex.get("Model", "")),
                            ex.get("LensModel", ""), ex.get("FNumber", ""),
                            ex.get("ExposureTime", ""), ex.get("ISOSpeedRatings", ""),
                            "%dx%d" % im.size, round(os.path.getsize(f) / 1024)])
            except Exception as exc:
                ctx.warn("%s: %s" % (os.path.basename(f), exc))
    ctx.info("导出 EXIF 清单")
    return [dst]


def do_dominant(files, output_dir, ctx):
    dst = os.path.join(output_dir, "图片主色调.csv")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "图片主色调_%d.csv" % i)
        i += 1
    with open(dst, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["文件", "主色HEX", "R", "G", "B"])
        for f in files:
            try:
                im = _open(f).resize((64, 64))
                q = im.quantize(colors=1, method=2)
                rgb = q.convert("RGB").getpixel((0, 0))
                w.writerow([os.path.basename(f), "#%02X%02X%02X" % rgb, rgb[0], rgb[1], rgb[2]])
            except Exception as exc:
                ctx.warn("%s: %s" % (os.path.basename(f), exc))
    ctx.info("导出主色调清单")
    return [dst]


def run_task(inputs, output_dir, params, ctx):
    if not deps.need(ctx, ["PIL"]):
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}
    mode = params.get("mode") or "resize"
    w = int(params.get("width") or 0)
    h = int(params.get("height") or 0)
    text = params.get("text") or "内部资料"
    blocks = int(params.get("blocks") or 12)
    direction = params.get("direction") or "vertical"
    quality = int(params.get("quality") or 90)
    os.makedirs(output_dir, exist_ok=True)

    files = sorted(collect_files(inputs, IMG_EXT))
    if not files:
        ctx.warn("没有找到图片文件")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    if mode == "resize":
        outs = do_resize(files, output_dir, w, h, quality, ctx)
    elif mode == "watermark":
        outs = do_watermark(files, output_dir, text, quality, ctx)
    elif mode == "mosaic":
        outs = do_mosaic(files, output_dir, blocks, quality, ctx)
    elif mode == "compose":
        outs = do_compose(files, output_dir, direction, ctx)
    elif mode == "to_pdf":
        outs = do_to_pdf(files, output_dir, ctx)
    elif mode == "exif":
        outs = do_exif(files, output_dir, ctx)
    else:
        outs = do_dominant(files, output_dir, ctx)

    return {"success_count": len(files), "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
