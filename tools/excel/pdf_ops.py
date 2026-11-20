"""PDF 工具：拆分、合并、图片转 PDF、PDF 转图片（PDF 转图片需系统 poppler）。"""
import glob
import os
import shutil
import subprocess

from utils import batch, deps
from utils.file_utils import collect_files
from utils.path_utils import batch_out_dir, safe_name

META = {
    "id": "pdf_ops",
    "cat": "导入与导出",
    "name": "PDF 拆分/合并/转换",
    "desc": "按页拆分 PDF、多个 PDF 合并、图片转 PDF、PDF 转图片（转图片需装 poppler-utils）。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "split",
         "options": [
             {"value": "split", "label": "按页拆分"},
             {"value": "merge", "label": "合并成一个 PDF"},
             {"value": "img2pdf", "label": "图片转 PDF"},
             {"value": "pdf2img", "label": "PDF 转图片"},
         ]},
        {"key": "pages", "label": "每个文件几页", "type": "number", "default": 1},
        {"key": "dpi", "label": "转图片分辨率", "type": "number", "default": 150},
    ],
}

IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}


def do_split(inputs, output_dir, pages, ctx):
    pypdf = deps.load("pypdf", "PDF 拆分")
    pages = max(1, int(pages or 1))
    outs = []
    for f in collect_files(inputs, {".pdf"}):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            reader = pypdf.PdfReader(f)
            sub = batch_out_dir(f, output_dir, "_拆分")
            stem = safe_name(os.path.splitext(os.path.basename(f))[0])
            total = len(reader.pages)
            n = 0
            for start in range(0, total, pages):
                writer = pypdf.PdfWriter()
                for i in range(start, min(start + pages, total)):
                    writer.add_page(reader.pages[i])
                dst = os.path.join(sub, "%s_p%03d-%03d.pdf" % (stem, start + 1, min(start + pages, total)))
                with open(dst, "wb") as fh:
                    writer.write(fh)
                outs.append(dst)
                n += 1
            ctx.info("%s: %d 页拆成 %d 个文件" % (os.path.basename(f), total, n))
        except Exception as exc:
            ctx.error("%s: %s" % (os.path.basename(f), exc))
    return outs


def do_merge(inputs, output_dir, ctx):
    pypdf = deps.load("pypdf", "PDF 合并")
    files = sorted(collect_files(inputs, {".pdf"}))
    if not files:
        ctx.warn("没有找到 PDF")
        return []
    writer = pypdf.PdfWriter()
    for f in files:
        try:
            writer.append(f)
        except Exception as exc:
            ctx.error("%s: %s" % (os.path.basename(f), exc))
    dst = os.path.join(output_dir, "合并结果.pdf")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "合并结果_%d.pdf" % i)
        i += 1
    with open(dst, "wb") as fh:
        writer.write(fh)
    ctx.info("合并 %d 个 PDF → %s" % (len(files), os.path.basename(dst)))
    return [dst]


def do_img2pdf(inputs, output_dir, ctx):
    from PIL import Image
    files = collect_files(inputs, IMG_EXT)
    if not files:
        ctx.warn("没有找到图片")
        return []
    imgs = []
    for f in files:
        try:
            im = Image.open(f).convert("RGB")
            imgs.append(im)
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
    ctx.info("%d 张图片 → %s" % (len(imgs), os.path.basename(dst)))
    return [dst]


def do_pdf2img(inputs, output_dir, dpi, ctx):
    exe = shutil.which("pdftoppm")
    if not exe:
        raise deps.MissingDep("pdftoppm", "PDF 转图片需要 poppler-utils，命令：sudo apt install poppler-utils")
    outs = []
    for f in collect_files(inputs, {".pdf"}):
        if ctx.cancelled:
            break
        sub = batch_out_dir(f, output_dir, "_图片")
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            subprocess.run([exe, "-png", "-r", str(int(dpi or 150)), f, os.path.join(sub, stem)],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
            made = sorted(glob.glob(os.path.join(sub, "*.png")))
            outs.extend(made)
            ctx.info("%s: 转出 %d 张图片" % (os.path.basename(f), len(made)))
        except subprocess.CalledProcessError as exc:
            err = exc.stderr.decode("utf-8", "replace")[:200] if exc.stderr else ""
            ctx.error("%s: 转换失败 %s" % (os.path.basename(f), err))
    return outs


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "split"
    pages = int(params.get("pages") or 1)
    dpi = int(params.get("dpi") or 150)
    os.makedirs(output_dir, exist_ok=True)

    if mode == "split":
        outs = do_split(inputs, output_dir, pages, ctx)
    elif mode == "merge":
        outs = do_merge(inputs, output_dir, ctx)
    elif mode == "img2pdf":
        outs = do_img2pdf(inputs, output_dir, ctx)
    else:
        outs = do_pdf2img(inputs, output_dir, dpi, ctx)

    return {"success_count": 1 if outs else 0, "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
