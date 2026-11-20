"""统一文档内图片尺寸与对齐。"""
import os

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_img_layout",
    "name": "统一图片布局",
    "desc": "把文档里所有内嵌图片统一到指定宽度，并设置所在段落对齐。",
    "fields": [
        {"key": "width_cm", "label": "宽度(厘米)", "type": "number", "default": 12},
        {"key": "align", "label": "对齐", "type": "select", "default": "center",
         "options": [{"value": "center", "label": "居中"}, {"value": "left", "label": "左对齐"}]},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    width = Cm(float(params.get("width_cm") or 12))
    align = WD_ALIGN_PARAGRAPH.CENTER if (params.get("align") or "center") == "center" \
        else WD_ALIGN_PARAGRAPH.LEFT

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            n = 0
            for shape in d.inline_shapes:
                ratio = shape.height / shape.width if shape.width else 1
                shape.width = width
                shape.height = int(width * ratio)
                n += 1
            for p in d.paragraphs:
                if p._element.findall(".//" + "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline"):
                    p.alignment = align
            dst = out_path(f, output_dir, suffix="_图片布局")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 调整 {n} 张图片")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
