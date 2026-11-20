"""批量设置文档页面样式：页边距、纸张、方向。"""
import os

from docx.enum.section import WD_ORIENT

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_page_style",
    "name": "页面样式统一",
    "desc": "批量设置页边距、纸张大小与横竖方向。",
    "fields": [
        {"key": "margin", "label": "页边距(厘米)", "type": "number", "default": 2.54},
        {"key": "paper", "label": "纸张", "type": "select", "default": "A4",
         "options": [{"value": "A4", "label": "A4"}, {"value": "A3", "label": "A3"}]},
        {"key": "landscape", "label": "横向", "type": "checkbox", "default": False},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    margin = float(params.get("margin") or 2.54)
    paper = params.get("paper") or "A4"
    landscape = bool(params.get("landscape"))
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            du.set_page(d, margin_cm=margin, paper=paper,
                        orientation=WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT)
            dst = out_path(f, output_dir, suffix="_页面")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {paper} 边距{margin}cm")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
