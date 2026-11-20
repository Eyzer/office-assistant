"""批量统一文档正文字体与字号。"""
import os

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_content_style",
    "name": "正文样式统一",
    "desc": "统一正文的中文字体、字号与加粗状态。",
    "fields": [
        {"key": "font", "label": "中文字体", "type": "text", "default": "仿宋_GB2312"},
        {"key": "size", "label": "字号(磅)", "type": "number", "default": 16},
        {"key": "bold", "label": "加粗", "type": "select", "default": "keep",
         "options": [{"value": "keep", "label": "保持原样"}, {"value": "on", "label": "全部加粗"},
                     {"value": "off", "label": "全部取消"}]},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    font = params.get("font") or "仿宋_GB2312"
    size = float(params.get("size") or 16)
    bold_opt = params.get("bold") or "keep"
    bold = True if bold_opt == "on" else (False if bold_opt == "off" else None)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            du.set_all_font(d, name=font, size=size, bold=bold)
            dst = out_path(f, output_dir, suffix="_样式")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {font} {size}磅")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
