"""文档只读保护（设置修改限制，不设打开密码）。"""
import os

from docx.oxml.ns import qn

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_protect",
    "name": "文档只读保护",
    "desc": "给文档加上只读修改限制；打开密码需 Office 环境，独立程序不支持。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "on",
         "options": [{"value": "on", "label": "加只读限制"}, {"value": "off", "label": "取消限制"}]},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    mode = params.get("mode") or "on"
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            settings = d.settings.element
            existing = settings.find(qn("w:documentProtection"))
            if existing is not None:
                settings.remove(existing)
            if mode == "on":
                el = settings.makeelement(qn("w:documentProtection"), {})
                el.set(qn("w:edit"), "readOnly")
                el.set(qn("w:enforcement"), "1")
                settings.append(el)
            dst = out_path(f, output_dir, suffix="_保护")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {'加' if mode == 'on' else '解除'}只读限制")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
