"""批量设置页眉页脚与页码。"""
import os

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_header_footer",
    "name": "页眉页脚设置",
    "desc": "统一页眉文字，页脚可插入页码；也可一键清空页眉页脚。",
    "fields": [
        {"key": "header", "label": "页眉文字", "type": "text", "default": ""},
        {"key": "footer", "label": "页脚文字", "type": "text", "default": ""},
        {"key": "page_number", "label": "页脚插入页码", "type": "checkbox", "default": False},
        {"key": "clear", "label": "先清空原页眉页脚", "type": "checkbox", "default": True},
    ],
}


def _add_page_field(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    header = params.get("header") or ""
    footer = params.get("footer") or ""
    page_number = bool(params.get("page_number"))
    clear = bool(params.get("clear", True))

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            if clear:
                du.clear_headers_footers(d)
            for s in d.sections:
                if header:
                    s.header.is_linked_to_previous = False
                    p = s.header.paragraphs[0] if s.header.paragraphs else s.header.add_paragraph()
                    p.text = header
                if footer or page_number:
                    s.footer.is_linked_to_previous = False
                    p = s.footer.paragraphs[0] if s.footer.paragraphs else s.footer.add_paragraph()
                    p.text = footer
                    if page_number:
                        _add_page_field(p)
            dst = out_path(f, output_dir, suffix="_页眉页脚")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 页眉页脚已设置")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
