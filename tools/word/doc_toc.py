"""在文档开头批量插入目录域（TOC）。"""
import os

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK

from utils import batch
from utils.file_utils import WORD_EXT
from utils.path_utils import out_path

META = {
    "id": "doc_toc",
    "cat": "文档结构",
    "name": "批量插入目录",
    "desc": "在每份文档开头插入自动目录域，打开后按 F9（或右键更新域）即可生成页码。",
    "fields": [
        {"key": "title", "label": "目录标题", "type": "text", "default": "目　录"},
        {"key": "levels", "label": "收录几级标题", "type": "number", "default": 3},
        {"key": "page_break", "label": "目录后分页", "type": "checkbox", "default": True},
    ],
}


def _fld(p, kind):
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), kind)
    p.add_run()._r.append(el)


def add_toc(doc, title, levels, page_break):
    from docx.shared import Pt
    anchor = doc.paragraphs[0] if doc.paragraphs else None
    if anchor is None:
        return
    tp = anchor.insert_paragraph_before(title)
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in tp.runs:
        r.bold = True
        r.font.size = Pt(16)

    p = anchor.insert_paragraph_before()
    _fld(p, "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-%d" \\h \\z \\u' % max(1, int(levels))
    p.add_run()._r.append(instr)
    _fld(p, "separate")
    p.add_run("（在此处右键 → 更新域，即可生成带页码的目录）")
    _fld(p, "end")

    if page_break:
        br = anchor.insert_paragraph_before()
        br.add_run().add_break(WD_BREAK.PAGE)


def run_task(inputs, output_dir, params, ctx):
    title = params.get("title") or "目　录"
    levels = int(params.get("levels") or 3)
    page_break = bool(params.get("page_break", True))

    def handler(f, ctx):
        import docx
        doc = docx.Document(f)
        add_toc(doc, title, levels, page_break)
        dst = out_path(f, output_dir, suffix="_带目录")
        doc.save(dst)
        ctx.info("%s: 已插入目录域" % os.path.basename(f))
        return [dst]

    return batch.process(inputs, ctx, handler, WORD_EXT)
