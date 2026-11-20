"""Word 底层封装：python-docx + 直接操作 XML 部件。"""
import copy
import os
import re
import shutil
import zipfile

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def load(path):
    return docx.Document(path)


def save(d, path):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    d.save(path)
    return path


def body_paragraph_text(d):
    return "\n".join(p.text for p in d.paragraphs)


def _iter_block_items(d):
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    body = d.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, d)
        elif child.tag == qn("w:tbl"):
            yield Table(child, d)


def _sect_pr(body):
    for child in body.iterchildren():
        if child.tag == qn("w:sectPr"):
            return child
    return None


def _remap_rels(element, src_part, dst_part):
    """把深拷贝元素里的 r:embed / r:id / r:link 重新挂到目标文档。"""
    for el in element.iter():
        for attr in (qn("r:embed"), qn("r:id"), qn("r:link"), qn("r:href")):
            rid = el.get(attr)
            if not rid:
                continue
            try:
                rel = src_part.rels[rid]
            except KeyError:
                continue
            try:
                new_rid = dst_part.relate_to(rel.target_part, rel.reltype)
            except Exception:
                continue
            el.set(attr, new_rid)


def append_document(target, src_path, page_break=True):
    """把一个 docx 的正文追加到目标文档末尾，保留图片与超链接关系。"""
    src = docx.Document(src_path)
    if page_break:
        target.add_page_break()

    body = target.element.body
    sect = _sect_pr(body)
    dst_part = target.part
    src_part = src.part

    for child in list(src.element.body.iterchildren()):
        if child.tag == qn("w:sectPr"):
            continue
        new_el = copy.deepcopy(child)
        _remap_rels(new_el, src_part, dst_part)
        if sect is not None:
            sect.addprevious(new_el)
        else:
            body.append(new_el)
    return target


def merge_files(paths, out_path):
    """按顺序合并多个文档。返回输出路径。"""
    if not paths:
        raise ValueError("没有要合并的文档")
    base = docx.Document(paths[0])
    base.add_page_break()
    for p in paths[1:]:
        append_document(base, p, page_break=(p != paths[-1]))
    return save(base, out_path)


def set_core_props(d, **kw):
    cp = d.core_properties
    for k, v in kw.items():
        if v is not None and hasattr(cp, k):
            setattr(cp, k, v)
    return d


def clear_core_props(d):
    cp = d.core_properties
    for k in ("author", "last_modified_by", "title", "subject",
              "keywords", "comments", "category", "content_status"):
        try:
            setattr(cp, k, "")
        except Exception:
            pass
    return d


def set_page(d, margin_cm=None, orientation=None, paper=None):
    for s in d.sections:
        if margin_cm is not None:
            s.top_margin = Cm(margin_cm)
            s.bottom_margin = Cm(margin_cm)
            s.left_margin = Cm(margin_cm)
            s.right_margin = Cm(margin_cm)
        if orientation is not None:
            s.orientation = orientation
        if paper == "A4":
            s.page_width, s.page_height = Cm(21), Cm(29.7)
        elif paper == "A3":
            s.page_width, s.page_height = Cm(29.7), Cm(42)
    return d


def _set_run_font(run, name=None, size=None, bold=None):
    if name:
        run.font.name = name
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:eastAsia"), name)
        rfonts.set(qn("w:ascii"), name)
        rfonts.set(qn("w:hAnsi"), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def set_all_font(d, name=None, size=None, bold=None):
    for p in d.paragraphs:
        for r in p.runs:
            _set_run_font(r, name, size, bold)
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        _set_run_font(r, name, size, bold)
    return d


def delete_paragraph(p):
    el = p._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def remove_blank_paragraphs(d, max_consecutive=0):
    removed, run = 0, 0
    for p in list(d.paragraphs):
        if p.text.strip() == "":
            run += 1
            if max_consecutive == 0 or run > max_consecutive:
                delete_paragraph(p)
                removed += 1
        else:
            run = 0
    return removed


def remove_comments(d):
    body = d.element.body
    removed = 0
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for el in list(body.iter(qn("w:" + tag))):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
                removed += 1
    return removed


def remove_hyperlinks(d):
    """超链接转纯文本，保留可见文字。"""
    count = 0
    body = d.element.body
    for h in list(body.iter(qn("w:hyperlink"))):
        parent = h.getparent()
        if parent is None:
            continue
        idx = list(parent).index(h)
        for run in h.findall(qn("w:r")):
            parent.insert(idx, copy.deepcopy(run))
            idx += 1
        parent.remove(h)
        count += 1
    return count


def extract_images(docx_path, out_dir, prefix=""):
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    with zipfile.ZipFile(docx_path) as z:
        for n in z.namelist():
            if n.startswith("word/media/"):
                base = os.path.basename(n)
                dst = os.path.join(out_dir, prefix + base)
                stem, ext = os.path.splitext(dst)
                i = 1
                while os.path.exists(dst):
                    dst = f"{stem}_{i}{ext}"
                    i += 1
                with z.open(n) as src, open(dst, "wb") as f:
                    shutil.copyfileobj(src, f)
                saved.append(dst)
    return saved


def replace_text(d, old, new, case_sensitive=True, whole_word=False):
    """段落与表格内替换，按 run 处理以保留格式。返回替换次数。"""
    count = 0
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.escape(old)
    if whole_word:
        pattern = r"\b" + pattern + r"\b"
    rx = re.compile(pattern, flags)

    def _do(p):
        nonlocal count
        for run in p.runs:
            t = run.text
            if not t:
                continue
            new_t, n = rx.subn(new, t)
            if n:
                run.text = new_t
                count += n

    for p in d.paragraphs:
        _do(p)
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _do(p)
    for s in d.sections:
        for p in s.header.paragraphs:
            _do(p)
        for p in s.footer.paragraphs:
            _do(p)
    return count


def remove_field_codes(d):
    """删除域代码，保留结果文本。"""
    count = 0
    body = d.element.body
    for fld in list(body.iter(qn("w:fldSimple"))):
        parent = fld.getparent()
        if parent is None:
            continue
        idx = list(parent).index(fld)
        for child in list(fld):
            parent.insert(idx, child)
            idx += 1
        parent.remove(fld)
        count += 1
    return count


def set_header_text(d, text):
    for s in d.sections:
        for p in list(s.header.paragraphs):
            delete_paragraph(p)
        s.header.paragraphs[0].text = text
    return d


def clear_headers_footers(d):
    for s in d.sections:
        for part in (s.header, s.footer):
            for p in list(part.paragraphs):
                delete_paragraph(p)
    return d
