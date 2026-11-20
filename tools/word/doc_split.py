"""按标题拆分文档为多份。"""
import copy
import os

import docx
from docx.oxml.ns import qn

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import batch_out_dir, safe_name, unique_path

META = {
    "id": "doc_split",
    "name": "按标题拆分文档",
    "desc": "按指定级别标题拆分成独立文档，标题文字作为文件名。",
    "fields": [
        {"key": "level", "label": "标题级别", "type": "select", "default": "1",
         "options": [{"value": "1", "label": "一级标题"}, {"value": "2", "label": "二级标题"}]},
    ],
}


def _heading_level(p):
    try:
        name = p.style.name or ""
    except Exception:
        return 0
    if name.lower().startswith("heading"):
        try:
            return int(name.split()[-1])
        except ValueError:
            return 0
    return 0


def _new_from(base_doc, blocks):
    new = docx.Document()
    body = new.element.body
    for child in list(body.iterchildren()):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    sect = next((c for c in body.iterchildren() if c.tag == qn("w:sectPr")), None)
    for el in blocks:
        new_el = copy.deepcopy(el)
        if sect is not None:
            sect.addprevious(new_el)
        else:
            body.append(new_el)
    return new


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    level = int(params.get("level") or 1)
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            groups = []
            title, blocks = "开头", []
            for child in list(d.element.body.iterchildren()):
                if child.tag == qn("w:sectPr"):
                    continue
                is_heading = False
                if child.tag == qn("w:p"):
                    from docx.text.paragraph import Paragraph
                    p = Paragraph(child, d)
                    if _heading_level(p) == level and p.text.strip():
                        is_heading = True
                        if blocks:
                            groups.append((title, blocks))
                        title, blocks = p.text.strip(), [child]
                if not is_heading:
                    blocks.append(child)
            if blocks:
                groups.append((title, blocks))

            if len(groups) <= 1:
                ctx.warn(f"{os.path.basename(f)}: 只识别到一段，未拆分")
                continue

            sub = batch_out_dir(f, output_dir, "_按标题拆分")
            for ti, (t, bl) in enumerate(groups, 1):
                new = _new_from(d, bl)
                dst = unique_path(os.path.join(sub, safe_name(t)[:60] + ".docx"))
                du.save(new, dst)
                outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 拆成 {len(groups)} 份")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
