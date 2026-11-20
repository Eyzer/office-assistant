"""页面级操作：按分页符提取或删除指定页（近似，依赖手动分页）。"""
import copy
import os

from docx.oxml.ns import qn

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_page_op",
    "name": "提取或删除指定页",
    "desc": "按分页符切页，提取或删除指定页。无分页符的文档按整篇处理。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "extract",
         "options": [{"value": "extract", "label": "提取指定页"}, {"value": "delete", "label": "删除指定页"}]},
        {"key": "pages", "label": "页码", "type": "text", "default": "1",
         "help": "多页用逗号分隔，如 1,3"},
    ],
}


def _split_pages(d):
    """按分页符把正文切成若干页块（元素列表）。"""
    pages, cur = [], []
    for child in list(d.element.body.iterchildren()):
        if child.tag == qn("w:sectPr"):
            continue
        cur.append(child)
        runs = child.iter(qn("w:br"))
        for br in runs:
            if br.get(qn("w:type")) == "page":
                pages.append(cur)
                cur = []
                break
    if cur:
        pages.append(cur)
    return pages


def _page_doc(src_doc, blocks):
    import docx
    new = docx.Document()
    # 清空默认正文
    body = new.element.body
    for child in list(body.iterchildren()):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    sect = None
    for c in body.iterchildren():
        if c.tag == qn("w:sectPr"):
            sect = c
    for el in blocks:
        new_el = copy.deepcopy(el)
        if sect is not None:
            sect.addprevious(new_el)
        else:
            body.append(new_el)
    return new


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    mode = params.get("mode") or "extract"
    try:
        targets = {int(x) for x in str(params.get("pages") or "1").replace("，", ",").split(",") if x.strip()}
    except ValueError:
        targets = {1}

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            pages = _split_pages(d)
            if mode == "extract":
                keep = [pg for idx, pg in enumerate(pages, 1) if idx in targets]
                blocks = [el for pg in keep for el in pg]
                if not blocks:
                    raise ValueError(f"没有第 {targets} 页")
                out_doc = _page_doc(d, blocks)
                dst = out_path(f, output_dir, suffix="_提取页")
            else:
                remove = [pg for idx, pg in enumerate(pages, 1) if idx in targets]
                for pg in remove:
                    for el in pg:
                        parent = el.getparent()
                        if parent is not None:
                            parent.remove(el)
                out_doc = d
                dst = out_path(f, output_dir, suffix="_删页")
            du.save(out_doc, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 共 {len(pages)} 页，处理页码 {sorted(targets)}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
