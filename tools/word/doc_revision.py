"""接受或拒绝全部修订。"""
import copy
import os

from docx.oxml.ns import qn

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_revision",
    "name": "修订批量处理",
    "desc": "接受全部修订（保留新内容）或拒绝全部修订（还原旧内容）。",
    "fields": [
        {"key": "action", "label": "操作", "type": "select", "default": "accept",
         "options": [{"value": "accept", "label": "接受全部修订"}, {"value": "reject", "label": "拒绝全部修订"}]},
    ],
}


def _unwrap(el):
    parent = el.getparent()
    if parent is None:
        return
    idx = list(parent).index(el)
    for child in list(el):
        parent.insert(idx, child)
        idx += 1
    parent.remove(el)


def _accept(d):
    body = d.element.body
    n = 0
    for tag in ("ins",):
        for el in list(body.iter(qn("w:" + tag))):
            _unwrap(el)
            n += 1
    for el in list(body.iter(qn("w:del"))):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
            n += 1
    return n


def _reject(d):
    body = d.element.body
    n = 0
    for el in list(body.iter(qn("w:ins"))):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
            n += 1
    for el in list(body.iter(qn("w:del"))):
        _unwrap(el)
        n += 1
    return n


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    action = params.get("action") or "accept"
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            n = _accept(d) if action == "accept" else _reject(d)
            dst = out_path(f, output_dir, suffix="_修订")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: {'接受' if action == 'accept' else '拒绝'} {n} 处")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
