"""表格转文本、表格行列翻转。"""
import os

from docx.oxml.ns import qn

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_table_convert",
    "name": "表格转文本/翻转",
    "desc": "把文档表格转成纯文本段落，或翻转表格行列。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "to_text",
         "options": [{"value": "to_text", "label": "表格转文本"}, {"value": "transpose", "label": "翻转行列"}]},
        {"key": "sep", "label": "列分隔符", "type": "text", "default": " | "},
    ],
}


def _table_to_text(d, sep):
    count = 0
    for table in list(d.tables):
        lines = []
        for row in table.rows:
            lines.append(sep.join(c.text.strip() for c in row.cells))
        parent = table._element.getparent()
        idx = list(parent).index(table._element)
        parent.remove(table._element)
        for k, line in enumerate(lines):
            p = d.add_paragraph(line)
            parent.insert(idx + k, p._element)
            count += 1
    return count


def _transpose(d):
    count = 0
    for table in d.tables:
        rows = [[c.text.strip() for c in row.cells] for row in table.rows]
        if not rows:
            continue
        cols = list(zip(*rows))
        # 清空原表并重填
        for row in list(table.rows):
            table._element.remove(row._tr)
        for r in cols:
            cells = table.add_row().cells
            for ci, v in enumerate(r):
                if ci < len(cells):
                    cells[ci].text = v
        count += 1
    return count


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    mode = params.get("mode") or "to_text"
    sep = params.get("sep", " | ")
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            n = _table_to_text(d, sep) if mode == "to_text" else _transpose(d)
            dst = out_path(f, output_dir, suffix="_表格处理")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 处理 {n} 行/表")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
