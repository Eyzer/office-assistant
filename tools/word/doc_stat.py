"""统计文档字数、段落、表格、图片，导出 CSV 台账。"""
import csv
import os

from utils import batch
from utils.file_utils import WORD_EXT
from utils.path_utils import safe_name

META = {
    "id": "doc_stat",
    "cat": "提取与转换",
    "name": "文档字数统计",
    "desc": "统计每份文档的字符数、段落数、表格数、图片数、页数，汇总成 CSV，可当交稿台账。",
    "fields": [],
}


def count_doc(path):
    import docx
    from docx.oxml.ns import qn
    doc = docx.Document(path)
    chars = 0
    for p in doc.paragraphs:
        chars += len(p.text.strip())
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                chars += len(cell.text.strip())
    imgs = len(doc.inline_shapes) + len(doc.element.body.findall('.//' + qn('a:blip')))
    pages = 0
    try:
        pages = int(doc.core_properties.revision or 0)
    except Exception:
        pages = 0
    return {
        "chars": chars,
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "images": imgs,
        "sections": len(doc.sections),
    }


def run_task(inputs, output_dir, params, ctx):
    from utils.file_utils import collect_files
    files = collect_files(inputs, WORD_EXT)
    if not files:
        ctx.warn("没有找到 docx 文件")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}
    os.makedirs(output_dir, exist_ok=True)
    dst = os.path.join(output_dir, "文档统计.csv")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "文档统计_%d.csv" % i)
        i += 1
    ok, fail = 0, 0
    with open(dst, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["文件名", "字符数", "段落数", "表格数", "图片数", "节数", "大小KB", "路径"])
        for idx, f in enumerate(files, 1):
            if ctx.cancelled:
                break
            try:
                st = count_doc(f)
                w.writerow([safe_name(os.path.basename(f)), st["chars"], st["paragraphs"],
                            st["tables"], st["images"], st["sections"],
                            round(os.path.getsize(f) / 1024), f])
                ok += 1
            except Exception as exc:
                fail += 1
                ctx.error("%s: %s" % (os.path.basename(f), exc))
            ctx.progress(idx, len(files), os.path.basename(f))
    ctx.info("统计 %d 份文档" % ok)
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
