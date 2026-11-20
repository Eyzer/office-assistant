"""两份文档文本差异对比。"""
import difflib
import os

from utils import doc_utils as du
from utils.encoding_utils import write_text
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_compare",
    "name": "文档文本对比",
    "desc": "对比两份文档的正文差异，输出差异报告（不比格式）。",
    "fields": [
        {"key": "context", "label": "上下文行数", "type": "number", "default": 2},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = sorted(collect_files(inputs, WORD_EXT))
    if len(files) < 2:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["需要正好两个 docx 文件"], "output_files": []}

    ctx_n = int(params.get("context") or 2)
    a, b = files[0], files[1]
    ta = du.body_paragraph_text(du.load(a)).splitlines()
    tb = du.body_paragraph_text(du.load(b)).splitlines()

    diff = list(difflib.unified_diff(ta, tb,
                                     fromfile=os.path.basename(a),
                                     tofile=os.path.basename(b),
                                     n=ctx_n, lineterm=""))
    body = "\n".join(diff) if diff else "两份文档正文完全一致"
    dst = out_path(a, output_dir, suffix="_对比", ext=".txt")
    write_text(dst, body)
    ctx.info(f"差异行数 {len(diff)}")
    ctx.progress(1, 1, "完成")

    return {"success_count": 1, "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
