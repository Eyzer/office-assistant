"""文本文件编码转换：自动识别源编码，统一转成目标编码。"""
import os

from utils.encoding_utils import detect, read_text
from utils.path_utils import out_path

META = {
    "id": "encoding_ops",
    "cat": "压缩与编码",
    "name": "批量转文件编码",
    "desc": "自动识别 txt/csv 原编码，统一转成 UTF-8（带或不带 BOM）或 GBK，解决打开乱码。",
    "fields": [
        {"key": "to", "label": "目标编码", "type": "select", "default": "utf-8-sig",
         "options": [{"value": "utf-8-sig", "label": "UTF-8（带 BOM，Excel 不乱码）"},
                     {"value": "utf-8", "label": "UTF-8（无 BOM）"},
                     {"value": "gb18030", "label": "GB18030（兼容 GBK）"}]},
        {"key": "suffix", "label": "文件名后缀", "type": "text", "default": "_转码"},
    ],
}

TXT = {".txt", ".csv", ".tsv", ".md", ".json", ".xml", ".html", ".htm", ".ini"}


def run_task(inputs, output_dir, params, ctx):
    from utils.file_utils import collect_files
    to = params.get("to") or "utf-8-sig"
    suffix = params.get("suffix") or "_转码"
    files = collect_files(inputs, TXT)
    if not files:
        ctx.warn("没有找到文本文件")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}
    ok, fail, outs = 0, 0, []
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            break
        try:
            src_enc = detect(f)
            text = read_text(f, [src_enc, "utf-8", "gb18030", "latin-1"])
            dst = out_path(f, output_dir, suffix=suffix)
            with open(dst, "w", encoding=to, newline="") as fh:
                fh.write(text)
            outs.append(dst)
            ok += 1
            ctx.info("%s: %s → %s" % (os.path.basename(f), src_enc, to))
        except Exception as exc:
            fail += 1
            ctx.error("%s: %s" % (os.path.basename(f), exc))
        ctx.progress(i, len(files), os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
