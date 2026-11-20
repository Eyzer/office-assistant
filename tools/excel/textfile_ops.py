"""文本文件与表格互转：txt/csv ↔ xlsx、文本按行拆分、文本合并。"""
import csv
import io
import os

import openpyxl

from utils import batch, excel_utils as xu
from utils.encoding_utils import read_text, write_text
from utils.path_utils import out_path, safe_name

META = {
    "id": "textfile_ops",
    "cat": "导入与导出",
    "name": "文本文件互转",
    "desc": "txt/csv 转表格、表格导成 CSV/TXT、文本按行数拆成多个文件、多个文本合并成一个。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "txt2excel",
         "options": [
             {"value": "txt2excel", "label": "文本/CSV → 表格"},
             {"value": "excel2txt", "label": "表格 → CSV"},
             {"value": "split", "label": "文本按行拆分"},
             {"value": "merge", "label": "多文本合并"},
         ]},
        {"key": "delim", "label": "分隔符", "type": "select", "default": "auto",
         "options": [{"value": "auto", "label": "自动识别"},
                     {"value": "tab", "label": "制表符"},
                     {"value": "comma", "label": "逗号"},
                     {"value": "space", "label": "空格"},
                     {"value": "pipe", "label": "竖线"}]},
        {"key": "lines", "label": "每个文件行数", "type": "number", "default": 1000},
        {"key": "header", "label": "合并时保留表头", "type": "checkbox", "default": True},
    ],
}
_DELIMS = {"tab": "\t", "comma": ",", "space": " ", "pipe": "|"}


def _split_rows(text, delim):
    if delim == "auto":
        sample = text[:4096]
        try:
            d = csv.Sniffer().sniff(sample, delimiters=",\t|;").delimiter
        except csv.Error:
            d = "\t" if "\t" in sample else ","
    else:
        d = _DELIMS.get(delim, ",")
    return list(csv.reader(io.StringIO(text), delimiter=d))


def txt2excel(f, output_dir, delim, ctx):
    rows = _split_rows(read_text(f), delim)
    if not rows:
        raise ValueError("文件是空的")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = safe_name(os.path.splitext(os.path.basename(f))[0])[:31] or "Sheet"
    for row in rows:
        ws.append(row)
    dst = out_path(f, output_dir, suffix="_表格", ext=".xlsx")
    wb.save(dst)
    ctx.info("%s: %d 行" % (os.path.basename(f), len(rows)))
    return [dst]


def excel2txt(f, output_dir, ctx):
    wb = xu.load(f, data_only=True)
    outs = []
    for idx, ws in enumerate(wb.worksheets):
        buf = io.StringIO()
        w = csv.writer(buf)
        for row in ws.iter_rows(values_only=True):
            w.writerow(["" if v is None else v for v in row])
        if len(wb.sheetnames) == 1:
            dst = out_path(f, output_dir, suffix="_导出", ext=".csv")
        else:
            sub = os.path.join(output_dir, safe_name(os.path.splitext(os.path.basename(f))[0]) + "_csv")
            os.makedirs(sub, exist_ok=True)
            dst = os.path.join(sub, safe_name(ws.title) + ".csv")
        write_text(dst, buf.getvalue())
        outs.append(dst)
    ctx.info("%s: 导出 %d 个 CSV" % (os.path.basename(f), len(outs)))
    return outs


def split_text(f, output_dir, lines, ctx):
    text = read_text(f)
    rows = text.splitlines(True)
    lines = max(1, int(lines or 1000))
    stem = safe_name(os.path.splitext(os.path.basename(f))[0])
    sub = os.path.join(output_dir, stem + "_拆分")
    os.makedirs(sub, exist_ok=True)
    ext = os.path.splitext(f)[1] or ".txt"
    outs = []
    for i in range(0, len(rows), lines):
        part = rows[i:i + lines]
        dst = os.path.join(sub, "%s_%03d%s" % (stem, i // lines + 1, ext))
        write_text(dst, "".join(part))
        outs.append(dst)
    ctx.info("%s: 拆成 %d 个文件" % (os.path.basename(f), len(outs)))
    return outs


def merge_text(files, output_dir, header, ctx):
    if not files:
        return []
    stem = "合并文本"
    dst = os.path.join(output_dir, stem + ".csv")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "%s_%d.csv" % (stem, i))
        i += 1
    parts = []
    for idx, f in enumerate(sorted(files)):
        text = read_text(f).strip("\r\n")
        if not text:
            continue
        lines = text.splitlines()
        if idx and header:
            lines = lines[1:]
        parts.extend(lines)
    write_text(dst, "\n".join(parts) + "\n")
    ctx.info("合并 %d 个文件 → %s" % (len(files), os.path.basename(dst)))
    return [dst]


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "txt2excel"
    delim = params.get("delim") or "auto"
    lines = int(params.get("lines") or 1000)
    header = bool(params.get("header", True))
    TXT = {".txt", ".csv", ".tsv"}
    os.makedirs(output_dir, exist_ok=True)

    if mode == "txt2excel":
        def handler(f, ctx):
            return txt2excel(f, output_dir, delim, ctx)
        return batch.process(inputs, ctx, handler, TXT)

    if mode == "excel2txt":
        def handler(f, ctx):
            return excel2txt(f, output_dir, ctx)
        return batch.process(inputs, ctx, handler, {".xlsx", ".xlsm", ".csv"})

    if mode == "split":
        def handler(f, ctx):
            return split_text(f, output_dir, lines, ctx)
        return batch.process(inputs, ctx, handler, TXT)

    # 合并：一次处理全部文件
    from utils.file_utils import collect_files
    files = collect_files(inputs, TXT)
    if not files:
        ctx.warn("没有找到文本文件")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}
    outs = merge_text(files, output_dir, header, ctx)
    return {"success_count": 1, "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
