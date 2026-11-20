"""按文件夹合并数据：把一个目录下的 xlsx/csv/txt 捞进同一张表。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import collect_files
from utils.path_utils import out_path

META = {
    "id": "merge_folder",
    "name": "按文件夹合并数据",
    "desc": "把一个文件夹（含子目录）里的表格、CSV、TXT 数据合并到一张总表。",
    "fields": [
        {"key": "add_source", "label": "加来源列", "type": "checkbox", "default": True},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}

EXTS = {".xlsx", ".xlsm", ".csv", ".txt"}


def _read_txt(path):
    from utils.encoding_utils import read_text
    text = read_text(path)
    lines = [ln for ln in text.splitlines() if ln.strip()]
    delim = "\t" if any("\t" in ln for ln in lines[:3]) else ","
    return [[c.strip() for c in ln.split(delim)] for ln in lines]


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXTS)
    if not files:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["未找到可合并的数据文件"], "output_files": []}

    add_source = bool(params.get("add_source", True))
    header_row = int(params.get("header_row") or 1)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "合并结果"
    header_written = False
    row_out = 1
    ok, fail = 0, 0
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            ext = os.path.splitext(f)[1].lower()
            if ext == ".txt":
                table = _read_txt(f)
                header, rows = (table[0], table[1:]) if table else ([], [])
            elif ext == ".csv":
                header, rows = xu.csv_rows(f)
            else:
                w = xu.load(f, data_only=True)
                s = w.worksheets[0]
                header = xu.header_values(s, header_row)
                rows = [[s.cell(row=r, column=c).value for c in range(1, s.max_column + 1)]
                        for r in range(header_row + 1, s.max_row + 1)]

            if not header_written:
                cols = (["来源文件"] + list(header)) if add_source else list(header)
                for c, v in enumerate(cols, 1):
                    ws.cell(row=1, column=c, value=v)
                row_out = 2
                header_written = True

            for row in rows:
                if all(v in (None, "") for v in row):
                    continue
                vals = ([os.path.basename(f)] + list(row)) if add_source else list(row)
                for c, v in enumerate(vals, 1):
                    ws.cell(row=row_out, column=c, value=v)
                row_out += 1
            ok += 1
            ctx.info(f"并入 {os.path.basename(f)}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    dst = out_path(files[0], output_dir, suffix="_文件夹合并")
    xu.save(wb, dst)
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
