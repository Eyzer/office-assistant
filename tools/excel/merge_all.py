"""合并多个表格到一个总表：自动识别表头，过滤重复表头。"""
import os

import openpyxl
from openpyxl.utils import get_column_letter

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "merge_all",
    "name": "合并表格到总表",
    "desc": "多文件纵向合并为一张总表，自动识别表头并过滤重复表头行。",
    "fields": [
        {"key": "sheet", "label": "合并的工作表", "type": "text", "default": "",
         "help": "留空=每个文件第一个工作表；填名称=同名表；填 * = 全部工作表"},
        {"key": "source_col", "label": "加来源列", "type": "checkbox", "default": True},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 0,
         "help": "0=自动识别"},
    ],
}


def _rows_of(path, sheet_opt, header_row):
    """返回 (header, rows) 列表，按 sheet 拆。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        h, rows = xu.csv_rows(path)
        return [("CSV", h, rows)]
    wb = xu.load(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        if sheet_opt == "*":
            pass
        elif sheet_opt and ws.title != sheet_opt:
            continue
        elif not sheet_opt and ws.title != wb.worksheets[0].title:
            continue
        hr = header_row or xu.detect_header_row(ws)
        h = xu.header_values(ws, hr)
        rows = []
        for r in range(hr + 1, ws.max_row + 1):
            vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
            if all(v in (None, "") for v in vals):
                continue
            rows.append(vals)
        out.append((ws.title, h, rows))
    return out


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    if not files:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["未找到表格文件"], "output_files": []}

    sheet_opt = (params.get("sheet") or "").strip()
    add_source = bool(params.get("source_col", True))
    header_row = int(params.get("header_row") or 0)

    dst = out_path(files[0], output_dir, suffix="_合并总表")
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "合并结果"

    header_written = False
    out_row = 1
    ok, fail = 0, 0
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            blocks = _rows_of(f, sheet_opt, header_row)
            fname = os.path.basename(f)
            for sheet_name, header, rows in blocks:
                if not header_written:
                    cols = list(header)
                    if add_source:
                        cols = ["来源文件", "来源工作表"] + cols
                    for c, v in enumerate(cols, 1):
                        ws_out.cell(row=1, column=c, value=v)
                    ws_out.freeze_panes = "A2"
                    header_written = True
                for row in rows:
                    if header and [str(x).strip() if x is not None else "" for x in row[:len(header)]] == \
                            [str(x).strip() if x is not None else "" for x in header]:
                        continue  # 过滤重复表头
                    vals = ([fname, sheet_name] + list(row)) if add_source else list(row)
                    for c, v in enumerate(vals, 1):
                        ws_out.cell(row=out_row + 1, column=c, value=v)
                    out_row += 1
            ok += 1
            ctx.info(f"已合并 {fname}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    xu.save(wb_out, dst)
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
