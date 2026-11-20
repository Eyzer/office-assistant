"""格式互转：xlsx ↔ csv。"""
import os

from utils import excel_utils as xu
from utils.encoding_utils import write_text
from utils.file_utils import collect_files
from utils.path_utils import out_path

META = {
    "id": "format_convert",
    "name": "表格格式互转",
    "desc": "xlsx 转 CSV（每表一个文件），CSV 转 xlsx。",
    "fields": [
        {"key": "direction", "label": "转换方向", "type": "select", "default": "to_csv",
         "options": [{"value": "to_csv", "label": "xlsx → CSV"}, {"value": "to_xlsx", "label": "CSV → xlsx"}]},
        {"key": "delim", "label": "CSV 分隔符", "type": "text", "default": ","},
    ],
}


def _xlsx_to_csv(path, output_dir, delim, ctx):
    import csv
    import io
    wb = xu.load(path, data_only=True)
    outs = []
    for ws in wb.worksheets:
        buf = io.StringIO()
        writer = csv.writer(buf, delimiter=delim, lineterminator="\n")
        for row in ws.iter_rows(values_only=True):
            writer.writerow(["" if v is None else v for v in row])
        dst = out_path(path, output_dir, suffix=f"_{ws.title}", ext=".csv")
        write_text(dst, buf.getvalue())
        outs.append(dst)
        ctx.info(f"{os.path.basename(path)}[{ws.title}] -> {os.path.basename(dst)}")
    return outs


def _csv_to_xlsx(path, output_dir, delim, ctx):
    import openpyxl
    header, rows = xu.csv_rows(path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for c, v in enumerate(header, 1):
        ws.cell(row=1, column=c, value=v)
    for ri, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            ws.cell(row=ri, column=c, value=v)
    dst = out_path(path, output_dir, suffix="", ext=".xlsx")
    xu.save(wb, dst)
    ctx.info(f"{os.path.basename(path)} -> {os.path.basename(dst)}")
    return [dst]


def run_task(inputs, output_dir, params, ctx):
    direction = params.get("direction") or "to_csv"
    delim = params.get("delim") or ","
    exts = {".xlsx", ".xlsm"} if direction == "to_csv" else {".csv"}
    files = collect_files(inputs, exts)
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            if direction == "to_csv":
                outs.extend(_xlsx_to_csv(f, output_dir, delim, ctx))
            else:
                outs.extend(_csv_to_xlsx(f, output_dir, delim, ctx))
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
