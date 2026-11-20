"""提取文档中的表格到 Excel。"""
import os

import openpyxl

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_table_extract",
    "name": "提取表格到 Excel",
    "desc": "把文档里的所有表格导出到一个 xlsx，每张表一个 sheet。",
    "fields": [],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            if not d.tables:
                ctx.warn(f"{os.path.basename(f)}: 无表格，跳过")
                continue
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            for ti, table in enumerate(d.tables, 1):
                ws = wb.create_sheet(title=f"表{ti}")
                for ri, row in enumerate(table.rows, 1):
                    for ci, cell in enumerate(row.cells, 1):
                        ws.cell(row=ri, column=ci, value=cell.text.strip())
            dst = out_path(f, output_dir, suffix="_表格", ext=".xlsx")
            from utils import excel_utils as xu
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 提取 {len(d.tables)} 张表")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
