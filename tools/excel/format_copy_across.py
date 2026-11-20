"""把一个模板文件的样式批量套到多个文件。"""
import os

import openpyxl

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "format_copy_across",
    "name": "格式批量套用",
    "desc": "取模板文件首表首行的字体/底纹/对齐，套到目标文件的所有单元格。",
    "fields": [
        {"key": "template", "label": "模板文件路径", "type": "text", "default": ""},
        {"key": "header_row", "label": "样式取样行", "type": "number", "default": 1},
        {"key": "auto_width", "label": "同时自动列宽", "type": "checkbox", "default": True},
    ],
}


def _style_of(cell):
    return {
        "font": cell.font.copy(),
        "fill": cell.fill.copy(),
        "border": cell.border.copy(),
        "alignment": cell.alignment.copy(),
        "number_format": cell.number_format,
    }


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    template = (params.get("template") or "").strip()
    if not template:
        template = next((p for p in inputs if p.lower().endswith((".xlsx", ".xlsm"))), "")
    if not template or not os.path.isfile(template):
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["未找到模板文件"], "output_files": []}

    header_row = int(params.get("header_row") or 1)
    auto_width = bool(params.get("auto_width", True))

    twb = xu.load(template)
    tws = twb.worksheets[0]
    proto = _style_of(tws.cell(row=header_row, column=1)) if tws.max_row >= header_row else None
    if proto is None:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["模板行为空"], "output_files": []}

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        cell.font = proto["font"].copy()
                        cell.fill = proto["fill"].copy()
                        cell.border = proto["border"].copy()
                        cell.alignment = proto["alignment"].copy()
                if auto_width:
                    from openpyxl.utils import get_column_letter
                    for c in range(1, ws.max_column + 1):
                        w = max((len(str(ws.cell(row=r, column=c).value or ""))
                                 for r in range(1, min(ws.max_row, 200) + 1)), default=8)
                        ws.column_dimensions[get_column_letter(c)].width = min(max(w * 1.6, 8), 50)
            dst = out_path(f, output_dir, suffix="_套格式")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 已套用模板样式")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
