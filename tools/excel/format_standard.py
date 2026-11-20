"""批量标准化单元格格式：字体、字号、对齐、行高列宽。"""
import os

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "format_standard",
    "name": "批量标准化格式",
    "desc": "统一字体、字号、对齐方式、边框底色，可自动适配列宽。",
    "fields": [
        {"key": "font", "label": "字体", "type": "text", "default": "宋体"},
        {"key": "size", "label": "字号", "type": "number", "default": 11},
        {"key": "align", "label": "对齐", "type": "select", "default": "center",
         "options": [{"value": "center", "label": "居中"}, {"value": "left", "label": "左对齐"},
                     {"value": "right", "label": "右对齐"}]},
        {"key": "auto_width", "label": "自动列宽", "type": "checkbox", "default": True},
        {"key": "header_bold", "label": "表头加粗", "type": "checkbox", "default": True},
        {"key": "fill", "label": "底色(十六进制)", "type": "text", "default": "",
         "help": "如 FFEEDD，留空不改"},
    ],
}

_ALIGN = {"center": "center", "left": "left", "right": "right"}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    font_name = params.get("font") or "宋体"
    size = int(params.get("size") or 11)
    align = _ALIGN.get(params.get("align") or "center", "center")
    auto_width = bool(params.get("auto_width", True))
    header_bold = bool(params.get("header_bold", True))
    fill_hex = (params.get("fill") or "").strip()

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
                        cell.font = Font(name=font_name, size=size, bold=cell.font.bold)
                        cell.alignment = Alignment(horizontal=align, vertical="center")
                        if fill_hex:
                            cell.fill = PatternFill("solid", fgColor=fill_hex)
                if header_bold and ws.max_row >= 1:
                    for cell in ws[1]:
                        cell.font = Font(name=font_name, size=size, bold=True)
                if auto_width:
                    for c in range(1, ws.max_column + 1):
                        width = max(
                            (len(str(ws.cell(row=r, column=c).value or "")) for r in range(1, min(ws.max_row, 200) + 1)),
                            default=8,
                        )
                        ws.column_dimensions[get_column_letter(c)].width = min(max(width * 1.6, 8), 60)
            dst = out_path(f, output_dir, suffix="_格式")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 格式已统一")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
