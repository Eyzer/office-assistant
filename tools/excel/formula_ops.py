"""公式处理：按模板批量填充、向下填充、相对/绝对引用互转、导出公式清单。"""
import re

from openpyxl.formula.translate import Translator

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "formula_ops",
    "cat": "数值与计算",
    "name": "公式批量填充与引用转换",
    "desc": "用模板公式整列填充（{r} 代表行号）、把首行公式向下填充、相对引用与绝对引用互转、导出全部公式清单。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "fill",
         "options": [
             {"value": "fill", "label": "按模板填充公式"},
             {"value": "fill_down", "label": "首行公式向下填充"},
             {"value": "ref_absolute", "label": "相对引用 → 绝对引用"},
             {"value": "ref_relative", "label": "绝对引用 → 相对引用"},
             {"value": "extract", "label": "导出公式清单"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "col", "label": "目标列", "type": "text", "default": "D"},
        {"key": "template", "label": "公式模板", "type": "text", "default": "=B{r}*C{r}",
         "help": "{r} 会被替换成行号"},
        {"key": "start_row", "label": "起始行", "type": "number", "default": 2},
    ],
}

_REF = re.compile(r"(?<![$\w'!:])([A-Z]{1,3})(\d{1,7})(?![\w(])")


def to_absolute(formula):
    return _REF.sub(lambda m: "$%s$%s" % (m.group(1), m.group(2)), formula)


def to_relative(formula):
    return formula.replace("$", "")


def do_fill(ws, col, template, header_row, start_row):
    t = str(template or "").strip()
    if not t:
        raise ValueError("公式模板不能为空，示例：=B{r}*C{r}")
    if not t.startswith("="):
        t = "=" + t
    c = batch.col_index(col, 4)
    n = 0
    for r in range(max(start_row, header_row + 1), ws.max_row + 1):
        ws.cell(row=r, column=c, value=t.replace("{r}", str(r)))
        n += 1
    return n


def do_fill_down(ws, col, header_row):
    c = batch.col_index(col, 4)
    src_row = header_row + 1
    origin = ws.cell(row=src_row, column=c).value
    if not isinstance(origin, str) or not origin.startswith("="):
        raise ValueError("第 %d 行的 %s 列不是公式，无法向下填充" % (src_row, col))
    origin_coord = ws.cell(row=src_row, column=c).coordinate
    n = 0
    for r in range(src_row + 1, ws.max_row + 1):
        coord = ws.cell(row=r, column=c).coordinate
        try:
            ws.cell(row=r, column=c, value=Translator(origin, origin=origin_coord).translate_formula(coord))
            n += 1
        except Exception:
            ws.cell(row=r, column=c, value=origin)
            n += 1
    return n


def do_ref(ws, absolute):
    n = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                nv = to_absolute(v) if absolute else to_relative(v)
                if nv != v:
                    cell.value = nv
                    n += 1
    return n


def do_extract(wb, ws):
    items = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                items.append((cell.coordinate, v))
    name = xu.safe_sheet_title("公式_" + ws.title, {s.title.lower() for s in wb.worksheets})
    out = wb.create_sheet(name)
    out.append(["单元格", "公式"])
    for coord, f in items:
        out.append([coord, f])
    return len(items)


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "fill"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    col = params.get("col") or "D"
    template = params.get("template") or ""
    start_row = int(params.get("start_row") or 2)

    def edit(wb, path, ctx):
        total = 0
        if mode == "extract":
            for ws in batch.target_sheets(wb, sheet):
                total += do_extract(wb, ws)
            return "导出 %d 条公式" % total
        for ws in batch.target_sheets(wb, sheet):
            if mode == "fill":
                total += do_fill(ws, col, template, header_row, start_row)
            elif mode == "fill_down":
                total += do_fill_down(ws, col, header_row)
            elif mode == "ref_absolute":
                total += do_ref(ws, True)
            elif mode == "ref_relative":
                total += do_ref(ws, False)
        return "处理 %d 个单元格" % total

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_公式处理", exts=EXCEL_EXT)
