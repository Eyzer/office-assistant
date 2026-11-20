"""求和类：末尾合计行、每隔 N 行小计、合并单元格求和。"""
from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "sum_ops",
    "cat": "单元格与区域",
    "name": "特殊求和（合计/小计/合并格）",
    "desc": "在末尾追加合计行、每隔 N 行算一次小计、把合并单元格对应区间的数值求和写回左上角。",
    "fields": [
        {"key": "mode", "label": "方式", "type": "select", "default": "total",
         "options": [
             {"value": "total", "label": "末尾追加合计行"},
             {"value": "every", "label": "每隔 N 行小计"},
             {"value": "merged", "label": "合并单元格求和"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "col", "label": "合并列", "type": "text", "default": "A",
         "help": "合并单元格求和使用：按这列的合并区域分组"},
        {"key": "value_col", "label": "数值列", "type": "text", "default": "B",
         "help": "求和、小计作用的数值列；合计模式可留空=所有数值列"},
        {"key": "every", "label": "每隔 N 行", "type": "number", "default": 10},
    ],
}


def _num(v):
    try:
        if isinstance(v, bool):
            return None
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def do_total(ws, header_row, value_col):
    last = ws.max_row
    vc = batch.col_index(value_col, 0) if str(value_col or "").strip() else 0
    cols = [vc] if vc else [c for c in range(1, ws.max_column + 1)
                            if any(_num(ws.cell(row=r, column=c).value) is not None
                                   for r in range(header_row + 1, last + 1))]
    if not cols:
        return 0
    at = last + 1
    label = ws.cell(row=at, column=1)
    label.value = "合计"
    from copy import copy as _copy
    f = _copy(label.font)
    f.bold = True
    label.font = f
    for c in cols:
        total = 0.0
        found = False
        for r in range(header_row + 1, last + 1):
            v = _num(ws.cell(row=r, column=c).value)
            if v is not None:
                total += v
                found = True
        if found:
            ws.cell(row=at, column=c, value=round(total, 6))
    return 1


def do_every(ws, header_row, value_col, every):
    c = batch.col_index(value_col, 2)
    every = max(1, int(every or 10))
    out = ws.max_column + 1
    ws.cell(row=header_row, column=out, value="小计")
    n = 0
    for start in range(header_row + 1, ws.max_row + 1, every):
        total = 0.0
        for r in range(start, min(start + every, ws.max_row + 1)):
            v = _num(ws.cell(row=r, column=c).value)
            if v is not None:
                total += v
        ws.cell(row=start, column=out, value=round(total, 6))
        n += 1
    return n


def do_merged(ws, header_row, col, value_col):
    mc = batch.col_index(col, 1)
    vc = batch.col_index(value_col, 2)
    ranges = [rng for rng in ws.merged_cells.ranges if rng.min_col == mc and rng.min_row >= header_row]
    if not ranges:
        # 没有合并单元格时，退化为「相同值视为同一组」
        ranges = []
        r = header_row + 1
        from openpyxl.worksheet.cell_range import CellRange
        while r <= ws.max_row:
            v = ws.cell(row=r, column=mc).value
            end = r
            while end + 1 <= ws.max_row and ws.cell(row=end + 1, column=mc).value == v:
                end += 1
            if end > r:
                ranges.append(CellRange(min_col=mc, max_col=mc, min_row=r, max_row=end))
            r = end + 1
    n = 0
    for rng in ranges:
        total = 0.0
        found = False
        for r in range(rng.min_row, rng.max_row + 1):
            v = _num(ws.cell(row=r, column=vc).value)
            if v is not None:
                total += v
                found = True
        if found:
            ws.cell(row=rng.min_row, column=vc, value=round(total, 6))
            n += 1
    return n


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "total"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    col = params.get("col") or "A"
    value_col = params.get("value_col") or "B"
    every = int(params.get("every") or 10)

    def edit(wb, path, ctx):
        total = 0
        for ws in batch.target_sheets(wb, sheet):
            if mode == "total":
                total += do_total(ws, header_row, value_col)
            elif mode == "every":
                total += do_every(ws, header_row, value_col, every)
            elif mode == "merged":
                total += do_merged(ws, header_row, col, value_col)
        return "处理 %d 处" % total

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_求和", exts=EXCEL_EXT)
