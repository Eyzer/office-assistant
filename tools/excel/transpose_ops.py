"""区域转换：行列转置、一维转二维、单行单列转多行多列、交换两个区域。"""
import os
from copy import copy

from openpyxl.utils import get_column_letter, range_boundaries

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "transpose_ops",
    "cat": "数据整理",
    "name": "区域转换（转置/一维转二维）",
    "desc": "行列转置、一维明细转二维交叉表、单列切成多行多列、交换两个区域，结果写新工作表。",
    "fields": [
        {"key": "mode", "label": "转换方式", "type": "select", "default": "transpose",
         "options": [
             {"value": "transpose", "label": "行列转置"},
             {"value": "one2two", "label": "一维明细 → 二维交叉表"},
             {"value": "single2multi", "label": "单行/单列 → 多行多列"},
             {"value": "swap", "label": "交换两个区域"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "留空=全部工作表"},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "key_col", "label": "行字段（列）", "type": "text", "default": "A",
         "help": "一维转二维：作为行标题的列"},
        {"key": "item_col", "label": "列字段（列）", "type": "text", "default": "B",
         "help": "一维转二维：作为列标题的列"},
        {"key": "value_col", "label": "值字段（列）", "type": "text", "default": "C",
         "help": "一维转二维：填入交叉点的值"},
        {"key": "cols_per_row", "label": "每组列数", "type": "number", "default": 5,
         "help": "单行单列转多行多列：每行放几个值"},
        {"key": "range1", "label": "区域1", "type": "text", "default": "",
         "help": "交换区域：如 A1:C3"},
        {"key": "range2", "label": "区域2", "type": "text", "default": "",
         "help": "交换区域：如 E1:G3，大小需与区域1一致"},
    ],
}


def _trim(rows):
    """去掉尾部全空的行与列。"""
    rows = [list(r) for r in rows]
    while rows and all(v is None or v == "" for v in rows[-1]):
        rows.pop()
    if not rows:
        return []
    width = 0
    for r in rows:
        for i, v in enumerate(r):
            if v is not None and v != "":
                width = max(width, i + 1)
    return [r[:width] for r in rows]


def _new_sheet(wb, base, title):
    name = xu.safe_sheet_title(title, {s.title.lower() for s in wb.worksheets})
    return wb.create_sheet(name)


def do_transpose(wb, ws):
    rows = _trim([[c.value for c in r] for r in ws.iter_rows()])
    if not rows:
        return 0
    out = _new_sheet(wb, ws.title, "转置_" + ws.title)
    data = [list(r) for r in zip(*rows)]
    for ri, row in enumerate(data, 1):
        for ci, v in enumerate(row, 1):
            out.cell(row=ri, column=ci, value=v)
    return 1


def do_one2two(wb, ws, header_row, key, item, value):
    ck, ci, cv = batch.col_index(key, 1), batch.col_index(item, 2), batch.col_index(value, 3)
    rows = list(ws.iter_rows(min_row=header_row + 1, values_only=True))
    rows = [r for r in rows if any(v is not None and v != "" for v in r)]
    if not rows:
        return 0
    keys, items, cell = [], [], {}
    for r in rows:
        k = r[ck - 1] if len(r) >= ck else None
        i = r[ci - 1] if len(r) >= ci else None
        v = r[cv - 1] if len(r) >= cv else None
        if k is None:
            continue
        k = str(k).strip()
        i = "" if i is None else str(i).strip()
        if k not in keys:
            keys.append(k)
        if i not in items:
            items.append(i)
        cell[(k, i)] = v
    out = _new_sheet(wb, ws.title, "二维_" + ws.title)
    out.cell(row=1, column=1, value=ws.cell(row=header_row, column=ck).value or "行\\列")
    for j, it in enumerate(items, 2):
        out.cell(row=1, column=j, value=it)
    for i, k in enumerate(keys, 2):
        out.cell(row=i, column=1, value=k)
        for j, it in enumerate(items, 2):
            out.cell(row=i, column=j, value=cell.get((k, it)))
    return 1


def do_single2multi(wb, ws, n):
    n = max(1, int(n or 5))
    vals = []
    for r in ws.iter_rows(values_only=True):
        for v in r:
            if v is not None and v != "":
                vals.append(v)
    if not vals:
        return 0
    out = _new_sheet(wb, ws.title, "多行多列_" + ws.title)
    for idx, v in enumerate(vals):
        out.cell(row=idx // n + 1, column=idx % n + 1, value=v)
    return 1


def do_swap(ws, r1, r2):
    try:
        a1, b1, a2, b2 = range_boundaries(r1.replace("$", "").upper())
        c1, d1, c2, d2 = range_boundaries(r2.replace("$", "").upper())
    except (TypeError, ValueError):
        raise ValueError("区域写法不对，示例：A1:C3 与 E1:G3")
    if (a2 - a1, b2 - b1) != (c2 - c1, d2 - d1):
        raise ValueError("两个区域的行列数必须一致")
    # 先快照两边的值与样式，避免写回时互相覆盖
    buf1 = [[(ws.cell(row=r, column=c).value, ws.cell(row=r, column=c).has_style,
              ws.cell(row=r, column=c)._style) for c in range(a1, a2 + 1)]
            for r in range(b1, b2 + 1)]
    buf2 = [[(ws.cell(row=r, column=c).value, ws.cell(row=r, column=c).has_style,
              ws.cell(row=r, column=c)._style) for c in range(c1, c2 + 1)]
            for r in range(d1, d2 + 1)]
    for i, row in enumerate(buf1):
        for j, (val, has, style) in enumerate(row):
            src = buf2[i][j]
            cell = ws.cell(row=b1 + i, column=a1 + j)
            cell.value = src[0]
            if src[1]:
                cell._style = copy(src[2])
    for i, row in enumerate(buf2):
        for j, (val, has, style) in enumerate(row):
            src = buf1[i][j]
            cell = ws.cell(row=d1 + i, column=c1 + j)
            cell.value = src[0]
            if src[1]:
                cell._style = copy(src[2])
    return 1


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "transpose"
    sheet = params.get("sheet")
    header_row = int(params.get("header_row") or 1)
    n = int(params.get("cols_per_row") or 5)
    key = params.get("key_col") or "A"
    item = params.get("item_col") or "B"
    value = params.get("value_col") or "C"
    r1, r2 = (params.get("range1") or "").strip(), (params.get("range2") or "").strip()

    if mode == "swap" and not (r1 and r2):
        ctx.error("交换区域需要填两个区域，如 A1:C3 与 E1:G3")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    def edit(wb, path, ctx):
        cnt = 0
        for ws in batch.target_sheets(wb, sheet):
            if mode == "transpose":
                cnt += do_transpose(wb, ws)
            elif mode == "one2two":
                cnt += do_one2two(wb, ws, header_row, key, item, value)
            elif mode == "single2multi":
                cnt += do_single2multi(wb, ws, n)
            elif mode == "swap":
                cnt += do_swap(ws, r1, r2)
        return "处理 %d 个工作表" % cnt

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_区域转换", exts=EXCEL_EXT)
