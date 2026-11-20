"""行列批操作：间隔插入空行空列、批量插入 N 行 N 列、间隔删除、整块复制 N 次。"""
from copy import copy

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "row_col_ops",
    "cat": "数据整理",
    "name": "行列批量插入/删除",
    "desc": "每隔 N 行插空行、顶部批量插 N 行 N 列、间隔删除行列、把整块数据复制 N 次。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "blank_row",
         "options": [
             {"value": "blank_row", "label": "每隔 N 行插入空行"},
             {"value": "blank_col", "label": "每隔 N 列插入空列"},
             {"value": "insert_top_rows", "label": "顶部插入 N 行"},
             {"value": "insert_top_cols", "label": "左侧插入 N 列"},
             {"value": "delete_rows", "label": "每隔 N 行删除 M 行"},
             {"value": "delete_cols", "label": "每隔 N 列删除 M 列"},
             {"value": "copy_n", "label": "整块复制 N 次"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "留空=全部工作表"},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "every", "label": "间隔 N", "type": "number", "default": 1},
        {"key": "amount", "label": "插入/删除 M", "type": "number", "default": 1},
        {"key": "times", "label": "复制次数", "type": "number", "default": 2},
    ],
}


def _style_row(ws, dst_row, src_row):
    for c in range(1, ws.max_column + 1):
        s = ws.cell(row=src_row, column=c)
        if s.has_style:
            ws.cell(row=dst_row, column=c)._style = copy(s._style)


def blank_rows(ws, header_row, every):
    every = max(1, int(every or 1))
    last = ws.max_row
    points = []
    r = header_row + every
    while r < last:            # 末尾之后插入无意义，openpyxl 也不认
        points.append(r)
        r += every
    for p in reversed(points):
        ws.insert_rows(p + 1, 1)
        _style_row(ws, p + 1, p)
    return len(points)


def blank_cols(ws, every):
    every = max(1, int(every or 1))
    last = ws.max_column
    points = []
    c = every
    while c < last:
        points.append(c)
        c += every
    for p in reversed(points):
        ws.insert_cols(p + 1, 1)
    return len(points)


def top_rows(ws, amount):
    ws.insert_rows(1, amount)
    return amount


def top_cols(ws, amount):
    ws.insert_cols(1, amount)
    return amount


def del_rows(ws, header_row, every, amount):
    every, amount = max(1, int(every or 1)), max(1, int(amount or 1))
    last = ws.max_row
    blocks = []
    start = header_row + 1
    while start + every - 1 <= last:
        d0 = start + every
        d1 = min(d0 + amount - 1, last)
        if d0 <= last:
            blocks.append((d0, d1 - d0 + 1))
        start = d1 + 1
    for d0, cnt in reversed(blocks):
        ws.delete_rows(d0, cnt)
    return sum(c for _, c in blocks)


def del_cols(ws, every, amount):
    every, amount = max(1, int(every or 1)), max(1, int(amount or 1))
    last = ws.max_column
    blocks = []
    start = 1
    while start + every - 1 <= last:
        d0 = start + every
        d1 = min(d0 + amount - 1, last)
        if d0 <= last:
            blocks.append((d0, d1 - d0 + 1))
        start = d1 + 1
    for d0, cnt in reversed(blocks):
        ws.delete_cols(d0, cnt)
    return sum(c for _, c in blocks)


def copy_n(ws, header_row, times):
    times = max(1, int(times or 1))
    width = ws.max_column
    block = []
    for r in range(header_row + 1, ws.max_row + 1):
        block.append([ws.cell(row=r, column=c).value for c in range(1, width + 1)])
    if not block:
        return 0
    at = ws.max_row + 1
    n = 0
    for _ in range(times - 1):
        for row in block:
            for ci, v in enumerate(row, 1):
                ws.cell(row=at, column=ci, value=v)
            at += 1
            n += 1
    return n


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "blank_row"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    every = int(params.get("every") or 1)
    amount = int(params.get("amount") or 1)
    times = int(params.get("times") or 2)

    def edit(wb, path, ctx):
        total = 0
        for ws in batch.target_sheets(wb, sheet):
            if mode == "blank_row":
                total += blank_rows(ws, header_row, every)
            elif mode == "blank_col":
                total += blank_cols(ws, every)
            elif mode == "insert_top_rows":
                total += top_rows(ws, amount)
            elif mode == "insert_top_cols":
                total += top_cols(ws, amount)
            elif mode == "delete_rows":
                total += del_rows(ws, header_row, every, amount)
            elif mode == "delete_cols":
                total += del_cols(ws, every, amount)
            elif mode == "copy_n":
                total += copy_n(ws, header_row, times)
        return "处理 %d 行/列" % total

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_行列处理", exts=EXCEL_EXT)
