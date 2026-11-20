"""对比工具：两列找相同与不同项、两张工作表逐格对比。"""
from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "compare_ops",
    "cat": "查找与对比",
    "name": "相同与不同项对比",
    "desc": "比较两列数据找出「两边都有 / 只在左 / 只在右」，或逐格对比两张工作表输出差异清单。",
    "fields": [
        {"key": "mode", "label": "对比方式", "type": "select", "default": "columns",
         "options": [
             {"value": "columns", "label": "两列对比（相同/不同项）"},
             {"value": "sheets", "label": "两张工作表逐格对比"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "两列对比：数据所在表；留空=第一张"},
        {"key": "col_a", "label": "A 列", "type": "text", "default": "A"},
        {"key": "col_b", "label": "B 列", "type": "text", "default": "B"},
        {"key": "sheet2", "label": "对比工作表", "type": "text", "default": "",
         "help": "逐格对比：另一张表的名字"},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "ignore_case", "label": "忽略大小写", "type": "checkbox", "default": False},
    ],
}


def _key(v, ignore_case):
    if v is None:
        return ""
    s = str(v).strip()
    return s.lower() if ignore_case else s


def compare_columns(wb, ws, col_a, col_b, header_row, ignore_case):
    ca, cb = batch.col_index(col_a, 1), batch.col_index(col_b, 2)
    left, right = {}, {}
    for r in range(header_row + 1, ws.max_row + 1):
        a = _key(ws.cell(row=r, column=ca).value, ignore_case)
        b = _key(ws.cell(row=r, column=cb).value, ignore_case)
        if a:
            left.setdefault(a, r)
        if b:
            right.setdefault(b, r)
    both = [k for k in left if k in right]
    only_a = [k for k in left if k not in right]
    only_b = [k for k in right if k not in left]

    name = xu.safe_sheet_title("对比_" + ws.title, {s.title.lower() for s in wb.worksheets})
    out = wb.create_sheet(name)
    out.append(["类别", "内容", "位置"])
    for k in both:
        out.append(["两边都有", k, "左%d行 / 右%d行" % (left[k], right[k])])
    for k in only_a:
        out.append(["只在左列", k, "第%d行" % left[k]])
    for k in only_b:
        out.append(["只在右列", k, "第%d行" % right[k]])
    return len(both) + len(only_a) + len(only_b)


def compare_sheets(wb, ws1, ws2, limit=20000):
    rows = max(ws1.max_row, ws2.max_row)
    cols = max(ws1.max_column, ws2.max_column)
    diffs = []
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            v1 = ws1.cell(row=r, column=c).value
            v2 = ws2.cell(row=r, column=c).value
            if (v1 if v1 is not None else "") != (v2 if v2 is not None else ""):
                diffs.append([ws1.cell(row=r, column=c).coordinate, v1, v2])
                if len(diffs) >= limit:
                    break
        if len(diffs) >= limit:
            break
    name = xu.safe_sheet_title("差异_" + ws1.title + "_" + ws2.title,
                               {s.title.lower() for s in wb.worksheets})
    out = wb.create_sheet(name)
    out.append(["单元格", ws1.title, ws2.title])
    for d in diffs:
        out.append(d)
    return len(diffs)


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "columns"
    sheet = params.get("sheet") or ""
    col_a = params.get("col_a") or "A"
    col_b = params.get("col_b") or "B"
    sheet2 = params.get("sheet2") or ""
    header_row = int(params.get("header_row") or 1)
    ic = bool(params.get("ignore_case"))

    if mode == "sheets" and not sheet2:
        ctx.error("逐格对比需要填「对比工作表」的名字")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    def edit(wb, path, ctx):
        if mode == "columns":
            ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
            return "对比出 %d 条结果" % compare_columns(wb, ws, col_a, col_b, header_row, ic)
        if sheet2 not in wb.sheetnames:
            raise ValueError("工作簿里没有名为 %s 的工作表" % sheet2)
        ws1 = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
        return "找出 %d 处差异" % compare_sheets(wb, ws1, wb[sheet2])

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_对比", exts=EXCEL_EXT)
