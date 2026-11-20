"""分类汇总：按分组列对数值列求和/计数/平均，输出汇总表或在原表插入小计行。"""
from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "subtotal_ops",
    "cat": "数据整理",
    "name": "分类汇总",
    "desc": "按某一列分组，对指定数值列做求和/计数/平均/最大/最小，生成汇总表或在原表插入小计行。",
    "fields": [
        {"key": "group_col", "label": "分组列", "type": "text", "default": "A",
         "help": "如「部门」「地区」"},
        {"key": "value_cols", "label": "数值列", "type": "text", "default": "",
         "help": "逗号分隔列字母，留空=所有数值列"},
        {"key": "func", "label": "汇总方式", "type": "select", "default": "sum",
         "options": [{"value": "sum", "label": "求和"},
                     {"value": "count", "label": "计数"},
                     {"value": "avg", "label": "平均值"},
                     {"value": "max", "label": "最大值"},
                     {"value": "min", "label": "最小值"}]},
        {"key": "inline", "label": "在原表插入小计行", "type": "checkbox", "default": False,
         "help": "不勾=生成汇总表（推荐，不动原数据布局）"},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
    ],
}

_FUNC_CN = {"sum": "求和", "count": "计数", "avg": "平均", "max": "最大", "min": "最小"}


def _num(v):
    try:
        if isinstance(v, bool):
            return None
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _agg(values, func):
    if func == "count":
        return len(values)
    if not values:
        return 0
    if func == "sum":
        return round(sum(values), 6)
    if func == "avg":
        return round(sum(values) / len(values), 6)
    if func == "max":
        return max(values)
    if func == "min":
        return min(values)
    return 0


def _value_cols(ws, header_row, spec):
    spec = str(spec or "").replace("，", ",").strip()
    if spec:
        return [batch.col_index(x, 1) for x in spec.split(",") if x.strip()]
    out = []
    for c in range(1, ws.max_column + 1):
        if any(_num(ws.cell(row=r, column=c).value) is not None
               for r in range(header_row + 1, ws.max_row + 1)):
            out.append(c)
    return out


def build_summary(wb, ws, group_col, value_cols, func, header_row):
    gc = batch.col_index(group_col, 1)
    cols = _value_cols(ws, header_row, value_cols)
    order, data = [], {}
    for r in range(header_row + 1, ws.max_row + 1):
        k = ws.cell(row=r, column=gc).value
        if k in (None, ""):
            continue
        k = str(k).strip()
        if k not in data:
            data[k] = {c: [] for c in cols}
            order.append(k)
        for c in cols:
            v = _num(ws.cell(row=r, column=c).value)
            if v is not None:
                data[k][c].append(v)

    name = xu.safe_sheet_title("汇总_" + ws.title, {s.title.lower() for s in wb.worksheets})
    out = wb.create_sheet(name)
    head = [ws.cell(row=header_row, column=gc).value or "分组"]
    for c in cols:
        h = ws.cell(row=header_row, column=c).value or ("列%d" % c)
        head.append("%s(%s)" % (h, _FUNC_CN.get(func, func)))
    out.append(head)

    totals = {c: [] for c in cols}
    for k in order:
        row = [k]
        for c in cols:
            vals = data[k][c]
            totals[c].extend(vals)
            row.append(_agg(vals, func))
        out.append(row)
    out.append(["合计"] + [_agg(totals[c], func if func != "count" else "sum") for c in cols])
    return len(order)


def inline_subtotal(wb, ws, group_col, value_cols, func, header_row):
    gc = batch.col_index(group_col, 1)
    cols = _value_cols(ws, header_row, value_cols)
    width = ws.max_column
    header = [ws.cell(row=header_row, column=c).value for c in range(1, width + 1)]
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        rows.append([ws.cell(row=r, column=c).value for c in range(1, width + 1)])
    if not rows:
        return 0

    grouped, order = {}, []
    for row in rows:
        k = row[gc - 1]
        k = "" if k is None else str(k).strip()
        if k not in grouped:
            grouped[k] = []
            order.append(k)
        grouped[k].append(row)

    new_rows = []
    for k in order:
        new_rows.extend(grouped[k])
        sub = [None] * width
        sub[gc - 1] = "%s 小计" % k
        for c in cols:
            vals = [x for x in (_num(r[c - 1]) for r in grouped[k]) if x is not None]
            sub[c - 1] = _agg(vals, func)
        new_rows.append(sub)

    ws.delete_rows(header_row + 1, ws.max_row - header_row)
    for i, row in enumerate(new_rows, header_row + 1):
        for c in range(1, width + 1):
            ws.cell(row=i, column=c, value=row[c - 1])
    return len(order)


def run_task(inputs, output_dir, params, ctx):
    group_col = params.get("group_col") or "A"
    value_cols = params.get("value_cols") or ""
    func = params.get("func") or "sum"
    inline = bool(params.get("inline"))
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)

    def edit(wb, path, ctx):
        total = 0
        for ws in batch.target_sheets(wb, sheet):
            if inline:
                total += inline_subtotal(wb, ws, group_col, value_cols, func, header_row)
            else:
                total += build_summary(wb, ws, group_col, value_cols, func, header_row)
        return "汇总 %d 组" % total

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_分类汇总", exts=EXCEL_EXT)
