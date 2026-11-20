"""层级处理：编号累加、单列拆解、多列合并、层级提取、层级汇总、层级比较。"""
from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "hierarchy_ops",
    "cat": "数据整理",
    "name": "层级处理（编号/拆解/合并）",
    "desc": "按层级列生成 1、1.1、1.1.1 这类累加编号；也能把「省/市/区」拆成多列或反向合并。",
    "fields": [
        {"key": "mode", "label": "处理方式", "type": "select", "default": "number",
         "options": [
             {"value": "number", "label": "层级编号累加"},
             {"value": "split", "label": "单列层级拆解成多列"},
             {"value": "merge", "label": "多列层级合并成一列"},
             {"value": "extract", "label": "提取第 N 层"},
             {"value": "summary", "label": "层级汇总（分组求和）"},
             {"value": "compare", "label": "层级比较（两列是否一致）"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "留空=全部工作表"},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "level_col", "label": "层级列", "type": "text", "default": "A",
         "help": "里面放 1/2/3 或 一级/二级/三级"},
        {"key": "sep", "label": "分隔符", "type": "text", "default": "/"},
        {"key": "col_start", "label": "起始列", "type": "text", "default": "A"},
        {"key": "col_count", "label": "列数", "type": "number", "default": 3},
        {"key": "n", "label": "第几层", "type": "number", "default": 1},
        {"key": "value_col", "label": "汇总值列", "type": "text", "default": ""},
    ],
}

_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def parse_level(v):
    """把 1 / "1" / "一级" / "第2层" 统一成整数层级。"""
    if v is None or v == "":
        return 1
    if isinstance(v, (int, float)):
        return max(1, int(v))
    s = str(v).strip().replace("第", "").replace("层", "").replace("级", "")
    if s.isdigit():
        return max(1, int(s))
    for k, num in _CN.items():
        if s.startswith(k):
            return num
    return 1


def _out_col(ws, header, name):
    c = ws.max_column + 1
    if ws.cell(row=1, column=1).value is not None or ws.max_row > 1:
        ws.cell(row=1, column=c, value=name)
    return c


def do_number(ws, header_row, level_col):
    lc = batch.col_index(level_col, 1)
    out = ws.max_column + 1
    ws.cell(row=header_row, column=out, value="层级编号")
    counts = [0] * 12
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=lc).value
        if v is None or v == "":
            continue
        lv = min(parse_level(v), 10)
        counts[lv - 1] += 1
        for i in range(lv, 12):
            counts[i] = 0
        code = ".".join(str(counts[i]) for i in range(lv))
        ws.cell(row=r, column=out, value=code)
        n += 1
    return n


def do_split(ws, header_row, level_col, sep):
    lc = batch.col_index(level_col, 1)
    max_parts = 1
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=lc).value
        parts = [] if v in (None, "") else [p.strip() for p in str(v).split(sep or "/")]
        max_parts = max(max_parts, len(parts))
        rows.append((r, parts))
    base = ws.max_column
    for i in range(max_parts):
        ws.cell(row=header_row, column=base + i + 1, value="层级%d" % (i + 1))
    for r, parts in rows:
        for i, p in enumerate(parts):
            ws.cell(row=r, column=base + i + 1, value=p)
    return len(rows)


def do_merge(ws, header_row, col_start, count, sep):
    c0 = batch.col_index(col_start, 1)
    out = ws.max_column + 1
    ws.cell(row=header_row, column=out, value="合并层级")
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        parts = []
        for c in range(c0, c0 + max(1, count)):
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                parts.append(str(v).strip())
        if parts:
            ws.cell(row=r, column=out, value=(sep or "/").join(parts))
            n += 1
    return n


def do_extract(ws, header_row, level_col, sep, n_layer):
    lc = batch.col_index(level_col, 1)
    out = ws.max_column + 1
    ws.cell(row=header_row, column=out, value="第%d层" % n_layer)
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=lc).value
        if v in (None, ""):
            continue
        parts = [p.strip() for p in str(v).split(sep or "/")]
        idx = max(1, int(n_layer)) - 1
        if idx < len(parts):
            ws.cell(row=r, column=out, value=parts[idx])
            n += 1
    return n


def do_summary(wb, ws, header_row, level_col, value_col, src_title):
    lc = batch.col_index(level_col, 1)
    vc = batch.col_index(value_col, 0)
    groups, order = {}, []
    sums = {}
    for r in range(header_row + 1, ws.max_row + 1):
        k = ws.cell(row=r, column=lc).value
        if k in (None, ""):
            continue
        k = str(k).strip()
        if k not in groups:
            groups[k] = 0
            order.append(k)
        groups[k] += 1
        if vc:
            v = ws.cell(row=r, column=vc).value
            try:
                sums[k] = sums.get(k, 0) + float(v)
            except (TypeError, ValueError):
                sums.setdefault(k, sums.get(k, 0))
    name = xu.safe_sheet_title("层级汇总_" + src_title, {s.title.lower() for s in wb.worksheets})
    out = wb.create_sheet(name)
    out.cell(row=1, column=1, value="层级")
    out.cell(row=1, column=2, value="条数")
    out.cell(row=1, column=3, value="合计")
    for i, k in enumerate(order, 2):
        out.cell(row=i, column=1, value=k)
        out.cell(row=i, column=2, value=groups[k])
        out.cell(row=i, column=3, value=round(sums.get(k, 0), 6))
    return len(order)


def do_compare(wb, ws, header_row, col_a, col_b, src_title):
    ca, cb = batch.col_index(col_a, 1), batch.col_index(col_b, 2)
    diffs = []
    for r in range(header_row + 1, ws.max_row + 1):
        a = ws.cell(row=r, column=ca).value
        b = ws.cell(row=r, column=cb).value
        if ("" if a is None else str(a).strip()) != ("" if b is None else str(b).strip()):
            diffs.append((r, a, b))
    if not diffs:
        return 0
    name = xu.safe_sheet_title("层级差异_" + src_title, {s.title.lower() for s in wb.worksheets})
    out = wb.create_sheet(name)
    out.append(["行号", "第1列", "第2列"])
    for r, a, b in diffs:
        out.append([r, a, b])
    return len(diffs)


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "number"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    level_col = params.get("level_col") or "A"
    sep = params.get("sep") or "/"
    col_start = params.get("col_start") or "A"
    count = int(params.get("col_count") or 3)
    n_layer = int(params.get("n") or 1)
    value_col = params.get("value_col") or ""

    def edit(wb, path, ctx):
        total = 0
        for ws in batch.target_sheets(wb, sheet):
            if mode == "number":
                total += do_number(ws, header_row, level_col)
            elif mode == "split":
                total += do_split(ws, header_row, level_col, sep)
            elif mode == "merge":
                total += do_merge(ws, header_row, col_start, count, sep)
            elif mode == "extract":
                total += do_extract(ws, header_row, level_col, sep, n_layer)
            elif mode == "summary":
                total += do_summary(wb, ws, header_row, level_col, value_col, ws.title)
            elif mode == "compare":
                total += do_compare(wb, ws, header_row, level_col, value_col, ws.title)
        return "处理 %d 行/组" % total

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_层级处理", exts=EXCEL_EXT)
