"""数值处理：四舍五入、四舍六入五留单、取整、改正负、按区间分级。"""
from decimal import Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN, ROUND_CEILING, ROUND_FLOOR

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "round_ops",
    "cat": "数值与计算",
    "name": "数值取整与分级",
    "desc": "四舍五入、四舍六入五留单（银行家舍入）、向上/向下取整、改数字正负、按区间划分级别。",
    "fields": [
        {"key": "mode", "label": "处理方式", "type": "select", "default": "half_up",
         "options": [
             {"value": "half_up", "label": "四舍五入"},
             {"value": "half_even", "label": "四舍六入五留单"},
             {"value": "ceil", "label": "向上取整"},
             {"value": "floor", "label": "向下取整"},
             {"value": "abs", "label": "取绝对值"},
             {"value": "negate", "label": "正负号取反"},
             {"value": "level", "label": "按区间分级"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "col", "label": "作用列", "type": "text", "default": "A"},
        {"key": "digits", "label": "保留位数", "type": "number", "default": 2},
        {"key": "rules", "label": "分级规则", "type": "text", "default": "0-60:不及格,60-80:及格,80-100:优秀",
         "help": "格式 下限-上限:名称，逗号分隔"},
    ],
}


def _num(v):
    if isinstance(v, bool):
        return None
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def parse_rules(text):
    out = []
    for part in str(text or "").replace("，", ",").split(","):
        part = part.strip()
        if not part or ":" not in part:
            continue
        rng, name = part.split(":", 1)
        rng = rng.replace("～", "-").replace("~", "-")
        if "-" not in rng:
            continue
        lo, hi = rng.split("-", 1)
        try:
            out.append((float(lo), float(hi), name.strip()))
        except ValueError:
            continue
    return out


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "half_up"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    col = params.get("col") or "A"
    digits = int(params.get("digits") or 0)
    rules = parse_rules(params.get("rules"))
    if mode == "level" and not rules:
        ctx.error("分级规则没解析出来，示例：0-60:不及格,60-80:及格,80-100:优秀")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    def edit(wb, path, ctx):
        n = 0
        for ws in batch.target_sheets(wb, sheet):
            c = batch.col_index(col, 1)
            for r in range(header_row + 1, ws.max_row + 1):
                cell = ws.cell(row=r, column=c)
                v = _num(cell.value)
                if v is None:
                    continue
                if mode == "level":
                    for lo, hi, name in rules:
                        if lo <= v <= hi:
                            cell.value = name
                            break
                    n += 1
                    continue
                d = Decimal(str(v))
                if mode == "half_up":
                    nv = d.quantize(Decimal(10) ** -digits, rounding=ROUND_HALF_UP)
                elif mode == "half_even":
                    nv = d.quantize(Decimal(10) ** -digits, rounding=ROUND_HALF_EVEN)
                elif mode == "ceil":
                    nv = d.to_integral_value(rounding=ROUND_CEILING)
                elif mode == "floor":
                    nv = d.to_integral_value(rounding=ROUND_FLOOR)
                elif mode == "abs":
                    nv = abs(d)
                elif mode == "negate":
                    nv = -d
                else:
                    continue
                f = float(nv)
                cell.value = int(f) if f.is_integer() and digits == 0 else f
                n += 1
        return "处理 %d 个数值" % n

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_数值处理", exts=EXCEL_EXT)
