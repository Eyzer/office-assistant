"""数字格式转换：文本数字↔数值、补零、中文大写金额、千分位。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "num_format",
    "name": "数字格式转换",
    "desc": "文本型数字转数值、数值转文本、整数补零、转中文大写金额、加千分位。",
    "fields": [
        {"key": "mode", "label": "转换方式", "type": "select", "default": "text2num",
         "options": [
             {"value": "text2num", "label": "文本数字 → 数值"},
             {"value": "num2text", "label": "数值 → 文本"},
             {"value": "pad_zero", "label": "整数补零"},
             {"value": "rmb_upper", "label": "金额 → 中文大写"},
             {"value": "thousand", "label": "加千分位"},
         ]},
        {"key": "col", "label": "作用列", "type": "text", "default": "",
         "help": "填列字母，留空=全部列"},
        {"key": "pad_len", "label": "补零位数", "type": "number", "default": 6},
    ],
}

_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_UNITS = ["", "拾", "佰", "仟"]
_BIG = ["", "万", "亿", "兆"]


def rmb_upper(num):
    """金额转中文大写。"""
    try:
        num = round(float(num), 2)
    except (TypeError, ValueError):
        return ""
    if num == 0:
        return "零元整"
    neg = num < 0
    num = abs(num)
    integer = int(num)
    frac = int(round((num - integer) * 100))

    def four(n):
        s, zero = "", False
        for i in range(3, -1, -1):
            d = (n // (10 ** i)) % 10
            if d == 0:
                zero = True
            else:
                if zero and s:
                    s += "零"
                s += _DIGITS[d] + _UNITS[i]
                zero = False
        return s

    parts, gi, out = [], 0, ""
    n = integer
    while n > 0:
        parts.append(four(n % 10000))
        n //= 10000
        gi += 1
    for idx in range(len(parts) - 1, -1, -1):
        seg = parts[idx]
        if seg:
            out += seg + _BIG[idx]
        elif out and not out.endswith("零"):
            out += "零"
    out = out.rstrip("零") or "零"
    res = out + "元"
    if frac == 0:
        res += "整"
    else:
        jiao, fen = frac // 10, frac % 10
        if jiao:
            res += _DIGITS[jiao] + "角"
        elif fen:
            res += "零"
        if fen:
            res += _DIGITS[fen] + "分"
    return ("负" if neg else "") + res


def _target_cols(ws, col):
    if not col:
        return list(range(1, ws.max_column + 1))
    if col.isalpha():
        return [__import__("openpyxl").utils.column_index_from_string(col.upper())]
    return list(range(1, ws.max_column + 1))


def _apply(ws, mode, col, pad_len):
    n = 0
    for ci in _target_cols(ws, col):
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(row=r, column=ci)
            v = cell.value
            if v in (None, ""):
                continue
            try:
                if mode == "text2num":
                    cell.value = float(str(v).replace(",", "").strip())
                    if cell.value.is_integer():
                        cell.value = int(cell.value)
                    cell.number_format = "General"
                elif mode == "num2text":
                    cell.value = str(v)
                elif mode == "pad_zero":
                    cell.value = str(int(float(v))).zfill(pad_len)
                elif mode == "rmb_upper":
                    cell.value = rmb_upper(v)
                elif mode == "thousand":
                    f = float(str(v).replace(",", ""))
                    cell.value = f
                    cell.number_format = "#,##0.00"
                n += 1
            except (ValueError, TypeError):
                continue
    return n


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    mode = params.get("mode") or "text2num"
    col = (params.get("col") or "").strip()
    pad_len = int(params.get("pad_len") or 6)
    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            cnt = sum(_apply(ws, mode, col, pad_len) for ws in wb.worksheets)
            dst = out_path(f, output_dir, suffix="_数字格式")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 处理 {cnt} 个单元格")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
