"""文字转换：汉字转拼音、数字转英文/英文金额、大写金额转数字、表达式求值、进制与 IP 转换。"""
import re

from utils import batch, deps, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "text_convert",
    "cat": "文本处理",
    "name": "文字数值互转",
    "desc": "汉字转拼音（全拼/首字母）、数字转英文与英文金额、大写金额转数字、算式求值、进制与 IP 互转。",
    "fields": [
        {"key": "mode", "label": "转换方式", "type": "select", "default": "pinyin",
         "options": [
             {"value": "pinyin", "label": "汉字转拼音"},
             {"value": "pinyin_first", "label": "汉字转拼音首字母"},
             {"value": "num2en", "label": "数字转英文"},
             {"value": "num2en_money", "label": "数字转英文金额"},
             {"value": "rmb2num", "label": "中文大写金额转数字"},
             {"value": "expr", "label": "算式求值"},
             {"value": "base", "label": "进制转换"},
             {"value": "ip2num", "label": "IP 转数值"},
             {"value": "num2ip", "label": "数值转 IP"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "col", "label": "作用列", "type": "text", "default": "A"},
        {"key": "sep", "label": "拼音分隔", "type": "text", "default": " ",
         "help": "留空=不分隔"},
        {"key": "from_base", "label": "源进制", "type": "number", "default": 10},
        {"key": "to_base", "label": "目标进制", "type": "number", "default": 16},
    ],
}

_ONES = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
         "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
         "seventeen", "eighteen", "nineteen"]
_TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]


def _below_thousand(n):
    out = []
    if n >= 100:
        out.append(_ONES[n // 100] + " hundred")
        n %= 100
        if n:
            out.append("and")
    if n >= 20:
        t, o = divmod(n, 10)
        out.append(_TENS[t] + ("-" + _ONES[o] if o else ""))
    elif n:
        out.append(_ONES[n])
    return " ".join(out)


def _to_float(v):
    try:
        return float(str(v).replace(",", "").replace("，", "").strip())
    except (TypeError, ValueError):
        return None


def num2en(num):
    n = _to_float(num)
    if n is None:
        return ""
    neg = n < 0
    n = abs(n)
    whole = int(n)
    frac = round(n - whole, 4)
    if whole == 0:
        words = "zero"
    else:
        parts = []
        for div, name in ((10 ** 9, "billion"), (10 ** 6, "million"), (1000, "thousand")):
            if whole >= div:
                parts.append(_below_thousand(whole // div) + " " + name)
                whole %= div
        if whole:
            parts.append(_below_thousand(whole))
        words = " ".join(parts)
    if frac:
        digits = ("%.4f" % frac).split(".")[1].rstrip("0")
        words += " point " + " ".join(_ONES[int(d)] if d != "0" else "zero" for d in digits)
    return ("minus " if neg else "") + words


def num2en_money(num):
    n = _to_float(num)
    if n is None:
        return ""
    n = round(n, 2)
    whole = int(abs(n))
    cents = int(round((abs(n) - whole) * 100))
    s = num2en(whole).upper() + " DOLLARS"
    if cents:
        s += " AND " + num2en(cents).upper() + " CENTS"
    return ("MINUS " if n < 0 else "") + s


_RMB_D = {"零": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5, "陆": 6, "柒": 7,
          "捌": 8, "玖": 9, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6,
          "七": 7, "八": 8, "九": 9, "两": 2}
_RMB_U = {"拾": 10, "十": 10, "佰": 100, "百": 100, "仟": 1000, "千": 1000}
_RMB_BIG = {"万": 10 ** 4, "亿": 10 ** 8}


def rmb2num(text):
    """中文大写金额转数字，如 壹佰贰拾叁元肆角伍分 → 123.45。"""
    s = str(text or "").strip()
    if not s:
        return None
    s = s.replace("整", "").replace("正", "").replace("圆", "元")
    yuan, _, rest = s.partition("元")
    if not _:
        rest = ""
        yuan = s

    def parse_int(part):
        total, section, num = 0, 0, 0
        for ch in part:
            if ch in _RMB_D:
                num = _RMB_D[ch]
            elif ch in _RMB_U:
                section += (num or 1) * _RMB_U[ch]
                num = 0
            elif ch in _RMB_BIG:
                section += num
                total += section * _RMB_BIG[ch]
                section, num = 0, 0
            elif ch in "，, ":
                continue
        return total + section + num

    value = float(parse_int(yuan))
    jiao = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九两])角", rest)
    fen = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九两])分", rest)
    if jiao:
        value += _RMB_D[jiao.group(1)] * 0.1
    if fen:
        value += _RMB_D[fen.group(1)] * 0.01
    return round(value, 2)


def safe_eval(expr):
    s = str(expr or "")
    s = (s.replace("×", "*").replace("÷", "/").replace("^", "**")
          .replace("（", "(").replace("）", ")").replace("，", ",").replace("＝", "="))
    s = s.lstrip("=")
    s = re.sub(r"[^0-9+\-*/().% ]", "", s)
    if not s.strip():
        return None
    try:
        return eval(compile(s, "<expr>", "eval"), {"__builtins__": {}}, {})
    except Exception:
        return None


def to_base(v, frm, to):
    try:
        n = int(str(v).strip(), int(frm))
    except (TypeError, ValueError):
        return None
    if not (2 <= int(to) <= 36):
        return None
    digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if n == 0:
        return "0"
    neg = n < 0
    n = abs(n)
    out = ""
    while n:
        n, r = divmod(n, int(to))
        out = digits[r] + out
    return ("-" if neg else "") + out


def ip2int(v):
    parts = str(v).strip().split(".")
    if len(parts) != 4:
        return None
    try:
        nums = [int(p) for p in parts]
    except ValueError:
        return None
    if any(x < 0 or x > 255 for x in nums):
        return None
    return (nums[0] << 24) + (nums[1] << 16) + (nums[2] << 8) + nums[3]


def int2ip(v):
    try:
        n = int(float(v))
    except (TypeError, ValueError):
        return None
    if not (0 <= n <= 0xFFFFFFFF):
        return None
    return ".".join(str((n >> s) & 255) for s in (24, 16, 8, 0))


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "pinyin"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    col = params.get("col") or "A"
    sep = params.get("sep")
    sep = " " if sep is None else str(sep)
    frm = int(params.get("from_base") or 10)
    to = int(params.get("to_base") or 16)

    if mode.startswith("pinyin"):
        if not deps.need(ctx, ["pypinyin"]):
            return {"success_count": 0, "fail_count": 0,
                    "output_files": [], "log_messages": []}
        from pypinyin import lazy_pinyin, Style

    def convert(v):
        if mode.startswith("pinyin"):
            style = Style.FIRST_LETTER if mode == "pinyin_first" else Style.NORMAL
            return sep.join(lazy_pinyin(str(v), style=style))
        if mode == "num2en":
            return num2en(v)
        if mode == "num2en_money":
            return num2en_money(v)
        if mode == "rmb2num":
            return rmb2num(v)
        if mode == "expr":
            return safe_eval(v)
        if mode == "base":
            return to_base(v, frm, to)
        if mode == "ip2num":
            return ip2int(v)
        if mode == "num2ip":
            return int2ip(v)
        return v

    def edit(wb, path, ctx):
        n = 0
        for ws in batch.target_sheets(wb, sheet):
            c = batch.col_index(col, 1)
            out = ws.max_column + 1
            ws.cell(row=header_row, column=out, value="转换结果")
            for r in range(header_row + 1, ws.max_row + 1):
                v = ws.cell(row=r, column=c).value
                if v in (None, ""):
                    continue
                try:
                    res = convert(v)
                except Exception:
                    continue
                if res is None or res == "":
                    continue
                cell = ws.cell(row=r, column=out, value=res)
                if not isinstance(res, (int, float)):
                    cell.number_format = "@"
                n += 1
        return "转换 %d 行" % n

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_文字转换", exts=EXCEL_EXT)
