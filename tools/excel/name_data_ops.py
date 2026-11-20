"""证件与号码数据：身份证解析/校验、脱敏、统一社会信用代码校验。"""
import datetime

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "name_data_ops",
    "cat": "数值与计算",
    "name": "身份证/手机号数据处理",
    "desc": "从身份证提取生日、性别、年龄并校验真伪；手机号/银行卡/姓名/身份证批量脱敏；统一社会信用代码校验。",
    "fields": [
        {"key": "mode", "label": "处理方式", "type": "select", "default": "idcard",
         "options": [
             {"value": "idcard", "label": "身份证解析 + 校验"},
             {"value": "mask", "label": "批量脱敏"},
             {"value": "credit", "label": "统一社会信用代码校验"},
         ]},
        {"key": "col", "label": "数据列", "type": "text", "default": "A"},
        {"key": "kind", "label": "脱敏类型", "type": "select", "default": "phone",
         "options": [{"value": "phone", "label": "手机号"},
                     {"value": "idcard", "label": "身份证"},
                     {"value": "bank", "label": "银行卡"},
                     {"value": "name", "label": "姓名"},
                     {"value": "email", "label": "邮箱"}]},
        {"key": "head", "label": "保留前几位", "type": "number", "default": 0,
         "help": "0=用该类型的默认规则"},
        {"key": "tail", "label": "保留后几位", "type": "number", "default": 0},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
    ],
}

_ID_W = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
_ID_CHECK = "10X98765432"
_CREDIT_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
_CREDIT_W = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]

_DEFAULT_MASK = {"phone": (3, 4), "idcard": (6, 4), "bank": (4, 4), "name": (1, 0), "email": (2, 0)}


def id_check(s):
    s = str(s or "").strip().upper()
    if len(s) != 18 or not s[:17].isdigit():
        return False
    total = sum(int(s[i]) * _ID_W[i] for i in range(17))
    return s[17] == _ID_CHECK[total % 11]


def id_info(s):
    s = str(s or "").strip().upper()
    if len(s) != 18:
        return None
    try:
        birth = datetime.date(int(s[6:10]), int(s[10:12]), int(s[12:14]))
    except ValueError:
        return None
    today = datetime.date.today()
    age = today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))
    return {"birth": birth.isoformat(), "gender": "男" if int(s[16]) % 2 else "女",
            "age": age, "ok": id_check(s), "area": s[:6]}


def credit_check(s):
    s = str(s or "").strip().upper()
    if len(s) != 18:
        return False
    try:
        total = sum(_CREDIT_CHARS.index(ch) * _CREDIT_W[i] for i, ch in enumerate(s[:17]))
    except ValueError:
        return False
    rem = 31 - total % 31
    if rem == 31:
        rem = 0
    return _CREDIT_CHARS[rem] == s[17]


def mask_value(v, kind, head, tail):
    s = str(v or "").strip()
    if not s:
        return s
    h, t = _DEFAULT_MASK.get(kind, (0, 0))
    if head:
        h = head
    if tail:
        t = tail
    if kind == "email" and "@" in s:
        user, _, dom = s.partition("@")
        keep = max(1, h or 1)
        return user[:keep] + "***@" + dom
    if kind == "name":
        if len(s) <= 1:
            return s
        return s[0] + "*" * (len(s) - 1)
    if len(s) <= h + t:
        return s[0] + "***"
    return s[:h] + "*" * (len(s) - h - t) + (s[-t:] if t else "")


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "idcard"
    col = batch.col_index(params.get("col") or "A", 1)
    kind = params.get("kind") or "phone"
    head = int(params.get("head") or 0)
    tail = int(params.get("tail") or 0)
    header_row = int(params.get("header_row") or 1)
    sheet = params.get("sheet") or ""

    def edit(wb, path, ctx):
        n = 0
        for ws in batch.target_sheets(wb, sheet):
            start = ws.max_column + 1
            if mode == "idcard":
                for i, h in enumerate(["生日", "性别", "年龄", "校验", "地区码"]):
                    ws.cell(row=header_row, column=start + i, value=h)
                for r in range(header_row + 1, ws.max_row + 1):
                    v = ws.cell(row=r, column=col).value
                    info = id_info(v)
                    if not info:
                        ws.cell(row=r, column=start, value="格式错误")
                        continue
                    ws.cell(row=r, column=start, value=info["birth"])
                    ws.cell(row=r, column=start + 1, value=info["gender"])
                    ws.cell(row=r, column=start + 2, value=info["age"])
                    ws.cell(row=r, column=start + 3, value="正确" if info["ok"] else "校验位错")
                    ws.cell(row=r, column=start + 4, value=info["area"])
                    n += 1
            elif mode == "mask":
                ws.cell(row=header_row, column=start, value="脱敏结果")
                for r in range(header_row + 1, ws.max_row + 1):
                    v = ws.cell(row=r, column=col).value
                    if v in (None, ""):
                        continue
                    cell = ws.cell(row=r, column=start, value=mask_value(v, kind, head, tail))
                    cell.number_format = "@"
                    n += 1
            else:
                ws.cell(row=header_row, column=start, value="校验结果")
                for r in range(header_row + 1, ws.max_row + 1):
                    v = ws.cell(row=r, column=col).value
                    if v in (None, ""):
                        continue
                    ws.cell(row=r, column=start, value="正确" if credit_check(v) else "不通过")
                    n += 1
        return "处理 %d 行" % n

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_号码处理", exts=EXCEL_EXT)
