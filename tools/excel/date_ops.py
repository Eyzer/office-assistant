"""日期时间：成分提取、时间戳互转、中文日期标准化、拆分合并、农历、工作日统计。"""
import datetime
import re

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "date_ops",
    "cat": "数值与计算",
    "name": "日期时间处理",
    "desc": "提取年月日星期、时间戳互转、中文日期标准化、清空非日期、日期时间拆分合并、公历转农历、工作日统计。",
    "fields": [
        {"key": "mode", "label": "处理方式", "type": "select", "default": "parts",
         "options": [
             {"value": "parts", "label": "提取成分（年/月/日/星期…）"},
             {"value": "timestamp", "label": "时间戳 ↔ 日期时间"},
             {"value": "normalize", "label": "中文日期标准化"},
             {"value": "clear", "label": "清空非日期时间"},
             {"value": "split", "label": "日期时间拆成两列"},
             {"value": "merge", "label": "日期 + 时间合并成一列"},
             {"value": "lunar", "label": "公历转农历"},
             {"value": "workday", "label": "工作日统计"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "留空=全部工作表"},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "col", "label": "日期列", "type": "text", "default": "A"},
        {"key": "col2", "label": "第二列", "type": "text", "default": "",
         "help": "合并模式填时间列；工作日模式填结束日期列"},
        {"key": "parts", "label": "提取哪些", "type": "text", "default": "年,月,日,星期",
         "help": "逗号分隔，可选：年,月,日,时,分,秒,星期,季度,周数"},
        {"key": "direction", "label": "时间戳方向", "type": "select", "default": "to_date",
         "options": [{"value": "to_date", "label": "时间戳 → 日期"},
                     {"value": "to_ts", "label": "日期 → 时间戳（秒）"}]},
        {"key": "holidays", "label": "节假日", "type": "text", "default": "",
         "help": "工作日统计用，逗号分隔，如 2024-10-01,2024-10-02"},
        {"key": "extra_work", "label": "调休上班日", "type": "text", "default": "",
         "help": "工作日统计用，这些周末算上班"},
    ],
}

_WEEK = ["一", "二", "三", "四", "五", "六", "日"]
_FORMATS = (
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d",
    "%Y.%m.%d", "%Y%m%d", "%Y%m%d%H%M%S",
    "%d/%m/%Y", "%m/%d/%Y", "%Y年%m月%d日",
)


def parse_dt(v):
    """尽力把各种写法解析成 datetime，失败返回 None。"""
    if v is None or v == "":
        return None
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime(v.year, v.month, v.day)
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        # Excel 序列号区间（1954~2119）
        if 20000 <= float(v) <= 80000:
            from openpyxl.utils.datetime import from_excel
            try:
                return from_excel(v)
            except Exception:
                return None
        if 1e9 <= float(v) <= 4e9:      # 10 位秒级时间戳
            return datetime.datetime.fromtimestamp(float(v))
        if 1e12 <= float(v) <= 4e12:    # 13 位毫秒
            return datetime.datetime.fromtimestamp(float(v) / 1000.0)
        return None
    s = str(v).strip()
    if not s:
        return None
    s2 = s.replace("年", "-").replace("月", "-").replace("日", "")
    s2 = s2.replace("号", "").replace("时", ":").replace("分", ":").replace("秒", "")
    s2 = re.sub(r"\s+", " ", s2).strip(" -:")
    for text in (s, s2):
        for fmt in _FORMATS:
            try:
                return datetime.datetime.strptime(text, fmt)
            except ValueError:
                continue
    m = re.match(r"^(\d{4})[-/.]?(\d{1,2})[-/.]?(\d{1,2})$", s2)
    if m:
        try:
            return datetime.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def _out_header(ws, header_row, names):
    base = ws.max_column
    for i, nm in enumerate(names):
        ws.cell(row=header_row, column=base + i + 1, value=nm)
    return base + 1


def do_parts(ws, header_row, col, parts):
    c = batch.col_index(col, 1)
    keys = [p.strip() for p in str(parts or "").replace("，", ",").split(",") if p.strip()]
    if not keys:
        keys = ["年", "月", "日"]
    start = _out_header(ws, header_row, keys)
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        dt = parse_dt(ws.cell(row=r, column=c).value)
        if not dt:
            continue
        for i, k in enumerate(keys):
            if k in ("年", "year"):
                v = dt.year
            elif k in ("月", "month"):
                v = dt.month
            elif k in ("日", "day"):
                v = dt.day
            elif k in ("时", "hour"):
                v = dt.hour
            elif k in ("分", "minute"):
                v = dt.minute
            elif k in ("秒", "second"):
                v = dt.second
            elif k in ("星期", "week"):
                v = "星期" + _WEEK[dt.weekday()]
            elif k in ("季度", "quarter"):
                v = "第%d季度" % ((dt.month - 1) // 3 + 1)
            elif k in ("周数", "weekno"):
                v = dt.isocalendar()[1]
            else:
                v = None
            ws.cell(row=r, column=start + i, value=v)
        n += 1
    return n


def do_timestamp(ws, header_row, col, direction):
    c = batch.col_index(col, 1)
    base = ws.max_column
    ws.cell(row=header_row, column=base + 1, value="转换结果")
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=c).value
        if v in (None, ""):
            continue
        if direction == "to_ts":
            dt = parse_dt(v)
            if not dt:
                continue
            ws.cell(row=r, column=base + 1, value=int(dt.timestamp()))
        else:
            dt = None
            try:
                f = float(v)
                dt = (datetime.datetime.fromtimestamp(f / 1000.0) if f > 4e10
                      else datetime.datetime.fromtimestamp(f))
            except (TypeError, ValueError):
                dt = parse_dt(v)
            if not dt:
                continue
            tgt = ws.cell(row=r, column=base + 1, value=dt)
            tgt.number_format = "yyyy-mm-dd hh:mm:ss"
        n += 1
    return n


def do_normalize(ws, header_row, col):
    c = batch.col_index(col, 1)
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        cell = ws.cell(row=r, column=c)
        dt = parse_dt(cell.value)
        if not dt:
            continue
        cell.value = dt
        cell.number_format = "yyyy-mm-dd"
        n += 1
    return n


def do_clear(ws, header_row, col):
    c = batch.col_index(col, 1)
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        cell = ws.cell(row=r, column=c)
        if cell.value in (None, ""):
            continue
        if parse_dt(cell.value) is None:
            cell.value = None
            n += 1
    return n


def do_split(ws, header_row, col):
    c = batch.col_index(col, 1)
    dc = _out_header(ws, header_row, ["日期", "时间"])
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        dt = parse_dt(ws.cell(row=r, column=c).value)
        if not dt:
            continue
        a = ws.cell(row=r, column=dc, value=dt.date())
        a.number_format = "yyyy-mm-dd"
        b = ws.cell(row=r, column=dc + 1, value=dt.strftime("%H:%M:%S"))
        b.number_format = "@"
        n += 1
    return n


def do_merge(ws, header_row, col, col2):
    c1 = batch.col_index(col, 1)
    c2 = batch.col_index(col2, 2)
    base = ws.max_column
    ws.cell(row=header_row, column=base + 1, value="合并时间")
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        d = parse_dt(ws.cell(row=r, column=c1).value)
        if not d:
            continue
        t = ws.cell(row=r, column=c2).value
        out = d
        if isinstance(t, datetime.time):
            out = datetime.datetime.combine(d.date(), t)
        elif isinstance(t, datetime.datetime):
            out = datetime.datetime.combine(d.date(), t.time())
        elif t not in (None, ""):
            dt = parse_dt("2000-01-01 " + str(t)) or parse_dt(str(t))
            if dt:
                out = (
                    datetime.datetime.combine(d.date(), dt.time())
                    if dt.year == 2000 else dt
                )
        cell = ws.cell(row=r, column=base + 1, value=out)
        cell.number_format = "yyyy-mm-dd hh:mm:ss"
        n += 1
    return n


_LUNAR_MONTH = ["正", "二", "三", "四", "五", "六", "七", "八", "九", "十", "冬", "腊"]
_LUNAR_DAY = [
    "初一", "初二", "初三", "初四", "初五", "初六", "初七", "初八", "初九", "初十",
    "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
    "廿一", "廿二", "廿三", "廿四", "廿五", "廿六", "廿七", "廿八", "廿九", "三十",
]


def do_lunar(ws, header_row, col, ctx):
    try:
        from lunardate import LunarDate
    except ImportError:
        ctx.warn("未安装 lunardate，公历转农历跳过（pip install lunardate 或放入 vendor/）")
        return 0
    c = batch.col_index(col, 1)
    base = ws.max_column
    ws.cell(row=header_row, column=base + 1, value="农历")
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        dt = parse_dt(ws.cell(row=r, column=c).value)
        if not dt:
            continue
        try:
            ld = LunarDate.fromSolarDate(dt.year, dt.month, dt.day)
            m = _LUNAR_MONTH[ld.month - 1]
            text = "%s%s月%s" % ("闰" if getattr(ld, "isLeapMonth", False) else "",
                                 m, _LUNAR_DAY[ld.day - 1])
            ws.cell(row=r, column=base + 1, value=text)
            n += 1
        except Exception:
            continue
    return n


def _parse_day_set(text):
    out = set()
    for p in str(text or "").replace("，", ",").split(","):
        p = p.strip()
        if not p:
            continue
        dt = parse_dt(p)
        if dt:
            out.add(dt.date())
    return out


def do_workday(ws, header_row, col, col2, holidays, extra):
    c1, c2 = batch.col_index(col, 1), batch.col_index(col2, 2)
    hol = _parse_day_set(holidays)
    ext = _parse_day_set(extra)
    base = ws.max_column
    ws.cell(row=header_row, column=base + 1, value="工作日数")
    n = 0
    for r in range(header_row + 1, ws.max_row + 1):
        d1 = parse_dt(ws.cell(row=r, column=c1).value)
        d2 = parse_dt(ws.cell(row=r, column=c2).value)
        if not (d1 and d2):
            continue
        a, b = d1.date(), d2.date()
        if a > b:
            a, b = b, a
        days = 0
        cur = a
        while cur <= b:
            is_work = cur.weekday() < 5 or cur in ext
            if is_work and cur not in hol:
                days += 1
            cur += datetime.timedelta(days=1)
        ws.cell(row=r, column=base + 1, value=days)
        n += 1
    return n


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "parts"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    col = params.get("col") or "A"
    col2 = params.get("col2") or ""
    parts = params.get("parts")
    direction = params.get("direction") or "to_date"
    holidays = params.get("holidays")
    extra = params.get("extra_work")

    if mode in ("merge", "workday") and not col2:
        ctx.error("%s 模式需要填「第二列」" % ("合并" if mode == "merge" else "工作日统计"))
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    def edit(wb, path, ctx):
        total = 0
        for ws in batch.target_sheets(wb, sheet):
            if mode == "parts":
                total += do_parts(ws, header_row, col, parts)
            elif mode == "timestamp":
                total += do_timestamp(ws, header_row, col, direction)
            elif mode == "normalize":
                total += do_normalize(ws, header_row, col)
            elif mode == "clear":
                total += do_clear(ws, header_row, col)
            elif mode == "split":
                total += do_split(ws, header_row, col)
            elif mode == "merge":
                total += do_merge(ws, header_row, col, col2)
            elif mode == "lunar":
                total += do_lunar(ws, header_row, col, ctx)
            elif mode == "workday":
                total += do_workday(ws, header_row, col, col2, holidays, extra)
        return "处理 %d 行" % total

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_日期处理", exts=EXCEL_EXT)
