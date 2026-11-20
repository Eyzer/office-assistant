"""批处理公共骨架：统一「遍历文件 → 处理 → 保存 → 进度 → 失败隔离」。

新增工具只需写业务函数，循环、日志、进度、异常都在这里收口，
避免每个工具各写一份、行为不一致。
"""
import os

from utils import excel_utils as xu
from utils.deps import MissingDep
from utils.file_utils import collect_files
from utils.path_utils import out_path

# 常见列参数
def col_index(value, default=1):
    """列参数转序号：支持 A / 3 / 'A,C' 里的单个值。"""
    s = str(value or "").strip()
    if not s:
        return default
    if s.isdigit():
        return int(s)
    from openpyxl.utils import column_index_from_string
    try:
        return column_index_from_string(s.upper())
    except (ValueError, KeyError):
        return default


def target_sheets(wb, names):
    """按名字挑工作表；留空或 * 表示全部。"""
    s = str(names or "").strip()
    if not s or s == "*":
        return list(wb.worksheets)
    want = [x.strip() for x in s.replace("，", ",").split(",") if x.strip()]
    got = [ws for ws in wb.worksheets if ws.title in want]
    return got or list(wb.worksheets)


def process(inputs, ctx, handler, exts=None):
    """通用循环。handler(path, ctx) 返回输出路径列表 / 消息 / None。"""
    files = collect_files(inputs, exts)
    total = len(files)
    if not total:
        ctx.warn("没有找到可处理的文件")
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ctx.task["log"].lines[-50:], "output_files": []}

    ok = fail = 0
    outs = []
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消，剩余 %d 个文件未处理" % (total - i + 1))
            break
        try:
            r = handler(f, ctx)
            if isinstance(r, (list, tuple)):
                outs.extend([x for x in r if x])
            ok += 1
        except MissingDep:
            raise                      # 缺依赖是全局问题，直接终止并上报
        except Exception as exc:
            fail += 1
            ctx.error("%s: %s" % (os.path.basename(f), exc))
        ctx.progress(i, total, os.path.basename(f))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}


def excel_edit(inputs, output_dir, ctx, edit, suffix="_处理后", exts=None):
    """载入工作簿 → edit(wb, path, ctx) → 另存新副本。"""
    def handler(f, ctx):
        wb = xu.load(f)
        msg = edit(wb, f, ctx)
        dst = out_path(f, output_dir, suffix=suffix)
        xu.save(wb, dst)
        if msg:
            ctx.info("%s: %s" % (os.path.basename(f), msg))
        return [dst]
    return process(inputs, ctx, handler, exts)


def each_sheet(wb, names, fn):
    """对指定工作表逐个执行 fn(ws)，返回累计计数。"""
    n = 0
    for ws in target_sheets(wb, names):
        n += fn(ws) or 0
    return n


def first_row(ws):
    return 1


def header_of(ws, row=1):
    return [str(c.value).strip() if c.value is not None else "" for c in ws[row]]


def insert_rows(ws, at, amount, sample_row=None):
    """插入行，并把样例行的格式复制到新行（保持外观一致）。"""
    ws.insert_rows(at, amount)
    if sample_row:
        from copy import copy
        for r in range(at, at + amount):
            for c in range(1, ws.max_column + 1):
                src = ws.cell(row=sample_row, column=c)
                dst = ws.cell(row=r, column=c)
                if src.has_style:
                    dst._style = copy(src._style)


def insert_cols(ws, at, amount, sample_col=None):
    ws.insert_cols(at, amount)
    if sample_col:
        from copy import copy
        for c in range(at, at + amount):
            for r in range(1, ws.max_row + 1):
                src = ws.cell(row=r, column=sample_col)
                dst = ws.cell(row=r, column=c)
                if src.has_style:
                    dst._style = copy(src._style)


def norm_header(name):
    """表头归一：去空白、去全半角差异，便于按名字找列。"""
    s = str(name or "").strip()
    return s.replace("　", "").replace(" ", "")


def find_col(ws, name, header_row=1):
    """按表头名字找列号，找不到返回 0。"""
    if not name:
        return 0
    want = norm_header(name)
    for c in range(1, ws.max_column + 1):
        if norm_header(ws.cell(row=header_row, column=c).value) == want:
            return c
    for c in range(1, ws.max_column + 1):
        if want and want in norm_header(ws.cell(row=header_row, column=c).value):
            return c
    return 0
