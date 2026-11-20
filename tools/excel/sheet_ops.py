"""工作表批量管理：排序、隐藏、标签色、删除、保护、冻结、按模板新建、生成目录。"""
import re

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "sheet_ops",
    "cat": "工作表与工作簿",
    "name": "工作表批量管理",
    "desc": "工作表排序、隐藏/显示、标签颜色、按名字删除、批量保护与取消保护、冻结首行、按模板批量新建、生成目录页。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "reorder",
         "options": [
             {"value": "reorder", "label": "工作表排序"},
             {"value": "hide", "label": "隐藏工作表"},
             {"value": "show", "label": "显示全部工作表"},
             {"value": "tab_color", "label": "设置标签颜色"},
             {"value": "delete", "label": "删除工作表"},
             {"value": "protect", "label": "批量保护工作表"},
             {"value": "unprotect", "label": "取消工作表保护"},
             {"value": "freeze", "label": "冻结首行/首列"},
             {"value": "new_blank", "label": "批量新建空白工作表"},
             {"value": "from_template", "label": "按模板批量新建"},
             {"value": "catalog", "label": "生成工作表目录"},
         ]},
        {"key": "sheets", "label": "目标工作表", "type": "text", "default": "",
         "help": "逗号分隔；留空=全部"},
        {"key": "color", "label": "标签颜色", "type": "text", "default": "2F6FE4",
         "help": "十六进制，如 2F6FE4"},
        {"key": "password", "label": "保护密码", "type": "text", "default": ""},
        {"key": "freeze", "label": "冻结方式", "type": "select", "default": "row",
         "options": [{"value": "row", "label": "冻结首行"},
                     {"value": "col", "label": "冻结首列"},
                     {"value": "both", "label": "首行+首列"}]},
        {"key": "names", "label": "新表名", "type": "text", "default": "",
         "help": "逗号分隔；留空则按模板首列内容命名"},
        {"key": "count", "label": "新建数量", "type": "number", "default": 1},
        {"key": "template", "label": "模板工作表", "type": "text", "default": "",
         "help": "按模板新建时指定，留空=第一个工作表"},
    ],
}


def _split(text):
    return [x.strip() for x in re.split(r"[,，]", str(text or "")) if x.strip()]


def _pick(wb, names):
    """按名字精确取表；留空=全部。名字写错不会误伤全部表。"""
    if not names:
        return list(wb.worksheets)
    return [ws for ws in wb.worksheets if ws.title in names]


def reorder(wb, desc):
    wb._sheets = sorted(wb._sheets, key=lambda s: s.title, reverse=bool(desc))
    return len(wb._sheets)


def hide(wb, names):
    got = _pick(wb, names)
    visible_left = len([s for s in wb.worksheets if s.sheet_state == "visible"])
    if len(got) >= visible_left and visible_left:
        got = got[:-1]        # 至少留一张可见，Excel 不允许全隐藏
    for ws in got:
        ws.sheet_state = "hidden"
    return len(got)


def show(wb):
    for ws in wb.worksheets:
        ws.sheet_state = "visible"
    return len(wb.worksheets)


def tab_color(wb, color):
    c = str(color or "2F6FE4").lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", c):
        c = "2F6FE4"
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = c
    return len(wb.worksheets)


def delete(wb, names):
    if not names:
        return 0
    n = 0
    for name in names:
        if name in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb[name]
            n += 1
    return n


def protect(wb, names, password):
    got = _pick(wb, names)
    for ws in got:
        ws.protection.sheet = True
        if password:
            ws.protection.set_password(password)
    return len(got)


def unprotect(wb):
    for ws in wb.worksheets:
        ws.protection.sheet = False
    return len(wb.worksheets)


def freeze(wb, how):
    panes = {"row": "A2", "col": "B1", "both": "B2"}
    if how not in panes:
        return 0
    for ws in wb.worksheets:
        ws.freeze_panes = panes[how]
    return len(wb.worksheets)


def new_blank(wb, names, count):
    used = {s.title.lower() for s in wb.worksheets}
    n = 0
    if names:
        for name in names:
            wb.create_sheet(xu.safe_sheet_title(name, used))
            n += 1
    else:
        for i in range(max(1, count)):
            wb.create_sheet(xu.safe_sheet_title("Sheet%d" % (i + 1), used))
            n += 1
    return n


def from_template(wb, template, names, count):
    src = wb[template] if template and template in wb.sheetnames else wb.worksheets[0]
    used = {s.title.lower() for s in wb.worksheets}
    n = 0
    if names:
        for name in names:
            ws = wb.copy_worksheet(src)
            ws.title = xu.safe_sheet_title(name, used)
            n += 1
    else:
        for i in range(max(1, count)):
            ws = wb.copy_worksheet(src)
            ws.title = xu.safe_sheet_title("%s_%d" % (src.title, i + 1), used)
            n += 1
    return n


def catalog(wb):
    used = {s.title.lower() for s in wb.worksheets}
    out = wb.create_sheet(xu.safe_sheet_title("目录", used))
    out.cell(row=1, column=1, value="工作表目录")
    out.cell(row=2, column=1, value="序号")
    out.cell(row=2, column=2, value="工作表名")
    out.cell(row=2, column=3, value="行数")
    n = 0
    for i, ws in enumerate(wb.worksheets, 1):
        if ws.title == out.title:
            continue
        r = i + 2
        out.cell(row=r, column=1, value=i)
        c = out.cell(row=r, column=2, value=ws.title)
        c.hyperlink = "#'%s'!A1" % ws.title
        out.cell(row=r, column=3, value=ws.max_row)
        n += 1
    return n


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "reorder"
    names = _split(params.get("sheets") or params.get("names"))
    color = params.get("color") or "2F6FE4"
    password = params.get("password") or ""
    how = params.get("freeze") or "row"
    count = int(params.get("count") or 1)
    template = params.get("template") or ""

    def edit(wb, path, ctx):
        if mode == "reorder":
            return "重排 %d 张工作表" % reorder(wb, False)
        if mode == "hide":
            return "隐藏 %d 张" % hide(wb, names)
        if mode == "show":
            return "显示 %d 张" % show(wb)
        if mode == "tab_color":
            return "设置 %d 张标签色" % tab_color(wb, color)
        if mode == "delete":
            n = delete(wb, names)
            if not names:
                ctx.warn("删除工作表要填「目标工作表」名字")
            return "删除 %d 张" % n
        if mode == "protect":
            return "保护 %d 张" % protect(wb, names, password)
        if mode == "unprotect":
            return "取消保护 %d 张" % unprotect(wb)
        if mode == "freeze":
            return "冻结 %d 张" % freeze(wb, how)
        if mode == "new_blank":
            return "新建 %d 张" % new_blank(wb, names, count)
        if mode == "from_template":
            return "按模板新建 %d 张" % from_template(wb, template, names, count)
        if mode == "catalog":
            n = catalog(wb)
            return "已生成目录页（%d 张）" % n
        return ""

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_工作表管理", exts=EXCEL_EXT)
