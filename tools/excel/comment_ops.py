"""批注批量处理：提取、删除、改内容、改作者、加前缀。"""
import os

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT
from utils.path_utils import out_path

META = {
    "id": "comment_ops",
    "cat": "批注",
    "name": "批注批量处理",
    "desc": "把全部批注导出成清单、批量删除、批量改内容/作者，或给每条批注加统一前缀。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "extract",
         "options": [
             {"value": "extract", "label": "导出批注清单"},
             {"value": "delete", "label": "删除批注"},
             {"value": "replace", "label": "查找替换批注内容"},
             {"value": "prefix", "label": "给批注加前缀"},
             {"value": "author", "label": "统一修改作者名"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "find", "label": "查找内容", "type": "text", "default": ""},
        {"key": "repl", "label": "替换为", "type": "text", "default": ""},
        {"key": "author", "label": "作者名", "type": "text", "default": ""},
        {"key": "prefix", "label": "前缀", "type": "text", "default": "【已阅】"},
    ],
}


def _comment_text(cell):
    c = cell.comment
    return "" if c is None or c.text is None else str(c.text)


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "extract"
    sheet = params.get("sheet") or ""
    find = params.get("find") or ""
    repl = params.get("repl") or ""
    author = params.get("author") or ""
    prefix = params.get("prefix") or ""

    def handler(f, ctx):
        wb = xu.load(f)
        total = 0
        if mode == "extract":
            name = xu.safe_sheet_title("批注清单", {s.title.lower() for s in wb.worksheets})
            out = wb.create_sheet(name)
            out.append(["工作表", "单元格", "作者", "内容"])
            for ws in batch.target_sheets(wb, sheet):
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.comment is None:
                            continue
                        out.append([ws.title, cell.coordinate,
                                    cell.comment.author or "", _comment_text(cell)])
                        total += 1
        else:
            for ws in batch.target_sheets(wb, sheet):
                for row in ws.iter_rows():
                    for cell in row:
                        cm = cell.comment
                        if cm is None:
                            continue
                        if mode == "delete":
                            cell.comment = None
                            total += 1
                            continue
                        text = _comment_text(cell)
                        new = text
                        if mode == "replace" and find:
                            new = text.replace(find, repl)
                        elif mode == "prefix":
                            new = prefix + text
                        if new != text:
                            cm.text = new
                            total += 1
                        if mode == "author" and author:
                            cm.author = author
                            total += 1
        dst = out_path(f, output_dir, suffix="_批注")
        xu.save(wb, dst)
        ctx.info("%s: %s %d 处" % (os.path.basename(f), mode, total))
        return [dst]

    return batch.process(inputs, ctx, handler, EXCEL_EXT)
