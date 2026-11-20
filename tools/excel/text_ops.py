"""文本批处理：加前后缀、间隔插入、删除文本/空格/特殊字符、大小写。"""
import re

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT

META = {
    "id": "text_ops",
    "cat": "文本处理",
    "name": "文本增删改",
    "desc": "批量加前缀后缀、每隔几个字插入字符、删除指定文本/空格/特殊字符、统一大小写。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "prefix",
         "options": [
             {"value": "prefix", "label": "加前缀"},
             {"value": "suffix", "label": "加后缀"},
             {"value": "insert", "label": "每隔 N 字插入字符"},
             {"value": "remove", "label": "删除指定文本"},
             {"value": "spaces", "label": "删除空格"},
             {"value": "special", "label": "删除特殊字符"},
             {"value": "case", "label": "大小写转换"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "col", "label": "作用列", "type": "text", "default": "A",
         "help": "列字母，留空=全部列"},
        {"key": "text", "label": "文本", "type": "text", "default": "",
         "help": "前缀/后缀/插入/要删除的内容"},
        {"key": "every", "label": "每隔 N 字", "type": "number", "default": 3},
        {"key": "case", "label": "大小写", "type": "select", "default": "upper",
         "options": [{"value": "upper", "label": "全部大写"},
                     {"value": "lower", "label": "全部小写"},
                     {"value": "title", "label": "每词首字母大写"}]},
    ],
}


def _cols(ws, col):
    if not str(col or "").strip():
        return list(range(1, ws.max_column + 1))
    return [batch.col_index(col, 1)]


def _apply(ws, header_row, col, fn):
    n = 0
    for c in _cols(ws, col):
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value in (None, ""):
                continue
            new = fn(str(cell.value))
            if new != cell.value:
                cell.value = new
                n += 1
    return n


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "prefix"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    col = params.get("col") or ""
    text = params.get("text")
    text = "" if text is None else str(text)
    every = max(1, int(params.get("every") or 3))
    case = params.get("case") or "upper"

    def transform(mode, text, every, case):
        if mode == "prefix":
            return lambda s: text + s
        if mode == "suffix":
            return lambda s: s + text
        if mode == "insert":
            def f(s):
                if not text:
                    return s
                out = []
                for i in range(0, len(s), every):
                    if i:
                        out.append(text)
                    out.append(s[i:i + every])
                return "".join(out)
            return f
        if mode == "remove":
            return lambda s: s.replace(text, "") if text else s
        if mode == "spaces":
            return lambda s: re.sub(r"\s+", "", s)
        if mode == "special":
            return lambda s: re.sub(r"[^\w\u4e00-\u9fa5.\-]", "", s)
        if mode == "case":
            if case == "lower":
                return lambda s: s.lower()
            if case == "title":
                return lambda s: s.title()
            return lambda s: s.upper()
        return lambda s: s

    fn = transform(mode, text, every, case)

    def edit(wb, path, ctx):
        total = 0
        for ws in batch.target_sheets(wb, sheet):
            total += _apply(ws, header_row, col, fn)
        return "修改 %d 个单元格" % total

    return batch.excel_edit(inputs, output_dir, ctx, edit,
                            suffix="_文本处理", exts=EXCEL_EXT)
