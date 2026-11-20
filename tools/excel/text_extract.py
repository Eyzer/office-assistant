"""文本提取：手机号、邮箱、网址、身份证、数值、字符、日期、唯一值。"""
import os
import re

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT
from utils.path_utils import out_path

META = {
    "id": "text_extract",
    "cat": "文本处理",
    "name": "文本信息提取",
    "desc": "从单元格里抽出手机号、邮箱、网址、身份证号、数字、汉字/字母，或把整列去重导出。",
    "fields": [
        {"key": "mode", "label": "提取什么", "type": "select", "default": "phone",
         "options": [
             {"value": "phone", "label": "手机号"},
             {"value": "email", "label": "邮箱"},
             {"value": "url", "label": "网址"},
             {"value": "idcard", "label": "身份证号"},
             {"value": "number", "label": "数字"},
             {"value": "cn", "label": "汉字"},
             {"value": "letter", "label": "字母"},
             {"value": "date", "label": "日期"},
             {"value": "unique", "label": "整列去重（导出到新表）"},
         ]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": ""},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "col", "label": "作用列", "type": "text", "default": "A"},
        {"key": "replace", "label": "覆盖原单元格", "type": "checkbox", "default": False,
         "help": "不勾=追加到最后一列右侧"},
    ],
}

_RULES = {
    "phone": r"1[3-9]\d{9}",
    "email": r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
    "url": r"(?:https?://|www\.)[^\s，。,；;]+",
    "idcard": r"\d{17}[\dXx]",
    "number": r"-?\d+(?:\.\d+)?",
    "cn": r"[\u4e00-\u9fa5]+",
    "letter": r"[A-Za-z]+",
    "date": r"\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?",
}


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "phone"
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    col = params.get("col") or "A"
    replace = bool(params.get("replace"))

    def handler(f, ctx):
        wb = xu.load(f)
        n = 0
        for ws in batch.target_sheets(wb, sheet):
            c = batch.col_index(col, 1)
            if mode == "unique":
                seen, vals = set(), []
                for r in range(header_row + 1, ws.max_row + 1):
                    v = ws.cell(row=r, column=c).value
                    if v in (None, ""):
                        continue
                    k = str(v).strip()
                    if k and k not in seen:
                        seen.add(k)
                        vals.append(k)
                name = xu.safe_sheet_title("唯一值_" + ws.title,
                                           {s.title.lower() for s in wb.worksheets})
                out = wb.create_sheet(name)
                out.cell(row=1, column=1, value="唯一值")
                for i, v in enumerate(vals, 2):
                    out.cell(row=i, column=1, value=v)
                n += len(vals)
                continue

            pat = re.compile(_RULES.get(mode, r".+"))
            out_col = c if replace else ws.max_column + 1
            if not replace:
                ws.cell(row=header_row, column=out_col, value="提取_" + mode)
            for r in range(header_row + 1, ws.max_row + 1):
                cell = ws.cell(row=r, column=c)
                if cell.value in (None, ""):
                    continue
                found = pat.findall(str(cell.value))
                if not found:
                    continue
                if isinstance(found[0], tuple):
                    found = [x for grp in found for x in grp if x]
                cell2 = ws.cell(row=r, column=out_col, value=found[0] if len(found) == 1 else ",".join(map(str, found)))
                cell2.number_format = "@"
                n += 1
        dst = out_path(f, output_dir, suffix="_信息提取")
        xu.save(wb, dst)
        ctx.info("%s: 提取 %d 处" % (os.path.basename(f), n))
        return [dst]

    return batch.process(inputs, ctx, handler, EXCEL_EXT)
