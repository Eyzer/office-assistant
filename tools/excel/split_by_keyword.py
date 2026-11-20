"""按关键字拆表：把一列里含指定关键字的行拆成独立工作表。"""
import os

import openpyxl

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT
from utils.path_utils import batch_out_dir, out_path

META = {
    "id": "split_by_keyword",
    "cat": "拆分与合并",
    "name": "按关键字拆表",
    "desc": "按某列是否包含关键字拆成多个工作表（如按部门/地区关键字），可另存为独立文件。",
    "fields": [
        {"key": "col", "label": "查找列", "type": "text", "default": "A"},
        {"key": "keywords", "label": "关键字", "type": "text", "default": "",
         "help": "逗号分隔，如 销售,财务,人事"},
        {"key": "sheet", "label": "源工作表", "type": "text", "default": "",
         "help": "留空=第一张工作表"},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "keep_header", "label": "保留表头", "type": "checkbox", "default": True},
        {"key": "rest", "label": "余下另存一张", "type": "checkbox", "default": True,
         "help": "不勾=不含关键字的行丢弃"},
        {"key": "to_file", "label": "拆成独立文件", "type": "checkbox", "default": False},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    col = batch.col_index(params.get("col") or "A", 1)
    raw = str(params.get("keywords") or "").replace("，", ",")
    keys = [k.strip() for k in raw.split(",") if k.strip()]
    sheet = params.get("sheet") or ""
    header_row = int(params.get("header_row") or 1)
    keep_header = bool(params.get("keep_header", True))
    want_rest = bool(params.get("rest", True))
    to_file = bool(params.get("to_file"))

    if not keys:
        ctx.error("请填写要拆的关键字")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    def handler(f, ctx):
        wb = xu.load(f)
        src = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
        header = [c.value for c in src[header_row]] if keep_header else None
        buckets = {k: [] for k in keys}
        rest_rows = []
        for r in range(header_row + 1, src.max_row + 1):
            v = src.cell(row=r, column=col).value
            text = "" if v is None else str(v)
            hit = [k for k in keys if k in text]
            if hit:
                for k in hit:
                    buckets[k].append([c.value for c in src[r]])
            else:
                rest_rows.append([c.value for c in src[r]])

        if to_file:
            base = batch_out_dir(f, output_dir, "_按关键字")
            outs = []
            for name, rows in list(buckets.items()) + ([("其他", rest_rows)] if want_rest else []):
                if not rows:
                    continue
                nwb = openpyxl.Workbook()
                nws = nwb.active
                nws.title = xu.safe_sheet_title(name)
                if header:
                    nws.append(header)
                for row in rows:
                    nws.append(row)
                dst = os.path.join(base, xu.safe_sheet_title(name) + ".xlsx")
                nwb.save(dst)
                outs.append(dst)
            ctx.info("%s: 拆成 %d 个文件" % (os.path.basename(f), len(outs)))
            return outs

        used = {s.title.lower() for s in wb.worksheets}
        made = 0
        for name, rows in list(buckets.items()) + ([("其他", rest_rows)] if want_rest else []):
            if not rows:
                continue
            ws = wb.create_sheet(xu.safe_sheet_title(name, used))
            if header:
                ws.append(header)
            for row in rows:
                ws.append(row)
            made += 1
        dst = out_path(f, output_dir, suffix="_按关键字拆")
        xu.save(wb, dst)
        ctx.info("%s: 生成 %d 张表" % (os.path.basename(f), made))
        return [dst]

    return batch.process(inputs, ctx, handler, EXCEL_EXT)
