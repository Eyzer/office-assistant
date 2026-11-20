"""JSON 与表格互转。"""
import json
import os

import openpyxl

from utils import batch, excel_utils as xu
from utils.encoding_utils import read_text, write_text
from utils.path_utils import out_path, safe_name

META = {
    "id": "json_ops",
    "cat": "导入与导出",
    "name": "JSON 与表格互转",
    "desc": "把 JSON 文件（对象数组/嵌套列表）转成表格，或把工作簿每个工作表导出成 JSON。",
    "fields": [
        {"key": "mode", "label": "方向", "type": "select", "default": "json2excel",
         "options": [{"value": "json2excel", "label": "JSON → 表格"},
                     {"value": "excel2json", "label": "表格 → JSON"}]},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "导出 JSON 时限定工作表，留空=全部"},
    ],
}


def _pick_rows(data):
    """把各种 JSON 结构归一成 (headers, rows)。"""
    if isinstance(data, dict):
        for v in data.values():
            if isinstance(v, list) and v:
                data = v
                break
        else:
            data = [data]
    if not isinstance(data, list):
        return [], []
    if data and isinstance(data[0], dict):
        headers = []
        for item in data:
            for k in item.keys():
                if k not in headers:
                    headers.append(k)
        rows = [[json.dumps(item.get(h), ensure_ascii=False)
                 if isinstance(item.get(h), (dict, list)) else item.get(h)
                 for h in headers] for item in data]
        return headers, rows
    if data and isinstance(data[0], list):
        width = max(len(r) for r in data)
        return ["列%d" % (i + 1) for i in range(width)], [list(r) + [None] * (width - len(r)) for r in data]
    return ["值"], [[x] for x in data]


def json2excel(f, output_dir, ctx):
    data = json.loads(read_text(f))
    headers, rows = _pick_rows(data)
    if not headers:
        raise ValueError("JSON 结构无法识别（需要对象数组）")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = safe_name(os.path.splitext(os.path.basename(f))[0])[:31] or "Sheet"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    dst = out_path(f, output_dir, suffix="_表格", ext=".xlsx")
    wb.save(dst)
    ctx.info("%s: %d 行 %d 列" % (os.path.basename(f), len(rows), len(headers)))
    return [dst]


def excel2json(f, output_dir, sheet, ctx):
    wb = xu.load(f, data_only=True)
    outs = []
    for ws in batch.target_sheets(wb, sheet):
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [str(h) if h is not None else "列%d" % (i + 1) for i, h in enumerate(rows[0])]
        items = []
        for row in rows[1:]:
            if all(v is None or v == "" for v in row):
                continue
            items.append({headers[i]: row[i] for i in range(min(len(headers), len(row)))})
        stem = safe_name(os.path.splitext(os.path.basename(f))[0])
        sub = None
        if len(wb.sheetnames) > 1:
            sub = os.path.join(output_dir, stem + "_json")
            os.makedirs(sub, exist_ok=True)
        dst = os.path.join(sub, safe_name(ws.title)) + ".json" if sub else \
            os.path.join(output_dir, stem + ".json")
        write_text(dst, json.dumps(items, ensure_ascii=False, indent=2, default=str))
        outs.append(dst)
        ctx.info("%s[%s]: %d 条" % (os.path.basename(f), ws.title, len(items)))
    return outs


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "json2excel"
    sheet = params.get("sheet") or ""
    os.makedirs(output_dir, exist_ok=True)

    if mode == "json2excel":
        def handler(f, ctx):
            return json2excel(f, output_dir, ctx)
        return batch.process(inputs, ctx, handler, {".json"})

    def handler(f, ctx):
        return excel2json(f, output_dir, sheet, ctx)
    return batch.process(inputs, ctx, handler, {".xlsx", ".xlsm"})
