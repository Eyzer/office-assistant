"""数据套打：Excel 每行数据填充 Word 模板，批量生成文档。"""
import os

from utils import doc_utils as du
from utils import excel_utils as xu
from utils.path_utils import ensure_dir, safe_name, unique_path

META = {
    "id": "doc_mail_merge",
    "name": "数据套打生成文档",
    "desc": "用 Excel 每行数据填充 Word 模板中的 {{字段}} 占位符，一行生成一份文档。",
    "fields": [
        {"key": "template", "label": "Word 模板路径", "type": "text", "default": "",
         "help": "模板里用 {{列名}} 作占位符"},
        {"key": "data", "label": "数据表路径", "type": "text", "default": ""},
        {"key": "name_field", "label": "命名用列名", "type": "text", "default": "",
         "help": "留空用序号命名"},
        {"key": "header_row", "label": "数据表头行", "type": "number", "default": 1},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    template = (params.get("template") or "").strip()
    data_path = (params.get("data") or "").strip()
    if not template:
        template = next((p for p in inputs if p.lower().endswith(".docx")), "")
    if not data_path:
        data_path = next((p for p in inputs if p.lower().endswith((".xlsx", ".xlsm"))), "")

    if not template or not os.path.isfile(template):
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["未找到 Word 模板"], "output_files": []}
    if not data_path or not os.path.isfile(data_path):
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["未找到数据表"], "output_files": []}

    header_row = int(params.get("header_row") or 1)
    name_field = (params.get("name_field") or "").strip()

    wb = xu.load(data_path, data_only=True)
    ws = wb.worksheets[0]
    headers = [str(xu.cell_text(c.value)) for c in ws[header_row]]
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v in (None, "") for v in vals):
            continue
        rows.append(dict(zip(headers, vals)))

    ensure_dir(output_dir)
    ok, fail, outs = 0, 0, []
    total = len(rows)

    for i, rec in enumerate(rows, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        base = f"第{i}份"
        try:
            d = du.load(template)
            for key, val in rec.items():
                if key:
                    du.replace_text(d, "{{" + key + "}}", "" if val is None else str(val),
                                    case_sensitive=True)
            if name_field and rec.get(name_field) not in (None, ""):
                base = safe_name(str(rec[name_field]))
            else:
                base = f"第{i}份"
            dst = unique_path(os.path.join(output_dir, base + ".docx"))
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"生成 {os.path.basename(dst)}")
        except Exception as e:
            fail += 1
            ctx.error(f"第{i}行: {e}")
        ctx.progress(i, total, base)

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
