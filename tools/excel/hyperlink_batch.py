"""超链接批量处理：把文件清单转成可点链接，或清除超链接。"""
import os

from openpyxl.styles import Font

from utils import excel_utils as xu
from utils.file_utils import collect_files
from utils.path_utils import out_path

META = {
    "id": "hyperlink_batch",
    "name": "超链接批处理",
    "desc": "把文件（夹）清单写成带超链接的表格，或清除表内超链接。",
    "fields": [
        {"key": "mode", "label": "操作", "type": "select", "default": "build",
         "options": [{"value": "build", "label": "文件清单转超链接表"},
                     {"value": "clear", "label": "清除已有超链接"}]},
        {"key": "path_depth", "label": "路径层级", "type": "select", "default": "full",
         "options": [{"value": "full", "label": "完整路径"}, {"value": "name", "label": "仅文件名"}]},
    ],
}


def _build(inputs, dst, depth):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "文件清单"
    ws.append(["序号", "文件名", "大小(KB)", "链接"])
    paths = collect_files(inputs)
    for i, p in enumerate(paths, 1):
        size = round(os.path.getsize(p) / 1024, 1)
        url = p if depth == "full" else p
        r = i + 1
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=os.path.basename(p))
        ws.cell(row=r, column=3, value=size)
        cell = ws.cell(row=r, column=4, value=os.path.basename(p) if depth == "name" else p)
        cell.hyperlink = "file://" + url.replace("\\", "/")
        cell.font = Font(color="0563C1", underline="single")
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["D"].width = 60
    xu.save(wb, dst)
    return len(paths)


def _clear(path, dst):
    wb = xu.load(path)
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.hyperlink is not None:
                    cell.hyperlink = None
                    n += 1
    xu.save(wb, dst)
    return n


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "build"
    depth = params.get("path_depth") or "full"
    outs = []

    if mode == "build":
        anchor = inputs[0] if inputs else "清单.xlsx"
        dst = out_path(anchor if os.path.isfile(anchor) else os.path.join(anchor, "清单.xlsx"),
                       output_dir, suffix="_清单", ext=".xlsx")
        n = _build(inputs, dst, depth)
        outs.append(dst)
        ctx.info(f"生成 {n} 条超链接")
        ctx.progress(1, 1, "完成")
        return {"success_count": 1, "fail_count": 0,
                "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}

    files = collect_files(inputs, {".xlsx", ".xlsm"})
    ok, fail = 0, 0
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            break
        try:
            dst = out_path(f, output_dir, suffix="_去链接")
            n = _clear(f, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 清除 {n} 个超链接")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, len(files), os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
