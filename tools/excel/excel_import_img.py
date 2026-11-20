"""按名称把图片批量导入表格。"""
import os

from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "excel_import_img",
    "name": "按名称导入图片",
    "desc": "按某列的文件名匹配图片目录，把图片插入指定列。",
    "fields": [
        {"key": "img_dir", "label": "图片目录", "type": "text", "default": "",
         "help": "图片文件名需与表格里某列内容对应（可带扩展名）"},
        {"key": "key_col", "label": "名称所在列", "type": "text", "default": "A"},
        {"key": "target_col", "label": "插入到列", "type": "text", "default": "E"},
        {"key": "size_px", "label": "图片边长(像素)", "type": "number", "default": 60},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def _index_images(img_dir):
    index = {}
    exts = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}
    for root, _, files in os.walk(img_dir):
        for n in files:
            if os.path.splitext(n)[1].lower() in exts:
                index.setdefault(os.path.splitext(n)[0], os.path.join(root, n))
    return index


def run_task(inputs, output_dir, params, ctx):
    img_dir = params.get("img_dir") or ""
    key_col = column_index_from_string((params.get("key_col") or "A").upper())
    tgt_col = column_index_from_string((params.get("target_col") or "E").upper())
    size = int(params.get("size_px") or 60)
    header_row = int(params.get("header_row") or 1)

    if not os.path.isdir(img_dir):
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["图片目录不存在"], "output_files": []}

    index = _index_images(img_dir)
    ctx.info(f"图片索引 {len(index)} 张")
    files = collect_files(inputs, EXCEL_EXT)
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            for ws in wb.worksheets:
                inserted = 0
                for r in range(header_row + 1, ws.max_row + 1):
                    name = ws.cell(row=r, column=key_col).value
                    if name in (None, ""):
                        continue
                    key = os.path.splitext(str(name).strip())[0]
                    img_path = index.get(key)
                    if not img_path:
                        continue
                    from openpyxl.utils import get_column_letter
                    img = XLImage(img_path)
                    img.width, img.height = size, size
                    ws.add_image(img, f"{get_column_letter(tgt_col)}{r}")
                    ws.row_dimensions[r].height = size * 0.78
                    inserted += 1
                ctx.info(f"[{ws.title}] 插入 {inserted} 张")
            dst = out_path(f, output_dir, suffix="_带图")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
