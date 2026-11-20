"""批量生成条形码（Code128 / Code39 / EAN13 / EAN8），可插入单元格旁边。"""
import os

from openpyxl.drawing.image import Image as XLImage

from utils import batch, deps, excel_utils as xu
from utils.file_utils import EXCEL_EXT
from utils.path_utils import batch_out_dir, out_path, safe_name, unique_path

META = {
    "id": "barcode_gen",
    "cat": "图片与条码",
    "name": "批量生成条形码",
    "desc": "按某列内容生成 Code128/Code39/EAN13 条形码图片，可自动插入到单元格旁边。",
    "fields": [
        {"key": "type", "label": "条码类型", "type": "select", "default": "code128",
         "options": [{"value": "code128", "label": "Code128（任意字符，最常用）"},
                     {"value": "code39", "label": "Code39（数字+大写字母）"},
                     {"value": "ean13", "label": "EAN13（13 位商品码）"},
                     {"value": "ean8", "label": "EAN8（8 位商品码）"}]},
        {"key": "col", "label": "内容列", "type": "text", "default": "A"},
        {"key": "header_row", "label": "表头行", "type": "number", "default": 1},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "留空=第一张"},
        {"key": "height", "label": "条码高度(mm)", "type": "number", "default": 12},
        {"key": "insert", "label": "插入到单元格右侧", "type": "checkbox", "default": True},
    ],
}


def _make(path_dir, value, kind, height):
    from barcode.writer import ImageWriter
    barcode = deps.load("barcode", "生成条形码")
    cls = barcode.get_barcode_class(kind)
    code = cls(str(value), writer=ImageWriter())
    stem = safe_name(str(value))[:60] or "code"
    dst = unique_path(os.path.join(path_dir, stem))
    opts = {"module_height": float(height or 12), "write_text": True,
            "font_size": 8, "text_distance": 2.0, "quiet_zone": 2.0}
    return code.save(dst, options=opts)


def run_task(inputs, output_dir, params, ctx):
    kind = params.get("type") or "code128"
    col = batch.col_index(params.get("col") or "A", 1)
    header_row = int(params.get("header_row") or 1)
    sheet = params.get("sheet") or ""
    height = float(params.get("height") or 12)
    insert = bool(params.get("insert", True))

    if not deps.need(ctx, ["barcode", "PIL"]):
        return {"success_count": 0, "fail_count": 0,
                "output_files": [], "log_messages": []}

    def handler(f, ctx):
        wb = xu.load(f)
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
        sub = batch_out_dir(f, output_dir, "_条形码")
        made, bad = 0, 0
        for r in range(header_row + 1, ws.max_row + 1):
            v = ws.cell(row=r, column=col).value
            if v in (None, ""):
                continue
            try:
                png = _make(sub, v, kind, height)
            except Exception as exc:
                bad += 1
                ctx.warn("%s 第%d行「%s」生成失败：%s" % (os.path.basename(f), r, v, exc))
                continue
            made += 1
            if insert:
                try:
                    img = XLImage(png)
                    img.width, img.height = 220, 70
                    ws.add_image(img, ws.cell(row=r, column=col + 1).coordinate)
                    ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 15, 55)
                except Exception as exc:
                    ctx.warn("插入图片失败：%s" % exc)
        if insert:
            dst = out_path(f, output_dir, suffix="_带条码")
            xu.save(wb, dst)
            ctx.info("%s: 生成 %d 个条码（失败 %d），已插入单元格" %
                     (os.path.basename(f), made, bad))
            return [dst, sub]
        ctx.info("%s: 生成 %d 个条码（失败 %d）" % (os.path.basename(f), made, bad))
        return [sub]

    return batch.process(inputs, ctx, handler, EXCEL_EXT)
