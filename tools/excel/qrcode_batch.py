"""批量生成二维码（需 qrcode 库，未装时给出明确提示）。"""
import os

from utils.file_utils import collect_files
from utils.path_utils import batch_out_dir

META = {
    "id": "qrcode_batch",
    "name": "批量生成二维码",
    "desc": "按表格某列内容生成二维码图片；需安装 qrcode 与 Pillow。",
    "fields": [
        {"key": "col", "label": "内容所在列", "type": "text", "default": "A"},
        {"key": "size", "label": "尺寸(像素)", "type": "number", "default": 300},
        {"key": "header_row", "label": "表头行号", "type": "number", "default": 1},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    try:
        import qrcode
    except ImportError:
        msg = "缺少 qrcode 库，请先在目标机执行：pip install qrcode"
        ctx.error(msg)
        return {"success_count": 0, "fail_count": 1,
                "log_messages": [msg], "output_files": []}

    from openpyxl.utils import column_index_from_string
    from utils import excel_utils as xu

    col = column_index_from_string((params.get("col") or "A").upper())
    size = int(params.get("size") or 300)
    header_row = int(params.get("header_row") or 1)

    files = collect_files(inputs, {".xlsx", ".xlsm", ".csv"})
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            if f.lower().endswith(".csv"):
                header, rows = xu.csv_rows(f)
                items = [r[col - 1] for r in rows if len(r) >= col]
            else:
                wb = xu.load(f, data_only=True)
                ws = wb.worksheets[0]
                items = [ws.cell(row=r, column=col).value for r in range(header_row + 1, ws.max_row + 1)]

            sub = batch_out_dir(f, output_dir, "_二维码")
            n = 0
            for v in items:
                if v in (None, ""):
                    continue
                qr = qrcode.QRCode(box_size=10, border=2)
                qr.add_data(str(v))
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")
                img = img.resize((size, size))
                from utils.path_utils import safe_name
                dst = os.path.join(sub, safe_name(str(v)) + ".png")
                if os.path.exists(dst):
                    stem, ext = os.path.splitext(dst)
                    k = 1
                    while os.path.exists(dst):
                        dst = f"{stem}_{k}{ext}"
                        k += 1
                img.save(dst)
                outs.append(dst)
                n += 1
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 生成 {n} 个二维码")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
