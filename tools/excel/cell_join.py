"""多列合并为一列（自定义分隔符）。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "cell_join",
    "name": "多列合并成一列",
    "desc": "把多列内容用分隔符拼成一列，可保留原列或只留结果。",
    "fields": [
        {"key": "cols", "label": "参与列", "type": "text", "default": "A,B,C",
         "help": "列字母，逗号分隔"},
        {"key": "sep", "label": "分隔符", "type": "text", "default": " "},
        {"key": "out_col", "label": "结果列", "type": "text", "default": "D"},
        {"key": "skip_blank", "label": "跳过空单元格", "type": "checkbox", "default": True},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    cols = [c.strip().upper() for c in (params.get("cols") or "A,B,C").split(",") if c.strip()]
    sep = params.get("sep", " ")
    out_col = (params.get("out_col") or "D").strip().upper()
    skip_blank = bool(params.get("skip_blank", True))

    from openpyxl.utils import column_index_from_string
    src_idx = [column_index_from_string(c) for c in cols]
    dst_idx = column_index_from_string(out_col)

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            for ws in wb.worksheets:
                for r in range(1, ws.max_row + 1):
                    parts = []
                    for ci in src_idx:
                        v = ws.cell(row=r, column=ci).value
                        if v in (None, ""):
                            if skip_blank:
                                continue
                            v = ""
                        parts.append(str(v))
                    ws.cell(row=r, column=dst_idx, value=sep.join(parts))
            dst = out_path(f, output_dir, suffix="_列合并")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 合并列 {','.join(cols)} -> {out_col}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
