"""单元格内容按分隔符拆成多列。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "cell_split",
    "name": "单元格拆分多列",
    "desc": "把一列内容按分隔符拆成多列，右侧列自动后移。",
    "fields": [
        {"key": "col", "label": "拆分列", "type": "text", "default": "A"},
        {"key": "sep", "label": "分隔符", "type": "text", "default": ","},
        {"key": "max_cols", "label": "最多拆成几列", "type": "number", "default": 3},
        {"key": "keep_origin", "label": "保留原列", "type": "checkbox", "default": True},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    from openpyxl.utils import column_index_from_string
    files = collect_files(inputs, EXCEL_EXT)
    col = (params.get("col") or "A").strip().upper()
    ci = column_index_from_string(col)
    sep = params.get("sep", ",")
    max_cols = max(1, int(params.get("max_cols") or 3))
    keep = bool(params.get("keep_origin", True))

    ok, fail, outs = 0, 0, []
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb = xu.load(f)
            for ws in wb.worksheets:
                values = []
                for r in range(1, ws.max_row + 1):
                    v = ws.cell(row=r, column=ci).value
                    parts = str(v).split(sep) if v not in (None, "") else [""]
                    parts = (parts + [""] * max_cols)[:max_cols]
                    values.append(parts)

                insert_at = ci + 1
                if max_cols > 1:
                    ws.insert_cols(insert_at, max_cols - 1)
                for r, parts in enumerate(values, 1):
                    base = ci + 1 if keep else ci
                    for k, p in enumerate(parts):
                        ws.cell(row=r, column=base + k, value=p)
            dst = out_path(f, output_dir, suffix="_列拆分")
            xu.save(wb, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 拆分 {col} 列")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
