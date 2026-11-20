"""公式转数值：把公式格替换成计算结果。"""
import os

from utils import excel_utils as xu
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "formula_to_value",
    "name": "公式转数值",
    "desc": "把单元格公式替换为已计算的结果值，避免拆分合并后引用错位。",
    "fields": [
        {"key": "keep_calc", "label": "保留计算值", "type": "checkbox", "default": True,
         "help": "勾选=用缓存结果；否则公式格清空"},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, EXCEL_EXT)
    keep = bool(params.get("keep_calc", True))
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            wb_formula = xu.load(f, data_only=False)
            wb_value = xu.load(f, data_only=True) if keep else None
            count = 0
            for ws in wb_formula.worksheets:
                vs = wb_value[ws.title] if wb_value else None
                for row in ws.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            if vs is not None:
                                cell.value = vs.cell(row=cell.row, column=cell.column).value
                            else:
                                cell.value = None
                            count += 1
            dst = out_path(f, output_dir, suffix="_数值化")
            xu.save(wb_formula, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 转换 {count} 个公式")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
