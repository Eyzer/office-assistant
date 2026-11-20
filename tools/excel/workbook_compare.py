"""对比两个工作簿，输出差异清单。"""
import os

from utils import excel_utils as xu
from utils.encoding_utils import write_text
from utils.file_utils import EXCEL_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "workbook_compare",
    "name": "工作簿差异对比",
    "desc": "对比两个表格的单元格差异，输出差异清单。差异超过上限自动停止。",
    "fields": [
        {"key": "limit", "label": "最多记录条数", "type": "number", "default": 5000},
        {"key": "compare_sheets", "label": "逐表对比", "type": "checkbox", "default": True},
    ],
}


def _diff_sheet(ws1, ws2, name, limit):
    rows = []
    max_r = max(ws1.max_row, ws2.max_row)
    max_c = max(ws1.max_column, ws2.max_column)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v1 = ws1.cell(row=r, column=c).value
            v2 = ws2.cell(row=r, column=c).value
            if (v1 or "") != (v2 or ""):
                rows.append([name, f"R{r}C{c}", v1, v2])
                if len(rows) >= limit:
                    return rows
    return rows


def run_task(inputs, output_dir, params, ctx):
    files = sorted(collect_files(inputs, EXCEL_EXT))
    if len(files) < 2:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["需要正好两个表格文件"], "output_files": []}

    limit = int(params.get("limit") or 5000)
    per_sheet = bool(params.get("compare_sheets", True))
    f1, f2 = files[0], files[1]

    wb1 = xu.load(f1, data_only=True)
    wb2 = xu.load(f2, data_only=True)

    lines = [f"对比文件: {os.path.basename(f1)}  <->  {os.path.basename(f2)}", ""]
    total_diff = 0

    pairs = []
    if per_sheet:
        for i in range(max(len(wb1.worksheets), len(wb2.worksheets))):
            s1 = wb1.worksheets[i] if i < len(wb1.worksheets) else None
            s2 = wb2.worksheets[i] if i < len(wb2.worksheets) else None
            pairs.append((s1, s2))
    else:
        pairs.append((wb1.worksheets[0], wb2.worksheets[0]))

    for idx, (s1, s2) in enumerate(pairs, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        if s1 is None or s2 is None:
            lines.append(f"[表{idx}] 一边缺少对应工作表")
            total_diff += 1
            continue
        rows = _diff_sheet(s1, s2, s1.title, limit)
        total_diff += len(rows)
        lines.append(f"[{s1.title}] 差异 {len(rows)} 处")
        for row in rows[:200]:
            lines.append(f"  {row[1]}  左={row[2]!r}  右={row[3]!r}")
        ctx.info(f"表 {s1.title}: {len(rows)} 处差异")
        ctx.progress(idx, len(pairs), s1.title)

    lines.append("")
    lines.append(f"合计差异: {total_diff}")
    dst = out_path(f1, output_dir, suffix="_差异报告", ext=".txt")
    write_text(dst, "\n".join(lines))

    return {"success_count": 1, "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
