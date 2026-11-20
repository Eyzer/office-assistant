"""Excel 底层封装：openpyxl 常用 + zip 直接改部件（保真用）。"""
import os
import re
import shutil
import zipfile

import openpyxl
from openpyxl.utils import get_column_letter


def load(path, data_only=False, read_only=False):
    return openpyxl.load_workbook(path, data_only=data_only, read_only=read_only)


def save(wb, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return v


def is_blank_row(ws, row_idx):
    for c in ws[row_idx]:
        if c.value not in (None, ""):
            return False
    return True


def is_blank_col(ws, col_idx):
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=col_idx).value not in (None, ""):
            return False
    return True


def detect_header_row(ws, scan=10):
    """找表头行：扫描前 N 行，取非空单元格最多的一行。"""
    best, best_score = 1, -1
    for r in range(1, min(scan, ws.max_row) + 1):
        score = sum(1 for c in ws[r] if c.value not in (None, ""))
        if score > best_score:
            best, best_score = r, score
    return best


def header_values(ws, row_idx):
    return [cell_text(c.value) for c in ws[row_idx]]


def unmerge_and_fill(ws, start_row=1):
    """取消合并，把左上角值填充到原区间。"""
    ranges = list(ws.merged_cells.ranges)
    for rng in ranges:
        top = ws.cell(row=rng.min_row, column=rng.min_col).value
        ws.unmerge_cells(str(rng))
        for r in range(rng.min_row, rng.max_row + 1):
            if r < start_row:
                continue
            for c in range(rng.min_col, rng.max_col + 1):
                ws.cell(row=r, column=c).value = top
    return len(ranges)


def merge_same_values(ws, col_idx, start_row=1):
    """同一列相邻同值单元格合并。返回合并次数。"""
    count = 0
    r = start_row
    max_r = ws.max_row
    while r <= max_r:
        v = ws.cell(row=r, column=col_idx).value
        if v in (None, ""):
            r += 1
            continue
        end = r
        while end + 1 <= max_r and ws.cell(row=end + 1, column=col_idx).value == v:
            end += 1
        if end > r:
            ws.merge_cells(start_row=r, start_column=col_idx, end_row=end, end_column=col_idx)
            count += 1
        r = end + 1
    return count


def column_widths(ws):
    return {get_column_letter(i): d.width for i, d in ws.column_dimensions.items() if d.width}


def copy_dimensions(src_ws, dst_ws):
    for k, d in src_ws.column_dimensions.items():
        if d.width:
            dst_ws.column_dimensions[k].width = d.width
    for k, d in src_ws.row_dimensions.items():
        if d.height:
            dst_ws.row_dimensions[k].height = d.height


def safe_sheet_title(name, used=None):
    """清洗 sheet 名，Excel 限 31 字符、禁用若干符号。"""
    t = re.sub(r'[\\/*?:\[\]]', "_", str(name)).strip() or "Sheet"
    t = t[:31]
    used = used if used is not None else set()
    base, i = t, 1
    while t.lower() in used:
        suf = f"_{i}"
        t = base[: 31 - len(suf)] + suf
        i += 1
    used.add(t.lower())
    return t


def csv_rows(path):
    """读 csv，自动探编码与分隔符。返回 (headers, rows)。"""
    import csv
    from utils.encoding_utils import read_text
    text = read_text(path)
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",\t;")
        delim = dialect.delimiter
    except csv.Error:
        delim = ","
    reader = list(csv.reader(text.splitlines(), delimiter=delim))
    if not reader:
        return [], []
    return reader[0], reader[1:]


# ---------- zip 底层：图片与部件（保真） ----------

IMAGE_REL = re.compile(rb'<Relationship[^>]*Target="([^"]*media/[^"]+)"')


def list_images(xlsx_path):
    """列出 xlsx 内所有图片部件路径。"""
    out = []
    with zipfile.ZipFile(xlsx_path) as z:
        for n in z.namelist():
            if "/media/" in n:
                out.append(n)
    return out


def extract_images(xlsx_path, out_dir, prefix=""):
    """导出 xlsx 内所有图片，按原文件名。返回导出路径列表。"""
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    with zipfile.ZipFile(xlsx_path) as z:
        for n in z.namelist():
            if "/media/" in n:
                base = os.path.basename(n)
                dst = os.path.join(out_dir, prefix + base)
                if os.path.exists(dst):
                    stem, ext = os.path.splitext(dst)
                    i = 1
                    while os.path.exists(dst):
                        dst = f"{stem}_{i}{ext}"
                        i += 1
                with z.open(n) as src, open(dst, "wb") as f:
                    shutil.copyfileobj(src, f)
                saved.append(dst)
    return saved
