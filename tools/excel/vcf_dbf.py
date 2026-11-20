"""导出通讯录 VCF 与 dBASE III 的 DBF（自实现，不依赖第三方）。"""
import datetime
import os

from utils import batch, excel_utils as xu
from utils.file_utils import EXCEL_EXT
from utils.path_utils import out_path, safe_name

META = {
    "id": "vcf_dbf",
    "cat": "导入与导出",
    "name": "导出 VCF 通讯录 / DBF",
    "desc": "按列生成手机通讯录 .vcf（可直接导入手机），或把表格导成 dBASE III 的 .dbf。",
    "fields": [
        {"key": "mode", "label": "导出格式", "type": "select", "default": "vcf",
         "options": [{"value": "vcf", "label": "VCF 通讯录"},
                     {"value": "dbf", "label": "DBF 数据表"}]},
        {"key": "name_col", "label": "姓名列", "type": "text", "default": "A"},
        {"key": "phone_col", "label": "手机列", "type": "text", "default": "B"},
        {"key": "email_col", "label": "邮箱列", "type": "text", "default": ""},
        {"key": "sheet", "label": "工作表", "type": "text", "default": "",
         "help": "留空=第一张"},
    ],
}


def _rows(ws, header_row):
    for r in range(header_row + 1, ws.max_row + 1):
        yield [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]


def build_vcf(f, ws, output_dir, nc, pc, ec, ctx):
    out = []
    for row in _rows(ws, 1):
        name = row[nc - 1] if len(row) >= nc else None
        phone = row[pc - 1] if len(row) >= pc else None
        if name in (None, "") and phone in (None, ""):
            continue
        lines = ["BEGIN:VCARD", "VERSION:3.0", "FN:%s" % (name or "")]
        if phone not in (None, ""):
            lines.append("TEL;TYPE=CELL:%s" % str(phone).strip())
        if ec and len(row) >= ec and row[ec - 1] not in (None, ""):
            lines.append("EMAIL:%s" % str(row[ec - 1]).strip())
        lines.append("END:VCARD")
        out.append("\r\n".join(lines))
        if len(out) >= 200:
            break
    if not out:
        raise ValueError("没读到有效数据（检查姓名列/手机列）")
    dst = out_path(f, output_dir, suffix="_通讯录", ext=".vcf")
    with open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write("\r\n".join(out) + "\r\n")
    return [dst]


def _dbf_value(v, ftype, length, dec=0):
    if v is None:
        return b" " * length
    if ftype == "N":
        try:
            num = float(str(v).replace(",", ""))
        except (TypeError, ValueError):
            num = 0.0
        s = ("%.*f" % (dec, num)) if dec else ("%.0f" % num)
        return s.rjust(length)[:length].encode("gb18030", "replace")
    if ftype == "D":
        try:
            if isinstance(v, (datetime.date, datetime.datetime)):
                s = v.strftime("%Y%m%d")
            else:
                s = str(v).replace("-", "").replace("/", "")[:8]
        except Exception:
            s = ""
        return s.ljust(length)[:length].encode("gb18030", "replace")
    s = str(v)
    return s.encode("gb18030", "replace")[:length].ljust(length, b" ")


def build_dbf(f, ws, output_dir, ctx):
    headers = [str(c.value).strip() if c.value is not None else "F%d" % i
               for i, c in enumerate(ws[1], 1)]
    if not headers:
        raise ValueError("第一行没有表头")
    data = [row for row in _rows(ws, 1) if any(v not in (None, "") for v in row)]
    if not data:
        raise ValueError("没有数据行")

    fields = []
    for i, h in enumerate(headers):
        col_vals = [r[i] if len(r) > i else None for r in data]
        nums = [v for v in col_vals if isinstance(v, (int, float))]
        if nums and len(nums) == len([v for v in col_vals if v not in (None, "")]):
            length = max(8, max(len(str(int(abs(v)))) for v in nums) + 4)
            fields.append((safe_name(h)[:10], "N", min(length, 18), 2))
        else:
            length = min(254, max(10, max((len(str(v or "")) for v in col_vals), default=10) + 2))
            fields.append((safe_name(h)[:10], "C", length, 0))

    today = datetime.date.today()
    header_len = 32 + 32 * len(fields) + 1
    rec_len = 1 + sum(f[2] for f in fields)
    buf = bytearray()
    buf += bytes([0x03, today.year - 1900, today.month, today.day])
    buf += len(data).to_bytes(4, "little")
    buf += header_len.to_bytes(2, "little")
    buf += rec_len.to_bytes(2, "little")
    buf += b"\x00" * 16          # 12~27 保留
    buf += bytes([0x00, 0x00])   # 28 MDX 标志 / 29 语言驱动
    buf += b"\x00\x00"           # 30~31 保留
    assert len(buf) == 32, len(buf)
    for name, ftype, length, dec in fields:
        nm = name.encode("gb18030", "replace")[:10]
        desc = bytearray(32)
        desc[0:len(nm)] = nm
        desc[11] = ord(ftype)
        desc[16] = length
        desc[17] = dec
        buf += desc
    buf += b"\x0d"

    for row in data:
        buf += b" "
        for i, (_, ftype, length, dec) in enumerate(fields):
            v = row[i] if len(row) > i else None
            buf += _dbf_value(v, ftype, length, dec)
    buf += b"\x1a"

    dst = out_path(f, output_dir, suffix="_导出", ext=".dbf")
    with open(dst, "wb") as fh:
        fh.write(buf)
    ctx.info("%s: %d 行 %d 字段" % (os.path.basename(f), len(data), len(fields)))
    return [dst]


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("mode") or "vcf"
    nc = batch.col_index(params.get("name_col") or "A", 1)
    pc = batch.col_index(params.get("phone_col") or "B", 2)
    ec = batch.col_index(params.get("email_col") or "", 0) if str(params.get("email_col") or "").strip() else 0
    sheet = params.get("sheet") or ""

    def handler(f, ctx):
        wb = xu.load(f, data_only=True)
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
        if mode == "vcf":
            return build_vcf(f, ws, output_dir, nc, pc, ec, ctx)
        return build_dbf(f, ws, output_dir, ctx)

    return batch.process(inputs, ctx, handler, EXCEL_EXT)
