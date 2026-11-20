"""文本编码探测：中文环境 GBK/UTF-8 混杂是常态。"""
import codecs


def read_text(path, encodings=None):
    encodings = encodings or ("utf-8-sig", "utf-8", "gb18030", "big5", "latin-1")
    with open(path, "rb") as f:
        raw = f.read()
    for enc in encodings:
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace")


def detect(path):
    with open(path, "rb") as f:
        raw = f.read(4096)
    for enc in ("utf-8-sig", "utf-8", "gb18030", "big5"):
        try:
            raw.decode(enc)
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return "utf-8"


def write_text(path, text, encoding="utf-8-sig"):
    """默认带 BOM，Excel 打开不乱码。"""
    with open(path, "w", encoding=encoding, newline="") as f:
        f.write(text)
    return path
