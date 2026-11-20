"""文件收集与过滤。"""
import os

EXCEL_EXT = {".xlsx", ".xlsm", ".csv"}
WORD_EXT = {".docx"}
ALL_DOC_EXT = EXCEL_EXT | WORD_EXT


def _is_temp(name):
    return name.startswith("~$") or name.startswith(".")


def collect_files(paths, exts=None):
    """把文件/文件夹混合列表摊平成文件列表。跳过隐藏与 ~$ 临时文件。"""
    out = []
    exts = {e.lower() for e in exts} if exts else None
    for p in paths:
        if not p:
            continue
        p = os.path.abspath(p)
        if os.path.isfile(p):
            if not _is_temp(os.path.basename(p)):
                if exts is None or os.path.splitext(p)[1].lower() in exts:
                    out.append(p)
        elif os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                dirs[:] = [d for d in dirs if not d.startswith(".")]
                for f in files:
                    if _is_temp(f):
                        continue
                    if exts is None or os.path.splitext(f)[1].lower() in exts:
                        out.append(os.path.join(root, f))
    # 去重保序
    seen, uniq = set(), []
    for f in out:
        if f not in seen:
            seen.add(f)
            uniq.append(f)
    return uniq


def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f}{unit}"
        n /= 1024
    return f"{n:.1f}TB"
