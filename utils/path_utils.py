"""路径与输出命名工具。绝不覆盖已有文件。"""
import os
import re


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def safe_name(name):
    """去掉文件名里的非法字符。"""
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(name))
    return name.strip() or "未命名"


def split_ext(path):
    base = os.path.basename(path)
    stem, ext = os.path.splitext(base)
    return stem, ext


def unique_path(path):
    """目标已存在时追加 _1 _2 ...，永不覆盖。"""
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    i = 1
    while True:
        cand = f"{stem}_{i}{ext}"
        if not os.path.exists(cand):
            return cand
        i += 1


def out_path(src, output_dir, suffix="_处理后", ext=None, subdir=None):
    """按源文件名生成输出路径，自动防覆盖。"""
    ensure_dir(output_dir)
    if subdir:
        output_dir = ensure_dir(os.path.join(output_dir, safe_name(subdir)))
    stem, src_ext = split_ext(src)
    ext = ext or src_ext
    return unique_path(os.path.join(output_dir, f"{safe_name(stem)}{suffix}{ext}"))


def batch_out_dir(src, output_dir, suffix="_拆分"):
    """拆分类任务的独立目录，按源文件名建子目录。"""
    stem, _ = split_ext(src)
    return ensure_dir(os.path.join(output_dir, f"{safe_name(stem)}{suffix}"))
