"""LibreOffice 无头转换封装。目标机需装 libreoffice。"""
import os
import shutil
import subprocess

CANDIDATES = ("soffice", "libreoffice", "/usr/bin/soffice", "/opt/libreoffice/program/soffice")


def find_soffice():
    for c in CANDIDATES:
        p = shutil.which(c) if not c.startswith("/") else (c if os.path.exists(c) else None)
        if p:
            return p
    return None


def to_pdf(src, out_dir, timeout=180):
    """转换单个文件为 PDF，返回输出路径或抛异常。"""
    exe = find_soffice()
    if not exe:
        raise RuntimeError("未找到 LibreOffice（soffice），请先安装")
    os.makedirs(out_dir, exist_ok=True)
    cmd = [exe, "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", out_dir, src]
    env = dict(os.environ)
    env.setdefault("HOME", out_dir)  # 避免无 HOME 时空配置报错
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                   timeout=timeout, env=env)
    base = os.path.splitext(os.path.basename(src))[0] + ".pdf"
    out = os.path.join(out_dir, base)
    if not os.path.exists(out):
        raise RuntimeError("转换未产出 PDF")
    return out
