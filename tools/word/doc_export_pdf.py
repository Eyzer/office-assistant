"""批量文档转 PDF（调用 LibreOffice）。"""
import os
import shutil

from utils import office_convert
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import ensure_dir, unique_path

META = {
    "id": "doc_export_pdf",
    "name": "文档转 PDF",
    "desc": "调用本机 LibreOffice 无头模式批量转 PDF；未安装时提示。",
    "fields": [
        {"key": "timeout", "label": "单文件超时(秒)", "type": "number", "default": 180},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    if not office_convert.find_soffice():
        msg = "未找到 LibreOffice，请先安装：sudo apt install libreoffice"
        ctx.error(msg)
        return {"success_count": 0, "fail_count": 1,
                "log_messages": [msg], "output_files": []}

    files = collect_files(inputs, WORD_EXT)
    ensure_dir(output_dir)
    timeout = int(params.get("timeout") or 180)
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            tmp = ensure_dir(os.path.join(output_dir, "_tmp_pdf"))
            pdf = office_convert.to_pdf(f, tmp, timeout)
            dst = unique_path(os.path.join(
                output_dir, os.path.splitext(os.path.basename(f))[0] + ".pdf"))
            shutil.move(pdf, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)} -> {os.path.basename(dst)}")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    for d in (os.path.join(output_dir, "_tmp_pdf"),):
        shutil.rmtree(d, ignore_errors=True)

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
