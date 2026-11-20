"""批量把 PPT/Word/Excel/ODF 转 PDF（走 LibreOffice）。"""
import os
import shutil

from utils import batch
from utils import office_convert as oc

META = {
    "id": "ppt2pdf",
    "cat": "格式转换",
    "name": "Office 批量转 PDF",
    "desc": "PPT/Word/Excel/ODT 批量转 PDF，依赖 LibreOffice（sudo apt install libreoffice-core）。",
    "fields": [
        {"key": "keep_name", "label": "输出到子目录", "type": "checkbox", "default": False,
         "help": "勾选则按源文件名建子目录，避免同名覆盖"},
    ],
}

EXT = {".ppt", ".pptx", ".doc", ".docx", ".xls", ".xlsx", ".odt", ".ods", ".odp"}


def run_task(inputs, output_dir, params, ctx):
    if not oc.find_soffice():
        ctx.error("未找到 LibreOffice（soffice）。请安装：sudo apt install libreoffice-core")
        return {"success_count": 0, "fail_count": 0,
                "output_files": [], "log_messages": ctx.task["log"].lines[-50:]}
    sub_dir = bool(params.get("keep_name"))

    def handler(f, ctx):
        target = output_dir
        if sub_dir:
            target = os.path.join(output_dir, os.path.splitext(os.path.basename(f))[0])
        out = oc.to_pdf(f, target)
        if target != output_dir:
            base = os.path.basename(out)
            dst = os.path.join(output_dir, base)
            if os.path.exists(dst):
                stem, ext = os.path.splitext(dst)
                i = 1
                while os.path.exists(dst):
                    dst = "%s_%d%s" % (stem, i, ext)
                    i += 1
            shutil.move(out, dst)
            shutil.rmtree(target, ignore_errors=True)
            out = dst
        ctx.info("%s → PDF" % os.path.basename(f))
        return [out]

    return batch.process(inputs, ctx, handler, EXT)
