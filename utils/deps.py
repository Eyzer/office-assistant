"""第三方依赖加载：优先系统包，回退内置 vendor/。

目标机是离线的麒麟 V10（aarch64），系统里未必装了 python3-pil / python3-pypinyin。
打包时把对应 wheel 解包进 vendor/，随 deb 一起分发，装完即可用，不需要联网。
vendor 放在 sys.path 末尾，系统里已有的版本优先，避免顶掉发行版自带的包。
"""
import importlib
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VENDOR = os.path.join(BASE, "vendor")

# 模块名 -> (中文名, 用途, 系统包名)
KNOWN = {
    "openpyxl": ("openpyxl", "读写 xlsx", "python3-openpyxl"),
    "docx": ("python-docx", "读写 docx", "python3-docx"),
    "lxml": ("lxml", "docx 底层 XML", "python3-lxml"),
    "PIL": ("Pillow", "图片处理", "python3-pil"),
    "qrcode": ("qrcode", "生成二维码", "python3-qrcode"),
    "pypinyin": ("pypinyin", "汉字转拼音", "python3-pypinyin"),
    "pypdf": ("pypdf", "PDF 拆分合并", "python3-pypdf2"),
}


class MissingDep(RuntimeError):
    """依赖缺失，消息可直接展示给用户。"""

    def __init__(self, name, hint=""):
        cn = KNOWN.get(name, (name, "", ""))[0]
        msg = "缺少依赖 %s" % cn
        if hint:
            msg += "（%s）" % hint
        pkg = KNOWN.get(name, ("", "", ""))[2]
        if pkg:
            msg += "。安装：sudo apt install %s，或将 wheel 解包到 vendor/ 目录" % pkg
        else:
            msg += "。请将 wheel 解包到 vendor/ 目录"
        RuntimeError.__init__(self, msg)


def ensure_vendor():
    """把 vendor/ 加入 sys.path（末尾，系统包优先）。"""
    if os.path.isdir(VENDOR) and VENDOR not in sys.path:
        sys.path.append(VENDOR)


def has(name):
    """依赖是否可用。不抛异常。"""
    ensure_vendor()
    try:
        importlib.import_module(name)
        return True
    except Exception:
        return False


def load(name, hint=""):
    """导入依赖，缺失时抛 MissingDep（消息友好）。"""
    ensure_vendor()
    try:
        return importlib.import_module(name)
    except Exception as exc:
        raise MissingDep(name, hint) from exc


def need(ctx, names):
    """批量检查依赖；缺了就写日志并返回 False，供工具提前退出。"""
    missing = [n for n in names if not has(n)]
    if not missing:
        return True
    for n in missing:
        ctx.error(str(MissingDep(n)))
    return False


def report():
    """给前端「环境自检」用的依赖状态表。"""
    ensure_vendor()
    out = []
    for name in ("openpyxl", "docx", "PIL", "qrcode", "pypinyin", "pypdf"):
        cn, use, pkg = KNOWN.get(name, (name, "", ""))
        ok = has(name)
        ver = ""
        if ok:
            try:
                mod = importlib.import_module(name)
                ver = str(getattr(mod, "__version__", "") or "")
            except Exception:
                ver = ""
        out.append({"module": name, "name": cn, "use": use,
                    "ok": ok, "version": ver, "apt": pkg})
    out.append({"module": "libreoffice", "name": "LibreOffice",
                "use": "转 PDF / PPT 转换", "ok": _which_soffice(),
                "version": "", "apt": "libreoffice-core"})
    return out


def _which_soffice():
    import shutil
    return bool(shutil.which("soffice") or shutil.which("libreoffice"))
