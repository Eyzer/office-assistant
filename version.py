"""应用信息：名称、版本、作者。改这里就能改全站显示。"""

APP_NAME = "office办公助手"
APP_NAME_EN = "office-assistant"
APP_TAGLINE = "文档与表格批量处理"
VERSION = "1.3.2"

# 作者信息（要换人改这几行即可）
AUTHOR = "Eyzer"
EMAIL = "eyzer@qq.com"
HOMEPAGE = "https://github.com/Eyzer"
LICENSE = "MIT"

DESCRIPTION = "脱离 Office 宿主的文档 / 表格批量处理工具，全部生成新副本，不改源文件。"

# 前端与后端握手用的接口版本号。改了 API 就加一，
# 老浏览器页面会据此提示「后台服务版本过旧，请重启程序」。
API_VERSION = 3

# 关于页展示的第三方组件（名称, 用途, 许可）
THIRD_PARTY = [
    ("openpyxl", "xlsx 读写", "MIT"),
    ("python-docx", "docx 读写", "MIT"),
    ("Pillow", "图片处理", "HPND"),
    ("qrcode", "二维码生成", "BSD"),
    ("pypinyin", "汉字转拼音", "MIT"),
    ("pypdf", "PDF 拆分合并", "BSD"),
    ("python-barcode", "条形码生成", "MIT"),
    ("lunardate", "农历换算", "MIT"),
    ("OpenSSL", "sjz 密文 AES-256 解密", "Apache-2.0"),
]


def about():
    """给 /api/about 用。"""
    import platform
    import sys
    return {
        "name": APP_NAME,
        "name_en": APP_NAME_EN,
        "tagline": APP_TAGLINE,
        "version": VERSION,
        "author": AUTHOR,
        "email": EMAIL,
        "homepage": HOMEPAGE,
        "license": LICENSE,
        "description": DESCRIPTION,
        "api_version": API_VERSION,
        "third_party": [{"name": a, "use": b, "license": c} for a, b, c in THIRD_PARTY],
        "python": "%s %s" % (platform.python_implementation(), platform.python_version()),
        "platform": "%s / %s" % (platform.system(), platform.machine()),
        "executable": sys.executable,
    }
