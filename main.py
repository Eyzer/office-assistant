#!/usr/bin/env python3
"""office办公助手 入口：起本地服务并打开浏览器。"""
import argparse
import os
import secrets
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

# 第三方库放在 vendor/（deb 安装包或离线依赖包解出来的），优先加载。
# 装了只读目录时会退到用户目录，两处都挂上。
for _vendor in (os.path.join(BASE, "vendor"),
                os.path.join(os.path.expanduser("~"), ".local", "share",
                             "office-assistant", "vendor")):
    if os.path.isdir(_vendor) and _vendor not in sys.path:
        sys.path.insert(0, _vendor)

import version as app_info  # noqa: E402
from utils import config_utils  # noqa: E402


def main():
    ap = argparse.ArgumentParser(
        description="%s v%s — %s" % (app_info.APP_NAME, app_info.VERSION, app_info.APP_TAGLINE),
        epilog="作者：%s <%s>" % (app_info.AUTHOR, app_info.EMAIL),
    )
    ap.add_argument("--port", type=int, default=0, help="端口，0=随机")
    ap.add_argument("--host", default="127.0.0.1",
                    help="监听地址，127.0.0.1=仅本机；0.0.0.0=允许局域网访问")
    ap.add_argument("--no-browser", action="store_true", help="不自动打开浏览器")
    ap.add_argument("--token", default="", help="自定义访问令牌")
    ap.add_argument("-v", "--version", action="version",
                    version="%s %s" % (app_info.APP_NAME, app_info.VERSION))
    args = ap.parse_args()

    config_utils.init_config()
    token = args.token or secrets.token_urlsafe(16)

    # 没有图形环境（服务器、SSH 会话）就别试图拉起浏览器，否则只会留下报错噪音
    no_browser = args.no_browser
    if not no_browser and sys.platform.startswith("linux"):
        if not (os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")):
            no_browser = True

    os.chdir(BASE)
    from server import serve
    serve(host=args.host, port=args.port, token=token, open_browser=not no_browser)


if __name__ == "__main__":
    main()
