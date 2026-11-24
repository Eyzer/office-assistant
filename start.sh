#!/bin/sh
# office办公助手 · Linux 启动脚本
#
# 用法：
#   ./start.sh                    仅本机访问，随机端口
#   ./start.sh --host 0.0.0.0     允许局域网访问（服务器上测试用这个）
#   ./start.sh --port 8765        固定端口
#   ./start.sh --token mytoken    固定访问令牌
#   ./start.sh --no-browser       不自动开浏览器
#
# 任何参数都会原样传给 main.py。

set -e
DIR=$(cd "$(dirname "$0")" && pwd)
cd "$DIR"

# ---------- 找 Python ----------
PY=""
for c in python3 python3.11 python3.10 python3.9 python3.8 python3.7 python; do
    if command -v "$c" >/dev/null 2>&1; then
        if "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 7) else 1)' 2>/dev/null; then
            PY="$c"
            break
        fi
    fi
done
if [ -z "$PY" ]; then
    echo "找不到 Python 3.7 或更高版本。" >&2
    echo "麒麟 / Debian 系可执行：sudo apt install python3" >&2
    exit 1
fi

# ---------- 内置依赖目录 ----------
# vendor/ 是离线依赖包解出来的位置；只读目录安装时会退到用户目录，两处都挂上。
VENDOR="$DIR/vendor"
USER_VENDOR="$HOME/.local/share/office-assistant/vendor"
export PYTHONPATH="$VENDOR${PYTHONPATH:+:$PYTHONPATH}"
if [ -d "$USER_VENDOR" ]; then
    export PYTHONPATH="$PYTHONPATH:$USER_VENDOR"
fi

# ---------- 必需依赖 ----------
# 顺序：先用随包带的离线 wheel（内网唯一可行），再试 apt，最后试 pip。
deps_ok() { "$PY" -c 'import openpyxl, docx' >/dev/null 2>&1; }

if ! deps_ok; then
    if [ -f "$DIR/offline-deps/install-deps.py" ] && [ -d "$DIR/offline-deps/wheels" ]; then
        echo "缺少依赖，使用内置离线包安装（不需要联网）……"
        "$PY" "$DIR/offline-deps/install-deps.py" || true
    fi
fi

if ! deps_ok; then
    echo "离线包未覆盖本机环境，尝试系统源 / pip……"
    if command -v apt-get >/dev/null 2>&1 && [ "$(id -u)" = "0" ]; then
        apt-get install -y python3-openpyxl python3-docx >/dev/null 2>&1 || true
    elif command -v apt-get >/dev/null 2>&1; then
        sudo apt-get install -y python3-openpyxl python3-docx >/dev/null 2>&1 || true
    fi
    if ! deps_ok; then
        "$PY" -m pip install -r requirements.txt 2>/dev/null || true
    fi
fi

if ! deps_ok; then
    echo "" >&2
    echo "依赖仍然缺失。内网机器请按下面任一条处理：" >&2
    echo "  1) 拷入对应架构与 Python 版本的离线包后执行：" >&2
    echo "     $PY $DIR/offline-deps/install-deps.py" >&2
    echo "  2) 有安装源时：sudo apt install python3-openpyxl python3-docx" >&2
    echo "" >&2
    echo "本机情况：$("$PY" -c 'import platform,sys;print(sys.version.split()[0], platform.machine())')" >&2
    exit 1
fi

# ---------- sjz 解密能力提示 ----------
if ! command -v openssl >/dev/null 2>&1; then
    if ! "$PY" -c 'import cryptography' 2>/dev/null && ! "$PY" -c 'import Crypto' 2>/dev/null; then
        echo ""
        echo "提示：没找到 openssl，也没有 cryptography / pycryptodome。"
        echo "      sjz 解密仍可用，但会退回纯 Python 实现，大文件很慢（约 1MB/s）。"
        echo "      建议安装：sudo apt install openssl"
        echo ""
    fi
fi

exec "$PY" main.py "$@"
