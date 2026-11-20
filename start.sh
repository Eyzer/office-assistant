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

# ---------- 必需依赖 ----------
need_install=0
"$PY" -c 'import openpyxl' 2>/dev/null || need_install=1
"$PY" -c 'import docx'     2>/dev/null || need_install=1

if [ "$need_install" = "1" ]; then
    echo "缺少 openpyxl 或 python-docx，尝试安装……"
    if command -v apt-get >/dev/null 2>&1 && [ "$(id -u)" = "0" ]; then
        apt-get install -y python3-openpyxl python3-docx >/dev/null 2>&1 || true
    elif command -v apt-get >/dev/null 2>&1; then
        sudo apt-get install -y python3-openpyxl python3-docx >/dev/null 2>&1 || true
    fi
    if ! "$PY" -c 'import openpyxl, docx' 2>/dev/null; then
        "$PY" -m pip install -r requirements.txt 2>/dev/null || true
    fi
    if ! "$PY" -c 'import openpyxl, docx' 2>/dev/null; then
        echo "" >&2
        echo "依赖仍然缺失，请手动安装其一：" >&2
        echo "  sudo apt install python3-openpyxl python3-docx" >&2
        echo "  $PY -m pip install -r requirements.txt" >&2
        exit 1
    fi
    echo "依赖已就绪。"
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
