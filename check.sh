#!/bin/sh
# 自检：Python 版本、依赖完整性、离线包覆盖情况、sjz 解密后端
DIR=$(cd "$(dirname "$0")" && pwd)
cd "$DIR"

PY=$(command -v python3 || command -v python)
if [ -z "$PY" ]; then
    echo "找不到 python3。" >&2
    exit 1
fi

export PYTHONPATH="$DIR/vendor${PYTHONPATH:+:$PYTHONPATH}"
if [ -d "$HOME/.local/share/office-assistant/vendor" ]; then
    export PYTHONPATH="$PYTHONPATH:$HOME/.local/share/office-assistant/vendor"
fi

echo "Python: $("$PY" -V 2>&1)  ->  $PY"
"$PY" - <<'EOF'
import os, platform, sys

ok = True

# 必需 / 可选分开看：缺可选只影响部分功能，缺必需整个程序用不了
REQUIRED = [("openpyxl", "openpyxl"), ("docx", "python-docx"), ("lxml", "lxml")]
OPTIONAL = [("PIL", "Pillow"), ("qrcode", "qrcode"), ("pypinyin", "pypinyin"),
            ("pypdf", "pypdf"), ("barcode", "python-barcode"),
            ("lunardate", "lunardate")]

print("平台: %s / Python %s" % (platform.machine(), sys.version.split()[0]))
print("")

missing = []
for mod, cn in REQUIRED + OPTIONAL:
    try:
        m = __import__(mod)
        ver = getattr(m, "__version__", "") or "已安装"
        tag = "必需" if (mod, cn) in REQUIRED else "可选"
        print("%-14s %-12s %s" % (cn, tag, ver))
    except Exception as e:
        missing.append((mod, cn, (mod, cn) in REQUIRED))
        print("%-14s %-12s 缺失" % (cn, "必需" if (mod, cn) in REQUIRED else "可选"))

if any(m[2] for m in missing):
    ok = False

# 离线包能不能补上本机缺的
root = os.path.join(os.path.abspath("."), "offline-deps", "wheels")
combo = "py%d%d-%s" % (sys.version_info[0], sys.version_info[1], platform.machine())
print("")
if missing:
    print("离线包目录: %s" % root)
    if os.path.isdir(root):
        have = sorted(d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d)))
        print("  已包含: %s" % "、".join(have))
        if combo in have:
            print("  本机组合 %s 在包里，执行下面这条即可补齐：" % combo)
            print("    python3 offline-deps/install-deps.py")
        else:
            print("  ! 包里没有 %s 这组，需要换对应离线包或在本机用 apt 安装" % combo)
    else:
        print("  没有 offline-deps/wheels 目录")

# sjz 解密后端
print("")
try:
    from utils import sjz
    info = sjz.self_check()
    print("sjz 解密后端  " + info["backend"])
    print("AES 标准向量  " + info["aes_vector"])
    print("密钥长度      %d 字节" % info["key_len"])
    print("FILL 判据     " + info["fill_hex"])
    if info["aes_vector"] != "PASS":
        ok = False
        print("AES 自检未通过，sjz 解密不可用！")
except Exception as e:
    ok = False
    print("sjz 模块加载失败：%s" % e)

print("")
if not missing:
    print("结论：就绪（依赖齐全）")
elif any(m[2] for m in missing):
    print("结论：不可用 —— 缺必需依赖 %s" % "、".join(m[1] for m in missing if m[2]))
else:
    print("结论：基本可用，缺可选依赖 %s（相关功能会报缺依赖）"
          % "、".join(m[1] for m in missing))
sys.exit(0 if ok else 1)
EOF
