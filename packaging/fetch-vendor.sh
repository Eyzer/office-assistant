#!/usr/bin/env bash
# 把内置第三方依赖拉到 vendor/（默认目标：aarch64 / Python 3.7，麒麟 V10 SP1）
#
# 用法：
#   bash packaging/fetch-vendor.sh              # 默认 3.7 + manylinux2014_aarch64
#   PYVER=3.8 bash packaging/fetch-vendor.sh
#   PLATFORM=manylinux_2_17_aarch64 PYVER=3.9 bash packaging/fetch-vendor.sh
#
# 产出：vendor/ 下直接是模块目录（qrcode/、PIL/、pypinyin/、pypdf/）
# 说明：纯 Python 包（qrcode/pypinyin/pypdf）任何平台通用；
#       二进制包（Pillow）必须匹配目标 CPU 与 Python ABI。
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

PYVER="${PYVER:-3.7}"
PLATFORM="${PLATFORM:-manylinux2014_aarch64}"
DEST="$ROOT/vendor"
TMP="$ROOT/build/vendor-wheels"

# 版本按「最低支持 py3.7」挑，避免目标机 Python 太老装不上
PKGS=(
  "qrcode==7.4.2"
  "pypinyin==0.49.0"
  "pypdf==3.17.4"
  "pillow==9.5.0"
  "python-barcode==0.15.1"
  "lunardate==0.2.2"
)

command -v python3 >/dev/null || { echo "! 缺少 python3"; exit 2; }

PYTAG="${PYVER//./}"
ABI="cp${PYTAG}m"
if [ "$PYTAG" -ge 38 ] 2>/dev/null; then ABI="cp${PYTAG}"; fi

echo "== 拉取依赖：py$PYVER / $PLATFORM =="
rm -rf "$TMP"; mkdir -p "$TMP" "$DEST"

if ! python3 -m pip download --only-binary=:all: \
      --platform "$PLATFORM" --python-version "$PYVER" \
      --implementation cp --abi "$ABI" \
      -d "$TMP" "${PKGS[@]}" 2>&1 | tail -6; then
  echo "! 带 abi 下载失败，去掉 abi 重试…"
  python3 -m pip download --only-binary=:all: \
      --platform "$PLATFORM" --python-version "$PYVER" \
      -d "$TMP" "${PKGS[@]}" 2>&1 | tail -6
fi

echo "== 解包到 vendor/ =="
python3 - "$TMP" "$DEST" <<'PY'
import os, sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
wheels = [f for f in os.listdir(src) if f.endswith(".whl")]
if not wheels:
    sys.exit("没有下载到 wheel")
for w in wheels:
    with zipfile.ZipFile(os.path.join(src, w)) as z:
        z.extractall(dst)
    print("  +", w)
PY

# 去掉 dist-info 里的 bytecode 与测试，减小体积
find "$DEST" -name "__pycache__" -type d -prune -exec rm -rf {} + 2>/dev/null || true
find "$DEST" -name "tests" -type d -prune -exec rm -rf {} + 2>/dev/null || true

echo
echo "完成。vendor 大小：$(du -sh "$DEST" | cut -f1)"
python3 - <<'PY'
import sys
sys.path.append("vendor")
for m in ("qrcode", "pypinyin", "pypdf", "PIL"):
    try:
        __import__(m)
        print("  ok  ", m)
    except Exception as e:
        print("  FAIL", m, e)
PY
