#!/usr/bin/env bash
# 把第三方依赖解包进 vendor/（打 deb 用）。
#
# 优先用 offline-deps/wheels 里已下好的 wheel（离线可用）；
# 没有对应组合时再调 build-wheels.py 联网下载。
#
# 用法：
#   bash packaging/fetch-vendor.sh                  # 默认 py3.7 + aarch64
#   PYVER=3.9 ARCH=aarch64 bash packaging/fetch-vendor.sh
#   PYVER=3.9 ARCH=x86_64  bash packaging/fetch-vendor.sh
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

PYVER="${PYVER:-3.7}"
ARCH="${ARCH:-aarch64}"
COMB="py${PYVER//./}-${ARCH}"
WH="$ROOT/offline-deps/wheels"

command -v python3 >/dev/null || { echo "! 缺少 python3"; exit 2; }

if [ ! -d "$WH/$COMB" ] || [ -z "$(ls -A "$WH/$COMB" 2>/dev/null)" ]; then
    echo "== 本地没有 $COMB，尝试联网下载 =="
    python3 offline-deps/build-wheels.py --arch "$ARCH" --py "$PYVER"
fi

echo "== 解包 $COMB 到 vendor/ =="
python3 - "$WH" "$COMB" "$ROOT/vendor" <<'PY'
import os, sys, zipfile
wh, comb, dest = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(dest, exist_ok=True)
n = 0
for d in ("common", comb):
    p = os.path.join(wh, d)
    if not os.path.isdir(p):
        continue
    for f in sorted(os.listdir(p)):
        if not f.endswith(".whl"):
            continue
        with zipfile.ZipFile(os.path.join(p, f)) as z:
            z.extractall(dest)
        print("  +", f)
        n += 1
if not n:
    sys.exit("没有可解包的 wheel")
PY

find "$ROOT/vendor" -name "__pycache__" -type d -prune -exec rm -rf {} + 2>/dev/null || true
find "$ROOT/vendor" -name "tests" -type d -prune -exec rm -rf {} + 2>/dev/null || true

echo
echo "完成。vendor 大小：$(du -sh "$ROOT/vendor" | cut -f1)"
PYTHONPATH="$ROOT/vendor" python3 - <<'PY'
for m in ("openpyxl", "docx", "lxml", "PIL", "qrcode", "pypinyin", "pypdf"):
    try:
        __import__(m)
        print("  ok  ", m)
    except Exception as e:
        print("  FAIL", m, e)
PY
