#!/usr/bin/env bash
# 构建 office办公助手 的 arm64 deb（纯 Python，含离线依赖）
#
# 要求：
#   - arm64 环境（原生机，或 qemu-user 的 arm64 rootfs 内）
#   - python3、pip、dpkg-deb、unzip
#   - 目标机 python3 版本需与 --pyver 一致（默认 3.8，麒麟 V10 常见）
#
# 用法：
#   bash packaging/build-deb.sh
#   PYVER=3.7 bash packaging/build-deb.sh
#   ALLOW_EMULATED=1 bash packaging/build-deb.sh   # 在 qemu 模拟环境里构建
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

APP=office-assistant
# 版本/作者统一从 version.py 读，避免两处不一致
VERSION="$(sed -n 's/^VERSION = "\(.*\)"$/\1/p' version.py | head -1)"
AUTHOR="$(sed -n 's/^AUTHOR = "\(.*\)"$/\1/p' version.py | head -1)"
EMAIL="$(sed -n 's/^EMAIL = "\(.*\)"$/\1/p' version.py | head -1)"
APP_NAME="$(sed -n 's/^APP_NAME = "\(.*\)"$/\1/p' version.py | head -1)"
: "${VERSION:=1.0.0}"
: "${AUTHOR:=Eyzer}"
: "${EMAIL:=eyzer@qq.com}"
: "${APP_NAME:=office办公助手}"
DESKTOP_NAME="${APP}.desktop"
PYVER="${PYVER:-3.8}"
PKG="${APP}_${VERSION}_arm64"
BUILD="$ROOT/build"
STAGE="$BUILD/$PKG"

ARCH="$(uname -m)"
if [ "$ARCH" != "aarch64" ] && [ "${ALLOW_EMULATED:-0}" != "1" ]; then
  echo "! 当前架构 $ARCH，不是 aarch64。"
  echo "  请在 arm64 机器上运行，或设 ALLOW_EMULATED=1 在 qemu arm64 rootfs 内构建。"
  exit 2
fi

command -v dpkg-deb >/dev/null || { echo "! 缺少 dpkg-deb"; exit 2; }
command -v python3 >/dev/null || { echo "! 缺少 python3"; exit 2; }

echo "== [1/5] 清理 =="
rm -rf "$BUILD"
mkdir -p "$STAGE/DEBIAN" \
         "$STAGE/opt/$APP/vendor" \
         "$STAGE/usr/bin" \
         "$STAGE/usr/share/applications"

echo "== [2/5] 复制程序 =="
cd "$ROOT"
for item in main.py server.py version.py config.ini utils service tools web profiles; do
  if [ -e "$item" ]; then cp -r "$item" "$STAGE/opt/$APP/"; fi
done
# 仓库里已内置好的 vendor/ 先带上（fetch-vendor.sh 会覆盖成目标平台版本）
if [ -d vendor ]; then cp -r vendor "$STAGE/opt/$APP/"; fi
mkdir -p "$STAGE/opt/$APP/logs" "$STAGE/opt/$APP/output"
find "$STAGE/opt/$APP" -name '__pycache__' -type d -prune -exec rm -rf {} + 2>/dev/null || true

echo "== [3/5] 下载离线依赖（aarch64 / py$PYVER）=="
TMPW="$BUILD/wheels"
mkdir -p "$TMPW"
# 核心：openpyxl / python-docx 及其 XML 依赖（lxml 为二进制轮子，需匹配 ABI）
PYTAG="${PYVER//./}"
ABI="cp${PYTAG}m"
if [ "$PYTAG" -ge 38 ] 2>/dev/null; then ABI="cp${PYTAG}"; fi
python3 -m pip download \
  --only-binary=:all: \
  --platform manylinux2014_aarch64 \
  --python-version "$PYVER" \
  --implementation cp \
  --abi "$ABI" \
  -d "$TMPW" \
  openpyxl python-docx lxml et-xmlfile 2>&1 | tail -5 || {
    echo "! 依赖下载失败。可改为目标机 apt 安装 python3-openpyxl python3-lxml。"
    echo "  或在联网机上先下载到 $TMPW 后重跑本脚本。"
    exit 3
  }

# 增强依赖（二维码/拼音/PDF/图片）：内置进 vendor，目标机免联网
echo "-- 内置增强依赖 --"
if PYVER="$PYVER" bash packaging/fetch-vendor.sh >/tmp/fetch-vendor.log 2>&1; then
  rm -rf "$STAGE/opt/$APP/vendor"
  mkdir -p "$STAGE/opt/$APP/vendor"
  cp -r vendor/. "$STAGE/opt/$APP/vendor/"
  echo "   vendor 已内置：$(du -sh "$STAGE/opt/$APP/vendor" | cut -f1)"
else
  echo "! 增强依赖拉取失败，跳过（相关工具将提示缺依赖）"
  tail -5 /tmp/fetch-vendor.log || true
fi

for whl in "$TMPW"/*.whl; do
  [ -e "$whl" ] || continue
  python3 - "$whl" "$STAGE/opt/$APP/vendor" <<'PY'
import sys, zipfile
whl, dst = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(whl) as z:
    z.extractall(dst)
PY
done

echo "== [4/5] 写启动器与控制信息 =="
cat > "$STAGE/usr/bin/$APP" <<EOF
#!/bin/sh
exec python3 /opt/$APP/main.py "\$@"
EOF
chmod 755 "$STAGE/usr/bin/$APP"

install -m 644 "packaging/$DESKTOP_NAME" "$STAGE/usr/share/applications/$DESKTOP_NAME"

INSTALLED=$(du -sk "$STAGE" | cut -f1)
cat > "$STAGE/DEBIAN/control" <<EOF
Package: $APP
Version: $VERSION
Section: office
Priority: optional
Architecture: arm64
Installed-Size: $INSTALLED
Depends: python3 (>= 3.7)
Recommends: zenity | kdialog
Maintainer: $AUTHOR <$EMAIL>
Homepage: https://github.com/Eyzer
Description: $APP_NAME — 文档与表格批量处理
 独立桌面程序，离线批量处理 docx 与 xlsx，不修改源文件。
 通过本地网页界面操作，不上传外网。
EOF

cat > "$STAGE/DEBIAN/postinst" <<EOF
#!/bin/sh
set -e
mkdir -p /opt/$APP/logs /opt/$APP/output /opt/$APP/uploads
chmod -R 777 /opt/$APP/logs /opt/$APP/output /opt/$APP/uploads 2>/dev/null || true
# 旧包名（docbatchtool）升级过来的，顺手提示一下
exit 0
EOF
chmod 755 "$STAGE/DEBIAN/postinst"

cat > "$STAGE/DEBIAN/postrm" <<'EOF'
#!/bin/sh
exit 0
EOF
chmod 755 "$STAGE/DEBIAN/postrm"

echo "== [5/5] 打包 =="
cd "$BUILD"
dpkg-deb --build --root-owner-group "$PKG" "$ROOT/${PKG}.deb"
echo
echo "完成: $ROOT/${PKG}.deb"
echo "安装: sudo dpkg -i ${PKG}.deb"
echo "启动: $APP   （或从开始菜单打开 ${APP_NAME}）"
echo "旧版清理: sudo dpkg -r docbatchtool  （若装过旧包名）"
