此目录用于内置第三方依赖（随 deb 分发）。

打包脚本包装时会执行 packaging/fetch-vendor.sh，把
  qrcode / Pillow / pypinyin / pypdf
的 aarch64 wheel 解包到这里。目录内容不纳入版本管理。
