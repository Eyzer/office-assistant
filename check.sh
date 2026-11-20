#!/bin/sh
# 自检：Python 版本、必需依赖、sjz 解密后端
DIR=$(cd "$(dirname "$0")" && pwd)
cd "$DIR"
PY=$(command -v python3 || command -v python)
echo "Python: $("$PY" -V 2>&1)  ->  $PY"
echo ""
"$PY" - <<'EOF'
import json, sys
ok = True
try:
    import openpyxl
    print("openpyxl      %s" % openpyxl.__version__)
except Exception as e:
    ok = False
    print("openpyxl      缺失（%s）" % e)
try:
    import docx
    print("python-docx   %s" % getattr(docx, "__version__", "已安装"))
except Exception as e:
    ok = False
    print("python-docx   缺失（%s）" % e)

from utils import sjz
info = sjz.self_check()
print("")
print("sjz 解密后端  " + info["backend"])
print("AES 标准向量  " + info["aes_vector"])
print("密钥长度      %d 字节" % info["key_len"])
print("FILL 判据     " + info["fill_hex"])
print("")
if info["aes_vector"] != "PASS":
    ok = False
    print("AES 自检未通过，sjz 解密不可用！")
print("结论：" + ("就绪" if ok else "有问题，见上"))
sys.exit(0 if ok else 1)
EOF
