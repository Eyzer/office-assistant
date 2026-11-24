#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""离线安装第三方依赖：把 offline-deps/wheels 里匹配本机的 wheel 解包到 vendor/。

内网 / 麒麟离线机器上没有 apt 源也没有 pip 源，这个脚本只用标准库，
把预置的 .whl 直接解压即可，不联网、不调用 pip。

用法：
    python3 offline-deps/install-deps.py            # 安装到 <程序目录>/vendor
    python3 offline-deps/install-deps.py --check    # 只检查，不写文件
    python3 offline-deps/install-deps.py --list     # 列出包里有哪些组合
    python3 offline-deps/install-deps.py --dest /opt/xx/vendor

退出码：0 成功 / 1 依赖不完整 / 2 没有匹配本机的 wheel 组合
"""

import argparse
import importlib
import os
import platform
import shutil
import sys
import zipfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WHEEL_ROOT = os.path.join(BASE, "offline-deps", "wheels")

# 装完要能导入才算成功（模块名, 中文名）
CHECKS = [
    ("openpyxl", "openpyxl"),
    ("docx", "python-docx"),
    ("lxml", "lxml"),
    ("PIL", "Pillow"),
    ("qrcode", "qrcode"),
    ("pypinyin", "pypinyin"),
    ("pypdf", "pypdf"),
    ("barcode", "python-barcode"),
    ("lunardate", "lunardate"),
]


def arch_tag():
    """归一化成 wheels/ 目录用的架构名。"""
    m = (platform.machine() or "").lower()
    if m in ("aarch64", "arm64"):
        return "aarch64"
    if m in ("x86_64", "amd64", "intel64"):
        return "x86_64"
    return m


def py_tag():
    return "py%d%d" % (sys.version_info[0], sys.version_info[1])


def wheel_dirs_for_this_box():
    """本机应该用的 wheel 目录列表（纯 Python 的 common 永远带）。"""
    return ["common", "%s-%s" % (py_tag(), arch_tag())]


def list_combos():
    if not os.path.isdir(WHEEL_ROOT):
        return []
    out = []
    for name in sorted(os.listdir(WHEEL_ROOT)):
        d = os.path.join(WHEEL_ROOT, name)
        if os.path.isdir(d):
            n = len([f for f in os.listdir(d) if f.endswith(".whl")])
            out.append((name, n))
    return out


def pick_dest(explicit=""):
    """优先装到程序目录的 vendor/；只读时退到用户目录。"""
    if explicit:
        return os.path.abspath(explicit)
    cand = os.path.join(BASE, "vendor")
    try:
        os.makedirs(cand, exist_ok=True)
        probe = os.path.join(cand, ".write-test")
        with open(probe, "w") as f:
            f.write("ok")
        os.remove(probe)
        return cand
    except OSError:
        home = os.path.expanduser("~")
        fallback = os.path.join(home, ".local", "share", "office-assistant", "vendor")
        os.makedirs(fallback, exist_ok=True)
        return fallback


def unzip_wheels(dirs, dest):
    os.makedirs(dest, exist_ok=True)
    done = []
    for d in dirs:
        path = os.path.join(WHEEL_ROOT, d)
        if not os.path.isdir(path):
            continue
        for f in sorted(os.listdir(path)):
            if not f.endswith(".whl"):
                continue
            with zipfile.ZipFile(os.path.join(path, f)) as z:
                z.extractall(dest)
            done.append(f)
    # 清掉字节码与测试，省空间也避免版本错配
    for root, subdirs, _ in os.walk(dest):
        for s in list(subdirs):
            if s == "__pycache__" or s == "tests":
                shutil.rmtree(os.path.join(root, s), ignore_errors=True)
                subdirs.remove(s)
    return done


def verify(dest):
    sys.path.insert(0, dest)
    rows = []
    for mod, cn in CHECKS:
        try:
            m = importlib.import_module(mod)
            ver = getattr(m, "__version__", "") or ""
            rows.append((cn, True, str(ver)))
        except Exception as exc:
            rows.append((cn, False, str(exc)[:60]))
    return rows


def print_rows(rows):
    width = max(len(r[0]) for r in rows) if rows else 8
    for cn, ok, extra in rows:
        print("  %s  %-*s %s" % ("ok  " if ok else "缺失", width, cn, extra))


def main():
    ap = argparse.ArgumentParser(description="离线安装第三方依赖到 vendor/")
    ap.add_argument("--dest", default="", help="解包目标目录，默认 <程序目录>/vendor")
    ap.add_argument("--check", action="store_true", help="只检查能否导入，不写文件")
    ap.add_argument("--list", action="store_true", help="列出包里包含的 Python/架构组合")
    args = ap.parse_args()

    if args.list:
        combos = list_combos()
        if not combos:
            print("没有找到 offline-deps/wheels 目录")
            return 2
        print("离线包内容（目录 -> wheel 数量）：")
        for name, n in combos:
            print("  %-18s %d" % (name, n))
        print("\n本机需要：%s-%s" % (py_tag(), arch_tag()))
        return 0

    dest = pick_dest(args.dest)

    if args.check:
        rows = verify(dest)
        print("依赖检查（vendor: %s）：" % dest)
        print_rows(rows)
        return 0 if all(r[1] for r in rows) else 1

    dirs = wheel_dirs_for_this_box()
    have = [d for d in dirs if os.path.isdir(os.path.join(WHEEL_ROOT, d))]
    if not have:
        print("离线包里没有适配本机的依赖组。")
        print("本机是：%s / %s（需要目录 %s）" % (py_tag(), arch_tag(),
                                                 " 或 ".join(dirs)))
        combos = list_combos()
        if combos:
            print("包里现有：%s" % "、".join(n for n, _ in combos))
        print("换对应组合的离线包，或按 DEPLOY-Linux.md 的说明在能联网的机器上重新生成。")
        return 2

    print("本机：%s / %s" % (py_tag(), arch_tag()))
    print("安装到：%s" % dest)
    done = unzip_wheels(have, dest)
    for f in done:
        print("  + " + f)

    rows = verify(dest)
    print("\n安装结果：")
    print_rows(rows)
    ok = all(r[1] for r in rows)
    print("\n结论：%s" % ("就绪" if ok else "仍有缺失"))
    if ok and dest != os.path.join(BASE, "vendor"):
        print("注意：装到了用户目录，启动时需带上 PYTHONPATH=%s" % dest)
        print("     start.sh 已自动包含该路径。")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
