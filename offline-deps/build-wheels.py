#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成离线依赖包：按「Python 版本 + 架构」下载 wheel 到 offline-deps/wheels/。

在有网的机器上跑一次，产物随程序一起拷进内网，目标机用 install-deps.py 解包。

用法：
    python3 offline-deps/build-wheels.py                       # 默认 aarch64 + x86_64，py3.7~3.11
    python3 offline-deps/build-wheels.py --arch aarch64
    python3 offline-deps/build-wheels.py --py 3.9 3.10
    python3 offline-deps/build-wheels.py --arch x86_64 --py 3.10 --clean

说明：
- 只下 manylinux2014（glibc 2.17 基线），麒麟 V10 的 glibc 2.28 完全覆盖；
  不用 manylinux_2_28/2_34，避免老 SP 包 glibc 不够新。
- 纯 Python 包（openpyxl / qrcode / pypinyin / pypdf / barcode / lunardate）
  放 wheels/common，任何架构与 Python 版本通用。
- 二进制包（lxml / Pillow）按 Python ABI 分开，放 wheels/pyXX-<arch>。
- Python 3.7 用 python-docx 1.0.1：1.1.x 依赖 typing-extensions>=4.9，而后者不支持 3.7。
"""

import argparse
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WHEEL_ROOT = os.path.join(BASE, "offline-deps", "wheels")

# 默认走清华镜像，比直连 PyPI 快一个数量级；失败自动回落官方源
INDEX = "https://pypi.tuna.tsinghua.edu.cn/simple"

# 纯 Python，跟架构无关
PURE = [
    "openpyxl==3.1.2",
    "et-xmlfile==1.1.0",
    "qrcode==7.4.2",
    "pypinyin==0.49.0",
    "pypdf==3.17.4",
    "python-barcode==0.15.1",
    "lunardate==0.2.2",
]


def binary_pkgs(pyver):
    """按 Python 版本挑二进制包版本，保证有 manylinux2014 轮子且支持该版本。"""
    v = tuple(int(x) for x in pyver.split("."))
    if v == (3, 7):
        # 1.1.x 要求 typing-extensions>=4.9，而 4.9 起不再支持 3.7；
        # 1.0.1 不锁版本，配 4.7.1（最后一个支持 3.7 的 typing-extensions）。
        return ["python-docx==1.0.1", "typing-extensions==4.7.1",
                "lxml==4.9.4", "pillow==9.5.0"]
    if v == (3, 8):
        return ["python-docx==1.1.2", "typing-extensions==4.12.2",
                "lxml==4.9.4", "pillow==9.5.0"]
    if v in ((3, 9), (3, 10)):
        return ["python-docx==1.1.2", "typing-extensions==4.12.2",
                "lxml==4.9.4", "pillow==9.5.0"]
    if v == (3, 11):
        return ["python-docx==1.1.2", "typing-extensions==4.12.2",
                "lxml==4.9.4", "pillow==9.5.0"]
    if v == (3, 12):
        return ["python-docx==1.1.2", "typing-extensions==4.12.2",
                "lxml==5.2.2", "pillow==10.3.0"]
    return None


def abi_of(pyver):
    tag = pyver.replace(".", "")
    return "cp%sm" % tag if pyver == "3.7" else "cp%s" % tag


def download(pkgs, dest, arch, pyver, abi=True, index=INDEX):
    plat = "manylinux2014_%s" % arch
    cmd = [sys.executable, "-m", "pip", "download", "--only-binary=:all:",
           "--platform", plat, "--python-version", pyver,
           "--implementation", "cp", "-d", dest]
    if index:
        cmd += ["-i", index]
    if abi:
        cmd += ["--abi", abi_of(pyver)]
    cmd += pkgs
    print("  $ pip download (%s / py%s)" % (arch, pyver))
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        if index:
            print("    镜像失败，换官方源重试…")
            return download(pkgs, dest, arch, pyver, abi=abi, index=None)
        if abi:
            print("    带 abi 失败，去掉 abi 重试…")
            return download(pkgs, dest, arch, pyver, abi=False, index=None)
        print("    失败：")
        for line in r.stdout.splitlines()[-6:]:
            print("      " + line)
        return False
    return True


def dir_size_mb(path):
    total = 0
    for root, _, files in os.walk(path):
        for f in files:
            total += os.path.getsize(os.path.join(root, f))
    return total / 1048576.0


def main():
    ap = argparse.ArgumentParser(description="下载离线依赖 wheel")
    ap.add_argument("--arch", nargs="+", default=["aarch64", "x86_64"],
                    help="目标架构：aarch64 x86_64")
    ap.add_argument("--py", nargs="+", default=["3.7", "3.8", "3.9", "3.10", "3.11"],
                    help="目标 Python 版本")
    ap.add_argument("--clean", action="store_true", help="先清空旧 wheel")
    args = ap.parse_args()

    if args.clean and os.path.isdir(WHEEL_ROOT):
        shutil.rmtree(WHEEL_ROOT)
    os.makedirs(WHEEL_ROOT, exist_ok=True)

    # 纯 Python 包只下一次
    common = os.path.join(WHEEL_ROOT, "common")
    os.makedirs(common, exist_ok=True)
    if not os.listdir(common):
        print("== 纯 Python 包 ==")
        if not download(PURE, common, "x86_64", "3.11", abi=False):
            print("! 纯 Python 包下载失败")
            return 1
    else:
        print("== 纯 Python 包已存在，跳过 ==")

    failed = []
    for arch in args.arch:
        for pyver in args.py:
            extra = binary_pkgs(pyver)
            if extra is None:
                print("== 跳过 py%s（没有固定版本策略）==" % pyver)
                continue
            dest = os.path.join(WHEEL_ROOT, "py%s-%s" % (pyver.replace(".", ""), arch))
            os.makedirs(dest, exist_ok=True)
            if os.listdir(dest):
                print("== py%s-%s 已存在，跳过 ==" % (pyver, arch))
                continue
            print("== py%s-%s ==" % (pyver, arch))
            if not download(extra, dest, arch, pyver):
                failed.append("%s-%s" % (pyver, arch))

    print("\n== 统计 ==")
    for name in sorted(os.listdir(WHEEL_ROOT)):
        p = os.path.join(WHEEL_ROOT, name)
        if os.path.isdir(p):
            n = len([f for f in os.listdir(p) if f.endswith(".whl")])
            print("  %-16s %2d 个 wheel  %6.1f MB" % (name, n, dir_size_mb(p)))
    print("  合计 %.1f MB" % dir_size_mb(WHEEL_ROOT))
    if failed:
        print("\n! 下载失败：%s" % "、".join(failed))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
