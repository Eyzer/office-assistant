"""列出文件与文件夹清单（导出 CSV，含大小与时间）。"""
import csv
import datetime
import os

from utils.file_utils import human_size

META = {
    "id": "file_list",
    "cat": "文件整理",
    "name": "列出文件清单",
    "desc": "扫描目录，把文件名、路径、大小、修改时间、扩展名导成 CSV，可当作台账用。",
    "fields": [
        {"key": "include_dirs", "label": "包含文件夹", "type": "checkbox", "default": True},
        {"key": "recursive", "label": "含子目录", "type": "checkbox", "default": True},
    ],
}


def _fmt(ts):
    try:
        return datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
    except (OSError, OverflowError):
        return ""


def run_task(inputs, output_dir, params, ctx):
    include_dirs = bool(params.get("include_dirs", True))
    recursive = bool(params.get("recursive", True))
    os.makedirs(output_dir, exist_ok=True)
    roots = [os.path.abspath(p) for p in inputs if p]
    if not roots:
        raise ValueError("没有输入：请先添加文件或文件夹再运行")

    rows = []
    missing = []
    for root in roots:
        if os.path.isfile(root):
            try:
                rows.append((root, os.path.basename(root), os.path.getsize(root),
                             os.path.getmtime(root), "文件"))
            except OSError as exc:
                missing.append("%s（%s）" % (root, exc))
        elif os.path.isdir(root):
            if recursive:
                for base, dirs, files in os.walk(root):
                    if include_dirs:
                        for d in dirs:
                            fp = os.path.join(base, d)
                            try:
                                rows.append((fp, d, 0, os.path.getmtime(fp), "文件夹"))
                            except OSError:
                                pass
                    for f in files:
                        fp = os.path.join(base, f)
                        try:
                            rows.append((fp, f, os.path.getsize(fp), os.path.getmtime(fp), "文件"))
                        except OSError:
                            pass
            else:
                try:
                    names = sorted(os.listdir(root))
                except OSError as exc:
                    missing.append("%s（%s）" % (root, exc))
                    continue
                for n in names:
                    fp = os.path.join(root, n)
                    try:
                        is_dir = os.path.isdir(fp)
                        rows.append((fp, n, 0 if is_dir else os.path.getsize(fp),
                                     os.path.getmtime(fp), "文件夹" if is_dir else "文件"))
                    except OSError:
                        pass
        else:
            missing.append(root)

    if missing:
        ctx.warn("跳过 %d 个无法访问的路径：%s" % (len(missing), "；".join(missing[:5])))
    if not rows:
        hint = ""
        if missing:
            hint = "输入的 %d 个路径都不存在或没有读取权限，例如 %s。" % (len(missing), missing[0])
        raise ValueError("没有可列出的内容。" + hint +
                         "请用「添加文件夹」选择要扫描的目录，或确认路径是否正确。")

    dst = os.path.join(output_dir, "文件清单.csv")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "文件清单_%d.csv" % i)
        i += 1
    total = 0
    with open(dst, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["序号", "名称", "类型", "扩展名", "大小(字节)", "可读大小", "修改时间", "所在目录", "完整路径"])
        for idx, (fp, name, size, mtime, kind) in enumerate(rows, 1):
            total += size
            w.writerow([idx, name, kind, os.path.splitext(name)[1].lower(), size,
                        human_size(size) if size else "", _fmt(mtime),
                        os.path.dirname(fp), fp])
    ctx.info("扫描 %d 个输入路径，共 %d 项（文件 %d / 文件夹 %d），合计 %s"
             % (len(roots), len(rows),
                sum(1 for r in rows if r[4] == "文件"),
                sum(1 for r in rows if r[4] == "文件夹"),
                human_size(total)))
    return {"success_count": 1, "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": [dst]}
