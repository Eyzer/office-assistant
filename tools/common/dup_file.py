"""找出内容相同或文件名相同的重复文件，输出清单（不删源文件）。"""
import hashlib
import os

from utils import batch
from utils.file_utils import collect_files, human_size

META = {
    "id": "dup_file",
    "cat": "文件整理",
    "name": "查找重复文件",
    "desc": "按内容指纹或文件名找出重复文件，生成清单；可将重复项复制到单独目录，不动源文件。",
    "fields": [
        {"key": "by", "label": "判定方式", "type": "select", "default": "hash",
         "options": [{"value": "hash", "label": "内容相同"},
                     {"value": "name", "label": "文件名相同"},
                     {"value": "size", "label": "大小相同"}]},
        {"key": "copy", "label": "复制重复项到输出目录", "type": "checkbox", "default": False},
    ],
}


def _hash(path, chunk=1 << 20):
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def run_task(inputs, output_dir, params, ctx):
    mode = params.get("by") or "hash"
    do_copy = bool(params.get("copy"))
    files = collect_files(inputs)
    if not files:
        ctx.warn("没有找到文件")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    groups = {}
    total = len(files)
    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            if mode == "hash":
                key = _hash(f)
            elif mode == "size":
                key = str(os.path.getsize(f))
            else:
                key = os.path.basename(f).lower()
            groups.setdefault(key, []).append(f)
        except OSError as exc:
            ctx.warn("%s: %s" % (os.path.basename(f), exc))
        ctx.progress(i, total, os.path.basename(f))

    dups = {k: v for k, v in groups.items() if len(v) > 1}
    import csv
    os.makedirs(output_dir, exist_ok=True)
    dst = os.path.join(output_dir, "重复文件清单.csv")
    i = 1
    while os.path.exists(dst):
        dst = os.path.join(output_dir, "重复文件清单_%d.csv" % i)
        i += 1
    waste = 0
    with open(dst, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["组", "序号", "路径", "大小", "修改时间"])
        for gi, (key, paths) in enumerate(sorted(dups.items()), 1):
            try:
                size = os.path.getsize(paths[0])
            except OSError:
                size = 0
            waste += size * (len(paths) - 1)
            for k, p in enumerate(paths, 1):
                mtime = ""
                try:
                    import datetime
                    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d %H:%M")
                except OSError:
                    pass
                w.writerow([gi, k, p, size, mtime])

    outs = [dst]
    if do_copy and dups:
        sub = os.path.join(output_dir, "重复文件")
        os.makedirs(sub, exist_ok=True)
        import shutil
        for gi, (key, paths) in enumerate(sorted(dups.items()), 1):
            for p in paths[1:]:
                target = os.path.join(sub, "组%03d_%s" % (gi, os.path.basename(p)))
                try:
                    shutil.copy2(p, target)
                    outs.append(target)
                except OSError as exc:
                    ctx.warn("复制失败 %s: %s" % (p, exc))

    ctx.info("发现 %d 组重复，可省 %s" % (len(dups), human_size(waste)))
    return {"success_count": len(dups), "fail_count": 0,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
