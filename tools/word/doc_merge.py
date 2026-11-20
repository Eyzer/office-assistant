"""合并多个文档为一个，可选按文件夹分组分别合并。"""
import os
from collections import defaultdict

import docx

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path, safe_name

META = {
    "id": "doc_merge",
    "name": "批量合并文档",
    "desc": "按选择顺序合并多个文档；分组模式按所在文件夹分别合并。",
    "fields": [
        {"key": "group_by_dir", "label": "按文件夹分组", "type": "checkbox", "default": False,
         "help": "勾选后每个文件夹各合并成一份"},
        {"key": "page_break", "label": "文档间插入分页", "type": "checkbox", "default": True},
        {"key": "sort_by_name", "label": "按文件名排序", "type": "checkbox", "default": True},
    ],
}


def _merge_group(files, out_file, page_break, ctx):
    base = docx.Document(files[0])
    for idx, p in enumerate(files[1:], 1):
        du.append_document(base, p, page_break=page_break)
        ctx.info(f"追加 {os.path.basename(p)}")
    du.save(base, out_file)
    return out_file


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    if not files:
        return {"success_count": 0, "fail_count": 0,
                "log_messages": ["未找到 docx 文件"], "output_files": []}

    group_by_dir = bool(params.get("group_by_dir"))
    page_break = bool(params.get("page_break", True))
    sort_by_name = bool(params.get("sort_by_name", True))

    if sort_by_name:
        files.sort(key=lambda p: os.path.basename(p))

    ok, fail, outs = 0, 0, []

    if group_by_dir:
        groups = defaultdict(list)
        for f in files:
            groups[os.path.dirname(f)].append(f)
        total = len(groups)
        for i, (folder, group) in enumerate(groups.items(), 1):
            if ctx.cancelled:
                ctx.warn("已取消")
                break
            try:
                tag = safe_name(os.path.basename(folder) or "根目录")
                dst = out_path(os.path.join(folder, "x.docx"), output_dir,
                               suffix=f"_{tag}_合并", ext=".docx")
                _merge_group(group, dst, page_break, ctx)
                outs.append(dst)
                ok += 1
                ctx.info(f"{tag}: 合并 {len(group)} 个文档")
            except Exception as e:
                fail += 1
                ctx.error(f"{folder}: {e}")
            ctx.progress(i, total, os.path.basename(folder))
    else:
        try:
            dst = out_path(files[0], output_dir, suffix="_合并", ext=".docx")
            _merge_group(files, dst, page_break, ctx)
            outs.append(dst)
            ok = 1
            ctx.progress(1, 1, "合并完成")
        except Exception as e:
            fail = 1
            ctx.error(str(e))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
