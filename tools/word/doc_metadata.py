"""清空或修改文档元数据（作者、公司、标题等），隐私处理刚需。"""
import os

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import out_path

META = {
    "id": "doc_metadata",
    "name": "文档元数据清理",
    "desc": "清空或统一修改作者、单位、标题等文档属性，去掉隐私信息。",
    "fields": [
        {"key": "clear", "label": "全部清空", "type": "checkbox", "default": True},
        {"key": "author", "label": "作者", "type": "text", "default": ""},
        {"key": "company", "label": "单位/公司", "type": "text", "default": ""},
        {"key": "title", "label": "标题", "type": "text", "default": ""},
    ],
}


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    clear = bool(params.get("clear", True))
    ok, fail, outs = 0, 0, []
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        try:
            d = du.load(f)
            if clear:
                du.clear_core_props(d)
            du.set_core_props(
                d,
                author=params.get("author") or None,
                title=params.get("title") or None,
                last_modified_by=params.get("author") or None,
            )
            dst = out_path(f, output_dir, suffix="_元数据")
            du.save(d, dst)
            outs.append(dst)
            ok += 1
            ctx.info(f"{os.path.basename(f)}: 元数据已处理")
        except Exception as e:
            fail += 1
            ctx.error(f"{f}: {e}")
        ctx.progress(i, total, os.path.basename(f))

    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
