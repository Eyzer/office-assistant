"""任务调度：发现工具、后台执行、进度回传、取消、失败隔离。"""
import importlib
import inspect
import os
import threading
import traceback
import uuid
from datetime import datetime

from utils.deps import MissingDep
from utils.log_utils import TaskLog
from utils.path_utils import ensure_dir

from service import catalog

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS_DIR = os.path.join(BASE, "tools")
TOOL_GROUPS = ("common", "excel", "word")

_registry = None
_lock = threading.Lock()
_tasks = {}


class Ctx:
    """传给工具的运行时上下文。"""

    def __init__(self, task, cancel_event):
        self.task = task
        self._cancel = cancel_event

    def progress(self, current, total, name=""):
        self.task["progress"] = {
            "current": current, "total": total, "name": name,
        }

    def log(self, level, msg):
        self.task["log"]._write(level, msg)

    def info(self, m):
        self.log("信息", m)

    def warn(self, m):
        self.log("警告", m)

    def error(self, m):
        self.log("错误", m)

    def ok(self, m):
        self.log("成功", m)

    @property
    def cancelled(self):
        return self._cancel.is_set()

    def check_cancel(self):
        if self._cancel.is_set():
            raise CancelledError()


class CancelledError(Exception):
    pass


def discover(force=False):
    """扫描 tools/<group>/*.py，加载 META 与 run_task。"""
    global _registry
    with _lock:
        if _registry is not None and not force:
            return _registry
        reg = {}
        for group in TOOL_GROUPS:
            gdir = os.path.join(TOOLS_DIR, group)
            if not os.path.isdir(gdir):
                continue
            for fn in sorted(os.listdir(gdir)):
                if not fn.endswith(".py") or fn.startswith("_"):
                    continue
                mod_name = f"tools.{group}.{fn[:-3]}"
                try:
                    mod = importlib.import_module(mod_name)
                except Exception:
                    continue
                meta = getattr(mod, "META", None)
                run = getattr(mod, "run_task", None)
                if not meta or not callable(run):
                    continue
                key = meta.get("id") or fn[:-3]
                reg[key] = {
                    "id": key,
                    "group": group,
                    "name": meta.get("name", key),
                    "desc": meta.get("desc", ""),
                    "fields": meta.get("fields", []),
                    "cat": meta.get("cat") or catalog.cat_of(key, group),
                    "run": run,
                }
        _registry = reg
        return reg


def tool_list():
    reg = discover()
    items = [
        {k: v for k, v in t.items() if k != "run"}
        for t in reg.values()
    ]
    return sorted(items, key=catalog.sort_key)


def _call_run(fn, inputs, output_dir, params, ctx):
    try:
        sig = inspect.signature(fn)
        wants_ctx = len(sig.parameters) >= 4
    except (TypeError, ValueError):
        wants_ctx = False
    if wants_ctx:
        return fn(inputs, output_dir, params, ctx)
    return fn(inputs, output_dir, params)


def start(tool_id, inputs, output_dir, params=None):
    reg = discover()
    if tool_id not in reg:
        raise ValueError(f"未知工具: {tool_id}")
    task_id = uuid.uuid4().hex[:12]
    cancel_event = threading.Event()
    task = {
        "id": task_id,
        "tool": tool_id,
        "tool_name": reg[tool_id]["name"],
        "status": "排队",
        "progress": {"current": 0, "total": 0, "name": ""},
        "result": None,
        "started": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "log": TaskLog(task_id, os.path.join(BASE, "logs")),
        "cancelled": False,
        "cancel_event": cancel_event,
    }
    _tasks[task_id] = task
    t = threading.Thread(
        target=_run, args=(task, reg[tool_id]["run"], inputs, output_dir, params or {}, cancel_event),
        daemon=True,
    )
    t.start()
    return task_id


def _run(task, fn, inputs, output_dir, params, cancel_event):
    ctx = Ctx(task, cancel_event)
    task["status"] = "运行中"
    ensure_dir(output_dir)
    dry = bool(params.pop("__dry_run__", False))
    try:
        res = _call_run(fn, inputs, output_dir, params, ctx)
        if dry:
            res = _discard_outputs(res, ctx)
        task["result"] = res
        task["status"] = "已取消" if cancel_event.is_set() else "完成"
        ctx.progress(task["progress"]["current"], task["progress"]["total"], "")
        task["log"].ok(f"任务结束：成功 {res.get('success_count', 0)}，失败 {res.get('fail_count', 0)}")
    except CancelledError:
        task["status"] = "已取消"
        task["log"].warn("用户取消")
    except Exception as e:
        task["status"] = "失败"
        task["log"].error(f"任务异常：{e}")
        # 输入路径不对、缺依赖这类"用户能自己解决"的错误不打堆栈，免得吓人
        if not isinstance(e, (ValueError, OSError, MissingDep)):
            task["log"].error(traceback.format_exc(limit=5))
    finally:
        task["finished"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _discard_outputs(res, ctx):
    """预览模式：跑完即删产物，只留统计。"""
    import shutil
    removed = 0
    for p in res.get("output_files", []):
        try:
            if os.path.isdir(p):
                shutil.rmtree(p, ignore_errors=True)
            elif os.path.exists(p):
                os.remove(p)
            removed += 1
        except OSError:
            pass
    ctx.warn(f"预览模式：已清理 {removed} 个产物，源文件未动")
    res["output_files"] = []
    res["dry_run"] = True
    return res


def cancel(task_id):
    t = _tasks.get(task_id)
    if not t:
        return False
    if t["status"] in ("完成", "失败", "已取消"):
        return False
    t["cancelled"] = True
    ev = t.get("cancel_event")
    if ev is not None:
        ev.set()
    t["status"] = "取消中"
    return True


def get(task_id):
    t = _tasks.get(task_id)
    if not t:
        return None
    return {
        "id": t["id"],
        "tool": t["tool"],
        "tool_name": t["tool_name"],
        "status": t["status"],
        "progress": t["progress"],
        "result": t["result"],
        "started": t.get("started"),
        "finished": t.get("finished"),
    }


def logs(task_id, since=0):
    t = _tasks.get(task_id)
    if not t:
        return []
    lines = t["log"].lines
    return lines[since:]


def cleanup(max_keep=50):
    if len(_tasks) <= max_keep:
        return
    done = [k for k, v in _tasks.items() if v["status"] in ("完成", "失败", "已取消")]
    for k in done[: len(_tasks) - max_keep]:
        _tasks.pop(k, None)
