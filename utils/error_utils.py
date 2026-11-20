"""错误隔离：单文件失败不中断批量。"""
import functools
import traceback


def guard(log, label=""):
    """装饰器：捕获异常写日志，返回失败标记，不抛出。"""
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*a, **kw):
            try:
                return True, fn(*a, **kw)
            except Exception as e:
                msg = f"{label or fn.__name__} 失败: {e}"
                if log:
                    log.error(msg)
                    log.error(traceback.format_exc(limit=3))
                return False, None
        return wrapper
    return deco


def run_isolated(fn, *a, log=None, label="", **kw):
    """直接调用版的隔离执行。返回 (ok, result)。"""
    try:
        return True, fn(*a, **kw)
    except Exception as e:
        if log:
            log.error(f"{label or getattr(fn, '__name__', '任务')} 失败: {e}")
            log.error(traceback.format_exc(limit=3))
        return False, None
