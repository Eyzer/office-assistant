"""日志：内存收集 + 落文件。任务内用它，不 print。"""
import datetime
import os
import threading


class TaskLog:
    """一个任务的日志收集器。线程安全。"""

    def __init__(self, task_id="", log_dir=None):
        self.task_id = task_id
        self.lines = []
        self._lock = threading.Lock()
        self.log_dir = log_dir
        self._file = None
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            self._file = os.path.join(log_dir, f"task_{stamp}_{task_id}.log")

    def _write(self, level, msg):
        line = f"[{datetime.datetime.now().strftime('%H:%M:%S')}][{level}] {msg}"
        with self._lock:
            self.lines.append(line)
            if self._file:
                try:
                    with open(self._file, "a", encoding="utf-8") as f:
                        f.write(line + "\n")
                except OSError:
                    pass

    def info(self, msg):
        self._write("信息", msg)

    def warn(self, msg):
        self._write("警告", msg)

    def error(self, msg):
        self._write("错误", msg)

    def ok(self, msg):
        self._write("成功", msg)


def tail(path, n=200):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return "".join(f.readlines()[-n:])
    except OSError:
        return ""
