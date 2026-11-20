"""本地 HTTP 服务：静态页面 + JSON API。只绑 127.0.0.1，token 鉴权。"""
import atexit
import base64
import json
import os
import posixpath
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import version as app_info
from service import task_service
from utils import config_utils, sjz

BASE = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE, "web")
UPLOAD_DIR = os.path.join(BASE, "uploads")
TOKEN = ""
CHUNK = 256 * 1024
API_VERSION = app_info.API_VERSION

# 后台服务版本与页面版本不一致时的提示语
STALE_HINT = "接口不存在：当前运行的后台服务可能是旧版本，请关闭程序重新启动"

MIME = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
}


def _safe_seg(s):
    """清洗单个路径片段：去掉分隔符与控制字符，保留中文。"""
    import re
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", str(s)).strip().strip(".")
    return s[:120] or "x"


def _safe_rel(name):
    """把浏览器给的相对路径变成安全的相对路径，禁止上跳。"""
    name = urllib.parse.unquote(str(name or "")).replace("\\", "/")
    parts = [p for p in name.split("/") if p not in ("", ".", "..")]
    if not parts:
        return ""
    return os.path.join(*[_safe_seg(p) for p in parts])


def cleanup_uploads(days=3):
    """清掉几天前的拖拽暂存目录，避免越积越大。"""
    import shutil
    import time
    if not os.path.isdir(UPLOAD_DIR):
        return 0
    cutoff = time.time() - days * 86400
    removed = 0
    for n in os.listdir(UPLOAD_DIR):
        fp = os.path.join(UPLOAD_DIR, n)
        try:
            if os.path.isdir(fp) and os.path.getmtime(fp) < cutoff:
                shutil.rmtree(fp, ignore_errors=True)
                removed += 1
        except OSError:
            pass
    return removed


def resolve_open_target(path, create=True):
    """把传入路径解析成"要打开的目录"；不合法抛 ValueError。

    create=True 时目录不存在会先建出来（点「打开输出目录」时还没跑过任务）。
    """
    if not path:
        raise ValueError("路径为空")
    p = os.path.abspath(path)
    if os.path.isfile(p):
        return os.path.dirname(p)
    if os.path.isdir(p):
        return p
    if not create:
        raise ValueError("目录不存在：%s" % p)
    try:
        os.makedirs(p, exist_ok=True)
    except OSError as exc:
        raise ValueError("目录不存在且无法创建：%s" % exc)
    return p


def open_path(path, create=True):
    """用系统文件管理器打开目录；传文件路径则定位到它所在的目录。"""
    try:
        target = resolve_open_target(path, create)
    except ValueError as exc:
        return {"ok": False, "error": str(exc)}

    try:
        if os.name == "nt":
            os.startfile(target)  # noqa: 仅 Windows 存在
            return {"ok": True, "path": target, "opener": "explorer"}
        if sys.platform == "darwin":
            subprocess.Popen(["open", target])
            return {"ok": True, "path": target, "opener": "open"}
        for name, cmd in (("xdg-open", ["xdg-open", target]),
                          ("gio", ["gio", "open", target]),
                          ("dde-file-manager", ["dde-file-manager", target]),
                          ("peony", ["peony", target]),
                          ("nautilus", ["nautilus", target])):
            if shutil.which(name):
                subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return {"ok": True, "path": target, "opener": name}
    except Exception as exc:
        return {"ok": False, "error": "打开失败：%s" % exc}
    return {"ok": False, "reason": "no-opener",
            "error": "系统里没有可用的文件管理器（xdg-open / gio / peony）"}


def _use_win_pick():
    """是否走 Windows 原生对话框（桌面端测试可覆盖此函数）。"""
    return os.name == "nt"


def native_pick(mode="files", start="", title=""):
    """调用系统原生的文件/目录选择框（zenity → kdialog）。

    返回 dict：{"ok":True,"paths":[...],"dialog":"zenity"}
      ok=False 时带 reason：no-dialog（没装）/ no-display（无图形会话）/ error
    取消选择返回 ok=True, paths=[], cancelled=True。
    """
    if mode == "dir":
        title = title or "选择文件夹"
    else:
        title = title or "选择文件"
    start = os.path.expanduser(start or "") or os.path.expanduser("~")

    # Windows：走系统 WinForms 对话框（Windows 上没有 zenity/kdialog）
    if _use_win_pick():
        return _pick_win(mode, start, title)

    if shutil.which("zenity"):
        cmd = ["zenity", "--file-selection", "--title", title]
        if mode == "dir":
            cmd.append("--directory")
        else:
            cmd += ["--multiple", "--separator", "\n"]
        if os.path.isdir(start):
            cmd += ["--filename", start + os.sep]
        return _run_pick(cmd, "zenity")

    if shutil.which("kdialog"):
        if mode == "dir":
            cmd = ["kdialog", "--title", title, "--getexistingdirectory", start]
        else:
            cmd = ["kdialog", "--title", title, "--getopenfilename", start, "*", "--multiple"]
        return _run_pick(cmd, "kdialog")

    if shutil.which("yad"):
        cmd = ["yad", "--file", "--title", title]
        if mode == "dir":
            cmd.append("--directory")
        else:
            cmd += ["--multiple", "--separator", "\n"]
        return _run_pick(cmd, "yad")

    return {"ok": False, "reason": "no-dialog",
            "error": "系统里没有 zenity / kdialog，无法弹出原生对话框"}


_PICK_PS = r"""
try {
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $enc  = New-Object System.Text.UTF8Encoding($false)
  $out  = $env:OA_PICK_OUT
  $mode = $env:OA_PICK_MODE
  $start= $env:OA_PICK_START
  $title= $env:OA_PICK_TITLE

  # 透明置顶宿主窗体，保证对话框弹在最前面
  $owner = New-Object System.Windows.Forms.Form
  $owner.TopMost = $true
  $owner.ShowInTaskbar = $false
  $owner.Opacity = 0
  $owner.Size = New-Object System.Drawing.Size(1,1)
  $owner.StartPosition = 'CenterScreen'
  $owner.Show()

  if ($mode -eq 'dir') {
    $d = New-Object System.Windows.Forms.FolderBrowserDialog
    $d.ShowNewFolderButton = $true
    if ($title) { $d.Description = $title }
    if ($start) { $d.SelectedPath = $start }
    $ok = ($d.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK)
    $val = $d.SelectedPath
  } else {
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Multiselect = $true
    $d.CheckFileExists = $true
    if ($title) { $d.Title = $title }
    if ($start) { $d.InitialDirectory = $start }
    $d.Filter = '办公文档|*.xlsx;*.xlsm;*.xls;*.csv;*.docx;*.doc;*.pptx;*.ppt;*.pdf;*.txt;*.json;*.md|图片|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|所有文件|*.*'
    $ok = ($d.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK)
    $val = ($d.FileNames -join "`n")
  }
  if ($ok) {
    [IO.File]::WriteAllText($out, $val, $enc)
    exit 0
  }
  [IO.File]::WriteAllText($out, '', $enc)
  exit 1
} catch {
  try {
    [IO.File]::WriteAllText($env:OA_PICK_OUT, '__ERR__' + $_.Exception.Message, $enc)
  } catch {}
  exit 2
}
"""


def _pick_win_once(mode, start, title):
    """一次性进程版（兜底）：每次启动一个 PowerShell，约 1.3 秒。"""
    fd, out = tempfile.mkstemp(prefix="oa_pick_", suffix=".txt")
    os.close(fd)
    env = dict(os.environ)
    env.update({"OA_PICK_MODE": mode, "OA_PICK_OUT": out,
                "OA_PICK_START": start or "", "OA_PICK_TITLE": title or ""})
    try:
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-STA",
             "-ExecutionPolicy", "Bypass", "-Command", _PICK_PS],
            capture_output=True, env=env, timeout=1800)
    except subprocess.TimeoutExpired:
        _rm_quiet(out)
        return {"ok": False, "reason": "error", "error": "对话框超时"}
    except OSError as exc:
        _rm_quiet(out)
        return {"ok": False, "reason": "error", "error": "对话框启动失败：%s" % exc}

    txt = ""
    try:
        with open(out, "r", encoding="utf-8", errors="replace") as f:
            txt = f.read()
    except OSError:
        pass
    finally:
        _rm_quiet(out)

    if txt.startswith("__ERR__"):
        return {"ok": False, "reason": "no-display" if "session" in txt.lower() else "error",
                "error": txt[7:].strip()[:300], "dialog": "windows"}

    paths = [s.strip() for s in txt.splitlines() if s.strip()]
    if p.returncode == 0:
        return {"ok": True, "paths": paths, "dialog": "windows"}
    if p.returncode == 1:
        return {"ok": True, "paths": [], "cancelled": True, "dialog": "windows"}
    err = (p.stderr or b"").decode("utf-8", "replace").strip()
    return {"ok": False, "reason": "error",
            "error": err[-300:] or ("对话框退出码 %d" % p.returncode),
            "dialog": "windows"}


_PICK_HOST_PS = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null
$enc = New-Object System.Text.UTF8Encoding($false)
$o = [Console]::Out
$i = [Console]::In

# 透明置顶宿主窗体：保证对话框弹在最前面，且不在任务栏留一个空窗口
$owner = New-Object System.Windows.Forms.Form
$owner.TopMost = $true
$owner.ShowInTaskbar = $false
$owner.Opacity = 0
$owner.Size = New-Object System.Drawing.Size(1,1)
$owner.StartPosition = 'CenterScreen'
$owner.Show()

$o.WriteLine('READY')
$o.Flush()

while ($true) {
  $line = $i.ReadLine()
  if ($null -eq $line) { break }
  $line = $line.Trim()
  if ($line -eq 'QUIT') { break }
  if ($line -eq 'PING') { $o.WriteLine('PONG'); $o.Flush(); continue }
  if (-not $line.StartsWith('PICK ')) { $o.WriteLine('DONE 2'); $o.Flush(); continue }

  $rc = 2
  $req = $null
  try {
    $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($line.Substring(5)))
    $req = $json | ConvertFrom-Json
    $mode  = $req.mode
    $start = $req.start
    $title = $req.title
    if ($mode -eq 'dir') {
      $d = New-Object System.Windows.Forms.FolderBrowserDialog
      $d.ShowNewFolderButton = $true
      if ($title) { $d.Description = $title }
      if ($start -and (Test-Path -LiteralPath $start)) { $d.SelectedPath = $start }
      $ok = ($d.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK)
      $val = $d.SelectedPath
    } else {
      $d = New-Object System.Windows.Forms.OpenFileDialog
      $d.Multiselect = $true
      $d.CheckFileExists = $true
      if ($title) { $d.Title = $title }
      if ($start -and (Test-Path -LiteralPath $start)) { $d.InitialDirectory = $start }
      $d.Filter = '办公文档|*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv;*.docx;*.doc;*.docm;*.pptx;*.ppt;*.pdf;*.txt;*.json;*.md|图片|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|所有文件|*.*'
      $ok = ($d.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK)
      $val = ($d.FileNames -join "`n")
    }
    $d.Dispose()
    if ($ok) {
      [IO.File]::WriteAllText($req.out, $val, $enc)
      $rc = 0
    } else {
      [IO.File]::WriteAllText($req.out, '', $enc)
      $rc = 1
    }
  } catch {
    $rc = 2
    try {
      if ($req -and $req.out) {
        [IO.File]::WriteAllText($req.out, '__ERR__' + $_.Exception.Message, $enc)
      }
    } catch {}
  }
  $o.WriteLine('DONE ' + $rc)
  $o.Flush()
}
"""


def _no_window():
    """Windows 下别给子进程弹控制台窗口。"""
    if os.name == "nt":
        return {"creationflags": 0x08000000}  # CREATE_NO_WINDOW
    return {}


class _PickerHost:
    """常驻 PowerShell 对话框宿主。

    每点一次都新起 powershell + 加载 WinForms 要约 1.3 秒，用户感觉是"点了没反应"。
    这里把进程与程序集加载只做一次（服务启动时预热），之后每次点击只发一行指令，
    实测响应 0~13 毫秒。
    """

    def __init__(self):
        self._proc = None
        self._q = None
        self._lock = threading.Lock()        # 同一时刻只允许一个对话框
        self._start_lock = threading.Lock()

    def alive(self):
        p = self._proc
        return p is not None and p.poll() is None

    def start(self, timeout=25):
        """启动宿主并等它报到 READY。返回 (ok, err)。"""
        with self._start_lock:
            if self.alive():
                return True, ""
            self._kill()
            try:
                proc = subprocess.Popen(
                    ["powershell", "-NoProfile", "-NonInteractive", "-STA",
                     "-ExecutionPolicy", "Bypass", "-Command", _PICK_HOST_PS],
                    stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                    stderr=subprocess.DEVNULL, **_no_window())
            except OSError as exc:
                return False, "对话框进程启动失败：%s" % exc

            q = queue.Queue()
            self._proc = proc
            self._q = q

            def _reader(p=proc, qq=q):
                try:
                    for line in iter(p.stdout.readline, b""):
                        qq.put(line.decode("utf-8", "replace").strip())
                except Exception:
                    pass
                qq.put(None)

            threading.Thread(target=_reader, daemon=True).start()

        try:
            first = q.get(timeout=timeout)
        except queue.Empty:
            self._kill()
            return False, "对话框进程没有响应（%d 秒超时）" % timeout
        if first == "READY":
            return True, ""
        self._kill()
        return False, "对话框进程启动异常：%s" % (first or "无输出")

    def pick(self, mode, start, title, timeout=1800):
        with self._lock:
            ok, err = self.start()
            if not ok:
                return {"ok": False, "reason": "error", "error": err, "dialog": "windows"}

            fd, out = tempfile.mkstemp(prefix="oa_pick_", suffix=".txt")
            os.close(fd)
            req = {"mode": mode, "start": start or "", "title": title or "", "out": out}
            payload = base64.b64encode(
                json.dumps(req, ensure_ascii=False).encode("utf-8")).decode("ascii")
            try:
                self._proc.stdin.write(("PICK " + payload + "\n").encode("ascii"))
                self._proc.stdin.flush()
            except OSError as exc:
                self._kill()
                _rm_quiet(out)
                return {"ok": False, "reason": "error",
                        "error": "对话框进程已退出：%s" % exc, "dialog": "windows"}

            rc = None
            timed_out = False
            try:
                while True:
                    line = self._q.get(timeout=timeout)
                    if line is None:
                        break
                    if line.startswith("DONE"):
                        try:
                            rc = int(line.split(None, 1)[1])
                        except (IndexError, ValueError):
                            rc = 2
                        break
            except queue.Empty:
                timed_out = True
                self._kill()

            txt = ""
            try:
                with open(out, "r", encoding="utf-8", errors="replace") as f:
                    txt = f.read()
            except OSError:
                pass
            finally:
                _rm_quiet(out)

            if txt.startswith("\ufeff"):
                txt = txt[1:]
            if txt.startswith("__ERR__"):
                return {"ok": False,
                        "reason": "no-display" if "session" in txt.lower() else "error",
                        "error": txt[7:].strip()[:300], "dialog": "windows"}
            if timed_out:
                return {"ok": False, "reason": "error", "error": "对话框超时",
                        "dialog": "windows"}
            if rc is None:
                return {"ok": False, "reason": "error",
                        "error": "对话框进程意外退出，请重试", "dialog": "windows"}
            if rc == 0:
                paths = [s.strip() for s in txt.splitlines() if s.strip()]
                return {"ok": True, "paths": paths, "dialog": "windows"}
            if rc == 1:
                return {"ok": True, "paths": [], "cancelled": True, "dialog": "windows"}
            return {"ok": False, "reason": "error",
                    "error": "对话框返回异常（代码 %d）" % rc, "dialog": "windows"}

    def _kill(self):
        p, self._proc, self._q = self._proc, None, None
        if p is None:
            return
        try:
            if p.poll() is None:
                try:
                    p.stdin.write(b"QUIT\n")
                    p.stdin.flush()
                except OSError:
                    pass
                try:
                    p.wait(timeout=1)
                except subprocess.TimeoutExpired:
                    p.kill()
        except Exception:
            pass

    def shutdown(self):
        self._kill()


_PICKER = _PickerHost()
atexit.register(_PICKER.shutdown)


def prewarm_picker():
    """服务启动时先把对话框宿主拉起来，用户点击时就不用再等 1.3 秒。"""
    if not _use_win_pick():
        return True
    ok, err = _PICKER.start()
    if not ok:
        # 预热失败不致命：真点击时还会再试一次，再不行就走一次性进程
        print("文件对话框预热失败：%s" % err)
    return ok


def _pick_win(mode, start, title):
    """优先走常驻宿主；宿主起不来就退回一次性进程。"""
    r = _PICKER.pick(mode, start, title)
    if r.get("ok") or r.get("reason") != "error":
        return r
    if not _PICKER.alive():
        twice = _pick_win_once(mode, start, title)
        if twice.get("ok") or twice.get("reason") != "error":
            return twice
    return r


def _rm_quiet(fp):
    try:
        os.remove(fp)
    except OSError:
        pass


def _run_pick(cmd, dialog):
    """跑一次对话框命令，把 stdout 拆成路径数组。"""
    env = dict(os.environ)
    if not env.get("DISPLAY") and not env.get("WAYLAND_DISPLAY"):
        return {"ok": False, "reason": "no-display",
                "error": "当前没有图形会话（DISPLAY 为空），无法弹出系统对话框"}
    try:
        p = subprocess.run(cmd, capture_output=True, env=env, timeout=1800)
    except subprocess.TimeoutExpired:
        return {"ok": False, "reason": "error", "error": "对话框超时"}
    except OSError as exc:
        return {"ok": False, "reason": "error", "error": "对话框启动失败：%s" % exc}

    out = (p.stdout or b"").decode("utf-8", "replace")
    paths = [s.strip() for s in out.splitlines() if s.strip()]
    if p.returncode == 0:
        return {"ok": True, "paths": paths, "dialog": dialog}
    if p.returncode in (1, 255):  # zenity/kdialog 取消
        return {"ok": True, "paths": [], "cancelled": True, "dialog": dialog}
    err = (p.stderr or b"").decode("utf-8", "replace").strip()
    return {"ok": False, "reason": "error",
            "error": err or ("对话框退出码 %d" % p.returncode),
            "dialog": dialog}


class Handler(BaseHTTPRequestHandler):
    server_version = "DocBatch/" + app_info.VERSION
    # 开持久连接：批量上传几十个文件时省掉每个请求的 TCP 握手
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass  # 不污染控制台

    # ---------- 基础 ----------
    def _send(self, code, body=b"", ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if body:
            self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def _err(self, msg, code=400):
        self._json({"ok": False, "error": msg}, code)

    def _token_ok(self, query):
        t = query.get("t", [""])[0] or self.headers.get("X-Token", "")
        return t == TOKEN

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return {}

    # ---------- 拖拽上传 ----------
    def _upload(self, query):
        """接收原始字节，落到 uploads/<会话>/<相对路径>，返回绝对路径。

        浏览器出于安全不给真实路径，拖进来的文件只能由本机服务落盘一次。
        """
        raw_name = query.get("name", [""])[0]
        sid = _safe_seg(query.get("sid", [""])[0] or "s")
        rel = _safe_rel(raw_name or "file")
        if not rel:
            return self._err("非法文件名")
        dst = os.path.join(UPLOAD_DIR, sid, rel)
        root = os.path.abspath(UPLOAD_DIR)
        if not os.path.abspath(dst).startswith(root + os.sep):
            return self._err("非法路径", 403)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        left = int(self.headers.get("Content-Length") or 0)
        written = 0
        try:
            with open(dst, "wb") as f:
                while left > 0:
                    chunk = self.rfile.read(min(CHUNK, left))
                    if not chunk:
                        break
                    f.write(chunk)
                    written += len(chunk)
                    left -= len(chunk)
        except OSError as exc:
            return self._err("写入失败：%s" % exc, 500)
        # sjz 密文自动解密：检测到就原地换成明文，失败不影响上传本身
        info = {"detected": False, "decrypted": False}
        try:
            hit, via = sjz.detect(dst)
            info["detected"] = hit
            if hit:
                info["via"] = via
                r = sjz.decrypt_file(dst)
                info["decrypted"] = r["changed"]
                if r["changed"]:
                    written = r["size"]
                elif r["reason"]:
                    info["error"] = r["reason"]
        except Exception as exc:  # 解密属于增强功能，绝不能拖垮上传
            info["error"] = str(exc)
        return self._json({"ok": True, "path": os.path.abspath(dst),
                           "size": written, "name": rel, "sjz": info})

    # ---------- 路由 ----------
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path.startswith("/api/"):
            if not self._token_ok(query):
                return self._err("鉴权失败", 403)
            try:
                return self._api_get(path, query)
            except Exception as exc:  # 单点异常不能让连接直接断掉
                return self._err("服务端异常：%s" % exc, 500)

        return self._static(path, query)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed.query)
        if not self._token_ok(query):
            # 请求体没读，持久连接会被残留字节污染，直接断开
            self.close_connection = True
            return self._err("鉴权失败", 403)
        # 拖拽上传走原始字节流，不能按 JSON 读
        if parsed.path == "/api/upload":
            return self._upload(query)
        data = self._body()
        try:
            return self._api_post(parsed.path, data)
        except Exception as e:
            return self._err(str(e), 500)

    def _static(self, path, query):
        if path in ("/", "/index.html"):
            fp = os.path.join(WEB_DIR, "index.html")
            with open(fp, "rb") as f:
                html = f.read().decode("utf-8")
            html = (html.replace("%%TOKEN%%", TOKEN)
                        .replace("%%APP_NAME%%", app_info.APP_NAME)
                        .replace("%%APP_TAGLINE%%", app_info.APP_TAGLINE)
                        .replace("%%VERSION%%", app_info.VERSION)
                        .replace("%%API_VER%%", str(API_VERSION)))
            return self._send(200, html.encode("utf-8"), MIME[".html"])

        rel = posixpath.normpath(path).lstrip("/")
        fp = os.path.join(WEB_DIR, rel)
        if not os.path.abspath(fp).startswith(os.path.abspath(WEB_DIR)):
            return self._err("非法路径", 403)
        if not os.path.isfile(fp):
            return self._err("未找到", 404)
        ext = os.path.splitext(fp)[1].lower()
        with open(fp, "rb") as f:
            self._send(200, f.read(), MIME.get(ext, "application/octet-stream"))

    def _api_get(self, path, query):
        if path == "/api/tools":
            return self._json({"ok": True, "tools": task_service.tool_list()})
        if path == "/api/task":
            tid = query.get("id", [""])[0]
            t = task_service.get(tid)
            if not t:
                return self._err("任务不存在", 404)
            since = int(query.get("since", ["0"])[0] or 0)
            t["logs"] = task_service.logs(tid, since)
            t["log_count"] = since + len(t["logs"])
            return self._json({"ok": True, "task": t})
        if path == "/api/profiles":
            return self._json({"ok": True, "profiles": config_utils.list_profiles()})
        if path == "/api/profile":
            name = query.get("name", [""])[0]
            try:
                return self._json({"ok": True, "data": config_utils.load_profile(name)})
            except OSError:
                return self._err("方案不存在", 404)
        if path == "/api/fs":
            p = query.get("path", [""])[0] or os.path.expanduser("~")
            return self._json(self._list_dir(p))
        if path == "/api/config":
            return self._json({
                "ok": True,
                "output_dir": config_utils.get("general", "output_dir", "output"),
                "workers": config_utils.get("run", "workers", "2"),
            })
        if path == "/api/deps":
            from utils import deps
            return self._json({"ok": True, "deps": deps.report(),
                               "upload_dir": UPLOAD_DIR,
                               "api_version": API_VERSION,
                               "version": app_info.VERSION})
        if path == "/api/ping":
            return self._json({"ok": True, "version": app_info.VERSION,
                               "api_version": API_VERSION,
                               "name": app_info.APP_NAME})
        if path == "/api/sjz":
            return self._json({"ok": True, "sjz": sjz.self_check()})
        if path == "/api/about":
            return self._json({"ok": True, "about": app_info.about()})
        if path == "/api/pick":
            mode = (query.get("mode", ["files"])[0] or "files").lower()
            if mode == "warm":
                # 页面加载时预热对话框进程，等用户点按钮时就已经热好了
                return self._json({"ok": prewarm_picker()})
            if mode not in ("files", "dir"):
                return self._err("mode 只能是 files 或 dir")
            start = query.get("start", [""])[0]
            r = native_pick(mode, start)
            r["ok"] = r.get("ok", False)
            return self._json(r)
        return self._err(STALE_HINT, 404)

    def _api_post(self, path, data):
        if path == "/api/run":
            tool = data.get("tool", "")
            inputs = data.get("inputs") or []
            out = data.get("output_dir") or config_utils.get("general", "output_dir", "output")
            params = data.get("params") or {}
            if not tool:
                return self._err("缺少工具")
            if not inputs:
                return self._err("请先选择文件或文件夹")
            tid = task_service.start(tool, inputs, out, params)
            config_utils.set("general", "last_tool", tool)
            return self._json({"ok": True, "task_id": tid})
        if path == "/api/open":
            p = (data.get("path") or "").strip()
            if not p:
                return self._err("缺少路径")
            if not os.path.isabs(p):
                p = os.path.join(BASE, p)
            r = open_path(p, create=bool(data.get("create", True)))
            r["ok"] = r.get("ok", False)
            return self._json(r)
        if path == "/api/cancel":
            tid = data.get("id", "")
            ok = task_service.cancel(tid)
            return self._json({"ok": ok})
        if path == "/api/profile/save":
            name = (data.get("name") or "").strip()
            if not name:
                return self._err("请填方案名")
            config_utils.save_profile(name, data.get("data") or {})
            return self._json({"ok": True})
        if path == "/api/profile/delete":
            name = (data.get("name") or "").strip()
            config_utils.delete_profile(name)
            return self._json({"ok": True})
        if path == "/api/config":
            if "output_dir" in data:
                config_utils.set("general", "output_dir", data["output_dir"])
            if "workers" in data:
                config_utils.set("run", "workers", data["workers"])
            return self._json({"ok": True})
        return self._err(STALE_HINT, 404)

    def _list_dir(self, path):
        path = os.path.abspath(os.path.expanduser(path))
        if not os.path.isdir(path):
            return {"ok": False, "error": "目录不存在"}
        dirs, files = [], []
        try:
            for n in sorted(os.listdir(path)):
                if n.startswith("."):
                    continue
                fp = os.path.join(path, n)
                if os.path.isdir(fp):
                    dirs.append(n)
                else:
                    files.append(n)
        except PermissionError:
            return {"ok": False, "error": "无权限"}
        parent = os.path.dirname(path)
        return {"ok": True, "path": path, "parent": parent, "dirs": dirs, "files": files}


def serve(host="127.0.0.1", port=0, token="", open_browser=True):
    global TOKEN
    TOKEN = token
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    dropped = cleanup_uploads(3)
    if dropped:
        print("已清理 %d 个过期暂存目录" % dropped)
    # 后台预热系统文件对话框，避免用户第一次点「添加文件」时等 1 秒多
    # 服务器 / SSH 会话没有图形环境，预热只会白跑一次，直接跳过
    if (not sys.platform.startswith("linux")) or os.environ.get("DISPLAY") \
            or os.environ.get("WAYLAND_DISPLAY"):
        threading.Thread(target=prewarm_picker, daemon=True).start()
    httpd = ThreadingHTTPServer((host, port), Handler)
    actual_port = httpd.server_address[1]
    url = f"http://{host}:{actual_port}/?t={token}"
    if open_browser:
        import webbrowser
        timer = threading.Timer(0.6, lambda: webbrowser.open(url))
        timer.daemon = True
        timer.start()
    print("%s v%s 已启动: %s" % (app_info.APP_NAME, app_info.VERSION, url))
    print("作者：%s <%s>" % (app_info.AUTHOR, app_info.EMAIL))
    print("按 Ctrl+C 退出")

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return url
