"""服务层接口测试：不启动浏览器，直接起一个本地 HTTP 服务打接口。"""
import io
import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer

import server
import version
from server import Handler


class ServerApiTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._old_token = server.TOKEN
        cls._old_upload = server.UPLOAD_DIR
        cls.tmp = tempfile.mkdtemp(prefix="dbt_api_")
        server.TOKEN = "testtok"
        server.UPLOAD_DIR = os.path.join(cls.tmp, "uploads")
        os.makedirs(server.UPLOAD_DIR, exist_ok=True)
        cls.httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        cls.port = cls.httpd.server_address[1]
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        server.TOKEN = cls._old_token
        server.UPLOAD_DIR = cls._old_upload
        shutil.rmtree(cls.tmp, ignore_errors=True)

    # ---------- 工具 ----------
    def _url(self, path, token="testtok"):
        url = "http://127.0.0.1:%d%s" % (self.port, path)
        if token:
            url += ("&" if "?" in path else "?") + "t=" + token
        return url

    def _open(self, path, token="testtok", data=None):
        op = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        req = urllib.request.Request(self._url(path, token), data=data)
        try:
            r = op.open(req, timeout=15)
            return r.status, r.read()
        except urllib.error.HTTPError as e:
            return e.code, e.read()

    def get_json(self, path, token="testtok"):
        code, body = self._open(path, token)
        return code, json.loads(body.decode("utf-8"))

    # ---------- 用例 ----------
    def test_ping(self):
        code, d = self.get_json("/api/ping")
        self.assertEqual(code, 200)
        self.assertTrue(d["ok"])
        self.assertEqual(d["version"], version.VERSION)
        self.assertEqual(d["api_version"], version.API_VERSION)
        self.assertEqual(d["name"], version.APP_NAME)

    def test_about_has_author(self):
        code, d = self.get_json("/api/about")
        self.assertEqual(code, 200)
        a = d["about"]
        self.assertEqual(a["author"], version.AUTHOR)
        self.assertEqual(a["email"], version.EMAIL)
        self.assertTrue(a["third_party"])

    def test_deps(self):
        code, d = self.get_json("/api/deps")
        self.assertEqual(code, 200)
        self.assertTrue(d["deps"])
        for item in d["deps"]:
            self.assertIn("module", item)
            self.assertIsInstance(item["ok"], bool)
            self.assertIn("apt", item)
        self.assertEqual(d["api_version"], version.API_VERSION)

    def test_unknown_api_tells_stale_server(self):
        code, d = self.get_json("/api/definitely_not_here")
        self.assertEqual(code, 404)
        self.assertFalse(d["ok"])
        self.assertIn("旧版本", d["error"])

    def test_pick_warm_does_not_open_dialog(self):
        """mode=warm 只预热，不能弹框（否则页面一加载就蹦窗口）。"""
        old = server.prewarm_picker
        server.prewarm_picker = lambda: True
        try:
            code, d = self.get_json("/api/pick?mode=warm")
        finally:
            server.prewarm_picker = old
        self.assertEqual(code, 200)
        self.assertTrue(d["ok"])
        self.assertNotIn("paths", d)

    def test_pick_rejects_unknown_mode(self):
        code, d = self.get_json("/api/pick?mode=folder2")
        self.assertFalse(d["ok"])
        self.assertIn("mode", d["error"])

    def test_token_required(self):
        code, d = self.get_json("/api/ping", token="")
        self.assertEqual(code, 403)
        self.assertFalse(d["ok"])

    def test_index_placeholders_replaced(self):
        code, body = self._open("/")
        html = body.decode("utf-8")
        self.assertEqual(code, 200)
        self.assertIn(version.APP_NAME, html)
        self.assertIn('"testtok"', html)
        self.assertNotIn("%%ASSET", html)
        for ph in ("%%TOKEN%%", "%%APP_NAME%%", "%%API_VER%%", "%%VERSION%%"):
            self.assertNotIn(ph, html)

    def test_upload_keeps_subdir_and_blocks_traversal(self):
        payload = b"PK\x03\x04fake"
        code, body = self._open(
            "/api/upload?sid=s1&name=%E7%9B%AE%E5%BD%95/a.xlsx", data=payload)
        d = json.loads(body.decode())
        self.assertEqual(code, 200)
        self.assertTrue(d["ok"])
        self.assertTrue(os.path.isfile(d["path"]))
        with open(d["path"], "rb") as f:
            self.assertEqual(f.read(), payload)
        # 上跳要被剥掉，最终仍落在 uploads/<sid> 里
        code, body = self._open("/api/upload?sid=s1&name=../../evil.txt", data=b"x")
        d = json.loads(body.decode())
        self.assertEqual(code, 200)
        root = os.path.abspath(server.UPLOAD_DIR)
        self.assertTrue(os.path.abspath(d["path"]).startswith(root + os.sep))
        self.assertNotIn("..", os.path.abspath(d["path"]))

    def test_native_pick_without_dialog(self):
        old = shutil.which
        old_win = server._use_win_pick
        shutil.which = lambda name: None
        server._use_win_pick = lambda: False  # 强制走 Linux 分支
        try:
            r = server.native_pick("files")
        finally:
            shutil.which = old
            server._use_win_pick = old_win
        self.assertFalse(r["ok"])
        self.assertEqual(r["reason"], "no-dialog")

    def test_native_pick_dir_mode_uses_directory_flag(self):
        """zenity 存在时应带 --directory；用假 which + 假 run 验证参数拼装。"""
        calls = {}

        class FakeRun:
            returncode = 0
            stdout = b"/home/user/Desktop\n"
            stderr = b""

        old_which = shutil.which
        old_run = server.subprocess.run
        old_disp = os.environ.get("DISPLAY")
        old_win = server._use_win_pick
        shutil.which = lambda n: "/usr/bin/" + n if n == "zenity" else None
        server.subprocess.run = lambda cmd, **kw: (calls.setdefault("cmd", cmd), FakeRun())[1]
        server._use_win_pick = lambda: False
        os.environ["DISPLAY"] = ":0"
        try:
            r = server.native_pick("dir")
        finally:
            shutil.which = old_which
            server.subprocess.run = old_run
            server._use_win_pick = old_win
            if old_disp is None:
                os.environ.pop("DISPLAY", None)
            else:
                os.environ["DISPLAY"] = old_disp
        self.assertTrue(r["ok"])
        self.assertEqual(r["paths"], ["/home/user/Desktop"])
        self.assertIn("--directory", calls["cmd"])
        self.assertEqual(r["dialog"], "zenity")

    def test_native_pick_cancel(self):
        class FakeRun:
            returncode = 1
            stdout = b""
            stderr = b""

        old_which = shutil.which
        old_run = server.subprocess.run
        old_disp = os.environ.get("DISPLAY")
        old_win = server._use_win_pick
        shutil.which = lambda n: "/usr/bin/zenity" if n == "zenity" else None
        server.subprocess.run = lambda cmd, **kw: FakeRun()
        server._use_win_pick = lambda: False
        os.environ["DISPLAY"] = ":0"
        try:
            r = server.native_pick("files")
        finally:
            shutil.which = old_which
            server.subprocess.run = old_run
            server._use_win_pick = old_win
            if old_disp is None:
                os.environ.pop("DISPLAY", None)
            else:
                os.environ["DISPLAY"] = old_disp
        self.assertTrue(r["ok"])
        self.assertTrue(r.get("cancelled"))
        self.assertEqual(r["paths"], [])


class OpenTargetTest(unittest.TestCase):
    """「打开输出目录」的路径解析：不实际弹文件管理器，只验目标目录。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="oa_open_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_dir_returns_itself(self):
        self.assertEqual(server.resolve_open_target(self.tmp), os.path.abspath(self.tmp))

    def test_file_returns_parent(self):
        fp = os.path.join(self.tmp, "a.csv")
        with open(fp, "w") as f:
            f.write("x")
        self.assertEqual(server.resolve_open_target(fp), os.path.abspath(self.tmp))

    def test_missing_dir_created(self):
        target = os.path.join(self.tmp, "新目录", "子目录")
        self.assertEqual(server.resolve_open_target(target), os.path.abspath(target))
        self.assertTrue(os.path.isdir(target))

    def test_missing_dir_without_create_raises(self):
        with self.assertRaises(ValueError):
            server.resolve_open_target(os.path.join(self.tmp, "没有这个"), create=False)

    def test_empty_path_raises(self):
        with self.assertRaises(ValueError):
            server.resolve_open_target("")


class PickWinTest(unittest.TestCase):
    """Windows 一次性进程分支（兜底）：mock 掉 PowerShell，只验结果解析。"""

    def _run_with(self, writer, rc):
        class FakeRun:
            returncode = rc
            stdout = b""
            stderr = b""

        def fake_run(cmd, env=None, **kw):
            with open(env["OA_PICK_OUT"], "w", encoding="utf-8") as f:
                f.write(writer)
            return FakeRun()

        old = server.subprocess.run
        server.subprocess.run = fake_run
        try:
            return server._pick_win_once("files", "C:\\", "选择文件")
        finally:
            server.subprocess.run = old

    def test_paths(self):
        r = self._run_with("C:\\a\\b.xlsx\nC:\\资料\\c.docx\n", 0)
        self.assertTrue(r["ok"])
        self.assertEqual(r["dialog"], "windows")
        self.assertEqual(r["paths"], ["C:\\a\\b.xlsx", "C:\\资料\\c.docx"])

    def test_cancelled(self):
        r = self._run_with("", 1)
        self.assertTrue(r["ok"])
        self.assertTrue(r.get("cancelled"))
        self.assertEqual(r["paths"], [])

    def test_error_is_reported(self):
        r = self._run_with("__ERR__无法创建窗口", 2)
        self.assertFalse(r["ok"])
        self.assertEqual(r["dialog"], "windows")
        self.assertIn("无法创建窗口", r["error"])


class FakeProc(object):
    """假 PowerShell 宿主：只吐出预先准备好的几行，不做真实对话框。"""

    def __init__(self, lines=b"READY\n", exit_code=None):
        self.stdout = io.BytesIO(lines)
        self.stdin = io.BytesIO()
        self._rc = exit_code

    def poll(self):
        return self._rc

    def wait(self, timeout=None):
        self._rc = 0
        return 0

    def kill(self):
        self._rc = -9


class PickerHostTest(unittest.TestCase):
    """常驻对话框宿主：只验进程生命周期与失败降级，不弹真窗口。"""

    def test_start_reads_ready(self):
        host = server._PickerHost()
        fake = FakeProc()
        old = server.subprocess.Popen
        server.subprocess.Popen = lambda *a, **k: fake
        try:
            ok, err = host.start(timeout=5)
            self.assertTrue(ok, err)
            self.assertTrue(host.alive())
            host.shutdown()
            self.assertFalse(host.alive())
        finally:
            server.subprocess.Popen = old

    def test_start_failure_is_reported(self):
        host = server._PickerHost()

        def boom(*a, **k):
            raise OSError("powershell 不存在")

        old = server.subprocess.Popen
        server.subprocess.Popen = boom
        try:
            ok, err = host.start(timeout=5)
        finally:
            server.subprocess.Popen = old
        self.assertFalse(ok)
        self.assertIn("powershell 不存在", err)
        self.assertFalse(host.alive())

    def test_pick_falls_back_to_once_when_host_dead(self):
        """宿主起不来时要退回一次性进程，而不是直接报错。"""
        class DeadHost:
            def pick(self, *a, **k):
                return {"ok": False, "reason": "error", "error": "宿主挂了"}

            def alive(self):
                return False

        calls = {}

        class FakeRun:
            returncode = 1
            stdout = b""
            stderr = b""

        def fake_run(cmd, env=None, **kw):
            calls["env"] = env
            open(env["OA_PICK_OUT"], "w", encoding="utf-8").close()
            return FakeRun()

        old_host, old_run = server._PICKER, server.subprocess.run
        server._PICKER, server.subprocess.run = DeadHost(), fake_run
        try:
            r = server._pick_win("files", "", "选择文件")
        finally:
            server._PICKER, server.subprocess.run = old_host, old_run
        self.assertTrue(r["ok"])
        self.assertTrue(r.get("cancelled"))


class VersionTest(unittest.TestCase):
    def test_about_payload(self):
        a = version.about()
        for k in ("name", "version", "author", "email", "homepage",
                  "license", "description", "api_version", "third_party"):
            self.assertIn(k, a)
        self.assertTrue(a["name"])
        self.assertTrue(a["version"])

    def test_version_format(self):
        parts = version.VERSION.split(".")
        self.assertEqual(len(parts), 3)
        for p in parts:
            self.assertTrue(p.isdigit())


if __name__ == "__main__":
    unittest.main()
