"""冒烟测试：每个工具都能被调用且不抛异常（失败计数允许）。"""
import os
import shutil
import tempfile
import unittest

from service import task_service
from tests.test_tools import SAMPLE, FakeCtx, load_mod


class TestSmoke(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_all_tools_callable(self):
        tools = task_service.tool_list()
        self.assertGreaterEqual(len(tools), 50)
        problems = []
        for t in tools:
            mod = load_mod(t["group"], t["id"])
            params = {}
            for f in t.get("fields", []):
                if "default" in f:
                    params[f["key"]] = f["default"]
            ctx = FakeCtx()
            try:
                res = mod.run_task([SAMPLE], self.tmp, params, ctx)
                assert isinstance(res, dict), "返回值不是 dict"
                for key in ("success_count", "fail_count", "output_files"):
                    assert key in res, f"缺少 {key}"
            except Exception as e:
                problems.append(f"{t['id']}: {e}")
        self.assertEqual(problems, [], "\n".join(problems))

    def test_no_tool_writes_source(self):
        """确认源目录文件数量不变（不原地改源）。"""
        before = sorted(os.listdir(SAMPLE))
        for t in task_service.tool_list():
            mod = load_mod(t["group"], t["id"])
            params = {f["key"]: f.get("default") for f in t.get("fields", []) if "default" in f}
            try:
                mod.run_task([SAMPLE], self.tmp, params, FakeCtx())
            except Exception:
                pass
        after = sorted(os.listdir(SAMPLE))
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main(verbosity=2)
