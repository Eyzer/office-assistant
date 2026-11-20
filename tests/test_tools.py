"""工具测试：直接调 run_task，验证输出存在且内容正确。"""
import os
import shutil
import sys
import tempfile
import unittest

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

from utils.log_utils import TaskLog  # noqa: E402


class FakeCtx:
    def __init__(self):
        self.task = {"log": TaskLog("test")}
        self.cancelled = False
        self.progress_log = []

    def progress(self, cur, total, name=""):
        self.progress_log.append((cur, total, name))

    def info(self, m):
        self.task["log"].info(m)

    def warn(self, m):
        self.task["log"].warn(m)

    def error(self, m):
        self.task["log"].error(m)

    def ok(self, m):
        self.task["log"].ok(m)

    def check_cancel(self):
        pass


SAMPLE = os.path.join(BASE, "samples", "data")


def load_mod(group, name):
    import importlib
    return importlib.import_module(f"tools.{group}.{name}")


def run(group, name, inputs, outdir, params=None):
    mod = load_mod(group, name)
    ctx = FakeCtx()
    return mod.run_task(inputs, outdir, params or {}, ctx), ctx


class TestExcel(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_merge_all(self):
        res, _ = run("excel", "merge_all",
                     [os.path.join(SAMPLE, "销售表1.xlsx"),
                      os.path.join(SAMPLE, "销售表2.xlsx")], self.tmp)
        self.assertEqual(res["fail_count"], 0)
        self.assertTrue(os.path.exists(res["output_files"][0]))
        import openpyxl
        wb = openpyxl.load_workbook(res["output_files"][0])
        ws = wb.active
        # 默认合并每文件第一个表：表头 1 行 + 数据 3+2 行
        self.assertEqual(ws.max_row, 6)

    def test_split_by_col(self):
        res, _ = run("excel", "split_by_col",
                     [os.path.join(SAMPLE, "大表.xlsx")], self.tmp,
                     {"col": "B", "mode": "files"})
        self.assertEqual(res["fail_count"], 0)
        self.assertGreaterEqual(len(res["output_files"]), 3)

    def test_split_by_row(self):
        res, _ = run("excel", "split_by_row",
                     [os.path.join(SAMPLE, "大表.xlsx")], self.tmp,
                     {"rows": 20, "mode": "sheets"})
        self.assertEqual(res["fail_count"], 0)
        import openpyxl
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertGreaterEqual(len(wb.worksheets), 3)

    def test_split_by_sheet(self):
        res, _ = run("excel", "split_by_sheet",
                     [os.path.join(SAMPLE, "销售表2.xlsx")], self.tmp)
        self.assertEqual(res["fail_count"], 0)
        self.assertEqual(len(res["output_files"]), 2)

    def test_unmerge_fill(self):
        res, _ = run("excel", "cell_unmerge_fill",
                     [os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp)
        self.assertEqual(res["fail_count"], 0)
        import openpyxl
        wb = openpyxl.load_workbook(res["output_files"][0])
        ws = wb.active
        self.assertEqual(len(ws.merged_cells.ranges), 0)
        self.assertEqual(ws["A3"].value, "A组")

    def test_data_clean(self):
        res, _ = run("excel", "data_clean",
                     [os.path.join(SAMPLE, "大表.xlsx")], self.tmp,
                     {"find": "员工", "replace": "职员"})
        self.assertEqual(res["fail_count"], 0)
        import openpyxl
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertIn("职员1", [c.value for row in wb.active.iter_rows() for c in row])

    def test_num_format(self):
        res, _ = run("excel", "num_format",
                     [os.path.join(SAMPLE, "大表.xlsx")], self.tmp,
                     {"mode": "thousand", "col": "D"})
        self.assertEqual(res["fail_count"], 0)


class TestWord(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_doc_clean(self):
        res, _ = run("word", "doc_clean",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp)
        self.assertEqual(res["fail_count"], 0)
        import docx
        d = docx.Document(res["output_files"][0])
        self.assertLessEqual(len(d.paragraphs), 3)

    def test_text_replace(self):
        res, _ = run("word", "doc_text_replace",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp,
                     {"pairs": "旧词=>新词"})
        self.assertEqual(res["fail_count"], 0)
        import docx
        text = "\n".join(p.text for p in docx.Document(res["output_files"][0]).paragraphs)
        self.assertIn("新词", text)
        self.assertNotIn("旧词", text)

    def test_doc_merge(self):
        res, _ = run("word", "doc_merge",
                     [os.path.join(SAMPLE, "文档1.docx"),
                      os.path.join(SAMPLE, "文档2.docx")], self.tmp)
        self.assertEqual(res["fail_count"], 0)
        import docx
        text = "\n".join(p.text for p in docx.Document(res["output_files"][0]).paragraphs)
        self.assertIn("另一份文档", text)

    def test_doc_metadata(self):
        res, _ = run("word", "doc_metadata",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp,
                     {"clear": True})
        self.assertEqual(res["fail_count"], 0)
        import docx
        d = docx.Document(res["output_files"][0])
        self.assertEqual(d.core_properties.author, "")


class TestCommon(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_batch_rename(self):
        res, _ = run("common", "batch_rename",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp,
                     {"pattern": "第{index}份_{name}"})
        self.assertEqual(res["fail_count"], 0)
        self.assertTrue(os.path.basename(res["output_files"][0]).startswith("第001份_"))

    def test_flatten(self):
        res, _ = run("common", "flatten_dir", [SAMPLE], self.tmp)
        self.assertEqual(res["fail_count"], 0)
        self.assertEqual(len(res["output_files"]), len(os.listdir(SAMPLE)))


if __name__ == "__main__":
    unittest.main(verbosity=2)
