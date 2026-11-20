"""P1 工具测试。"""
import os
import shutil
import tempfile
import unittest

from tests.test_tools import SAMPLE, run


class TestExcelP1(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_merge_modes_align(self):
        res, _ = run("excel", "merge_modes",
                     [os.path.join(SAMPLE, "销售表1.xlsx"),
                      os.path.join(SAMPLE, "销售表2.xlsx")], self.tmp,
                     {"mode": "align"})
        self.assertEqual(res["fail_count"], 0)
        self.assertTrue(os.path.exists(res["output_files"][0]))

    def test_merge_folder(self):
        res, _ = run("excel", "merge_folder", [SAMPLE], self.tmp, {})
        self.assertEqual(res["fail_count"], 0)
        self.assertTrue(os.path.exists(res["output_files"][0]))

    def test_pivot_to_long(self):
        res, _ = run("excel", "pivot_flatten",
                     [os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp,
                     {"mode": "to_long"})
        self.assertEqual(res["fail_count"], 0)
        import openpyxl
        ws = openpyxl.load_workbook(res["output_files"][0]).active
        self.assertEqual(ws.cell(row=1, column=1).value, "行标签")

    def test_cell_join(self):
        res, _ = run("excel", "cell_join",
                     [os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp,
                     {"cols": "B,C", "sep": "-", "out_col": "F"})
        self.assertEqual(res["fail_count"], 0)
        import openpyxl
        ws = openpyxl.load_workbook(res["output_files"][0]).active
        self.assertIn("-", str(ws["F2"].value))

    def test_cell_split(self):
        res, _ = run("excel", "cell_split",
                     [os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp,
                     {"col": "C", "sep": "部", "max_cols": 2})
        self.assertEqual(res["fail_count"], 0)

    def test_cell_merge_same(self):
        res, _ = run("excel", "cell_merge_same",
                     [os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp,
                     {"cols": "B", "start_row": 2})
        self.assertEqual(res["fail_count"], 0)

    def test_format_standard(self):
        res, _ = run("excel", "format_standard",
                     [os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp,
                     {"font": "黑体", "size": 10})
        self.assertEqual(res["fail_count"], 0)

    def test_sheet_batch(self):
        res, _ = run("excel", "sheet_batch",
                     [os.path.join(SAMPLE, "销售表2.xlsx")], self.tmp,
                     {"rename": "表{index}"})
        self.assertEqual(res["fail_count"], 0)
        import openpyxl
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertEqual(wb.worksheets[0].title, "表1")

    def test_formula_to_value(self):
        import openpyxl
        src = os.path.join(self.tmp, "f.xlsx")
        wb = openpyxl.Workbook()
        wb.active["A1"] = 1
        wb.active["A2"] = 2
        wb.active["A3"] = "=SUM(A1:A2)"
        wb.save(src)
        res, _ = run("excel", "formula_to_value", [src], self.tmp, {"keep_calc": False})
        self.assertEqual(res["fail_count"], 0)
        ws = openpyxl.load_workbook(res["output_files"][0]).active
        self.assertIsNone(ws["A3"].value)

    def test_workbook_compare(self):
        res, _ = run("excel", "workbook_compare",
                     [os.path.join(SAMPLE, "销售表1.xlsx"),
                      os.path.join(SAMPLE, "销售表2.xlsx")], self.tmp, {})
        self.assertEqual(res["fail_count"], 0)
        self.assertIn("差异", open(res["output_files"][0], encoding="utf-8-sig").read())


class TestWordP1(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_extract_text(self):
        res, _ = run("word", "doc_extract_text",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp, {})
        self.assertEqual(res["fail_count"], 0)
        self.assertIn("第一段", open(res["output_files"][0], encoding="utf-8-sig").read())

    def test_extract_img(self):
        res, _ = run("word", "doc_extract_img",
                     [os.path.join(SAMPLE, "文档2.docx")], self.tmp, {})
        self.assertEqual(res["fail_count"], 0)

    def test_compress(self):
        res, _ = run("word", "doc_compress",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp, {})
        self.assertEqual(res["fail_count"], 0)
        self.assertTrue(os.path.exists(res["output_files"][0]))

    def test_page_style(self):
        res, _ = run("word", "doc_page_style",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp,
                     {"margin": 2, "paper": "A4"})
        self.assertEqual(res["fail_count"], 0)

    def test_content_style(self):
        res, _ = run("word", "doc_content_style",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp,
                     {"font": "黑体", "size": 14})
        self.assertEqual(res["fail_count"], 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
