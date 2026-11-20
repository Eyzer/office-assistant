"""P2 工具测试。"""
import os
import shutil
import tempfile
import unittest

from tests.test_tools import SAMPLE, run


class TestExcelP2(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_advanced_sort(self):
        res, _ = run("excel", "advanced_sort",
                     [os.path.join(SAMPLE, "大表.xlsx")], self.tmp,
                     {"cols": "D", "desc": True})
        self.assertEqual(res["fail_count"], 0)

    def test_super_find(self):
        res, _ = run("excel", "super_find", [SAMPLE], self.tmp, {"keyword": "员工"})
        self.assertEqual(res["fail_count"], 0)
        self.assertIn("命中", open(res["output_files"][0], encoding="utf-8-sig").read())

    def test_data_dup(self):
        res, _ = run("excel", "data_dup",
                     [os.path.join(SAMPLE, "大表.xlsx")], self.tmp, {"cols": "B"})
        self.assertEqual(res["fail_count"], 0)

    def test_random_sample(self):
        res, _ = run("excel", "random_sample",
                     [os.path.join(SAMPLE, "大表.xlsx")], self.tmp, {"count": 5})
        self.assertEqual(res["fail_count"], 0)

    def test_hyperlink_build(self):
        res, _ = run("excel", "hyperlink_batch", [SAMPLE], self.tmp, {"mode": "build"})
        self.assertEqual(res["fail_count"], 0)
        self.assertTrue(os.path.exists(res["output_files"][0]))

    def test_format_convert(self):
        res, _ = run("excel", "format_convert",
                     [os.path.join(SAMPLE, "销售表2.xlsx")], self.tmp, {"direction": "to_csv"})
        self.assertEqual(res["fail_count"], 0)
        self.assertEqual(len(res["output_files"]), 2)

    def test_extract_img_runs(self):
        res, _ = run("excel", "excel_extract_img",
                     [os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp, {})
        self.assertEqual(res["fail_count"], 0)


class TestWordP2(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_header_footer(self):
        res, _ = run("word", "doc_header_footer",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp,
                     {"header": "内部文件", "page_number": True})
        self.assertEqual(res["fail_count"], 0)

    def test_page_op_extract(self):
        res, _ = run("word", "doc_page_op",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp,
                     {"mode": "extract", "pages": "1"})
        self.assertEqual(res["fail_count"], 0)

    def test_img_layout(self):
        res, _ = run("word", "doc_img_layout",
                     [os.path.join(SAMPLE, "文档1.docx")], self.tmp, {"width_cm": 10})
        self.assertEqual(res["fail_count"], 0)

    def test_mail_merge(self):
        import docx
        tpl = os.path.join(self.tmp, "模板.docx")
        d = docx.Document()
        d.add_paragraph("姓名：{{姓名}}")
        d.add_paragraph("部门：{{部门}}")
        d.save(tpl)
        res, _ = run("word", "doc_mail_merge",
                     [tpl, os.path.join(SAMPLE, "销售表1.xlsx")], self.tmp,
                     {"template": tpl, "data": os.path.join(SAMPLE, "销售表1.xlsx"),
                      "name_field": "姓名"})
        self.assertEqual(res["fail_count"], 0)
        self.assertGreaterEqual(len(res["output_files"]), 1)
        text = "\n".join(p.text for p in docx.Document(res["output_files"][0]).paragraphs)
        self.assertNotIn("{{", text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
