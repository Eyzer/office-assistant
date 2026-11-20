"""人员信息提取的单元测试。

样本全部现造 docx，不依赖外部文件。
"""
import os
import shutil
import tempfile
import unittest

import docx

from tools.word import doc_info_extract as m


def make_doc(path, blocks):
    """blocks: [('p', 文本) 或 ('t', [[r1c1,r1c2...], ...])]"""
    d = docx.Document()
    for kind, data in blocks:
        if kind == "p":
            d.add_paragraph(data)
        else:
            rows, cols = len(data), max(len(r) for r in data)
            t = d.add_table(rows=rows, cols=cols)
            t.style = "Table Grid"
            for ri, row in enumerate(data):
                for ci, v in enumerate(row):
                    t.cell(ri, ci).text = v
    d.save(path)
    return path


class IdChecksumTest(unittest.TestCase):
    def test_valid(self):
        for s in ("110101199005121230", "320102198807154529", "440305199512030010"):
            self.assertTrue(m.id_checksum_ok(s), s)

    def test_invalid(self):
        # 校验位错一位、含字母、长度不对
        self.assertFalse(m.id_checksum_ok("110101199005121234"))
        self.assertFalse(m.id_checksum_ok("11010119900512123"))
        self.assertFalse(m.id_checksum_ok("11010119900512123A"))


class ExtractLayoutTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _one(self, blocks):
        p = os.path.join(self.tmp, "t.docx")
        make_doc(p, blocks)
        return m.extract(p)

    def test_grid_horizontal(self):
        """[姓名][张三][性别][男] 横向铺开。"""
        bag = self._one([("t", [
            ["姓名", "张三", "性别", "男", "民族", "汉族"],
            ["身份证号", "110101199005121230", "单位", "某部"],
        ])])
        self.assertEqual(bag["姓名"], "张三")
        self.assertEqual(bag["性别"], "男")
        self.assertEqual(bag["民族"], "汉族")
        self.assertEqual(bag["身份证号"], "110101199005121230")
        self.assertEqual(bag["_id_status"], "校验通过")

    def test_cell_with_colon(self):
        """单元格内「姓名：张三」。"""
        bag = self._one([("t", [
            ["姓名：", "李四", "性别：", "女"],
            ["身份证号：", "320102198807154529", "民族：", "回族"],
        ])])
        self.assertEqual(bag["姓名"], "李四")
        self.assertEqual(bag["身份证号"], "320102198807154529")
        self.assertEqual(bag["民族"], "回族")

    def test_vertical_header(self):
        """第一行全是标签、第二行全是值。"""
        bag = self._one([("t", [
            ["姓名", "性别", "民族"],
            ["王五", "男", "满族"],
        ])])
        self.assertEqual(bag["姓名"], "王五")
        self.assertEqual(bag["性别"], "男")
        self.assertEqual(bag["民族"], "满族")

    def test_paragraph_inline(self):
        """段落里一行多组，空格分隔。"""
        bag = self._one([("p", "姓名：赵六    性别：男    民族：汉族"),
                         ("p", "身份证号：110101199005121230")])
        self.assertEqual(bag["姓名"], "赵六")
        self.assertEqual(bag["性别"], "男")
        self.assertEqual(bag["身份证号"], "110101199005121230")

    def test_single_column_vertical(self):
        """单列竖排：[姓名]/[张三]/[性别]/[男]，一格标签一格值交替。"""
        bag = self._one([("t", [
            ["姓名"], ["张三"], ["性别"], ["男"],
        ])])
        self.assertEqual(bag["姓名"], "张三")
        self.assertEqual(bag["性别"], "男")

    def test_empty_column_between(self):
        """标签和值之间夹一个空列也要能连上。"""
        bag = self._one([("t", [
            ["姓名", "", "张三", "性别", "", "男"],
        ])])
        self.assertEqual(bag["姓名"], "张三")
        self.assertEqual(bag["性别"], "男")

    def test_header_only_no_value(self):
        """纯表头行不能把右边的标签当值。"""
        bag = self._one([("t", [
            ["姓名", "性别", "民族"],
            ["", "", ""],
        ])])
        self.assertNotIn("姓名", bag)
        self.assertNotIn("性别", bag)

    def test_no_false_positive(self):
        """普通文档不该抓出任何字段。"""
        bag = self._one([("p", "今天下午三点召开连务会，讨论训练计划安排。"),
                         ("p", "参会人员：全体干部骨干。")])
        self.assertNotIn("姓名", bag)
        self.assertNotIn("身份证号", bag)
        self.assertEqual(bag["_id_status"], "")


class IdFallbackTest(unittest.TestCase):
    """身份证号没有标签也要抓到——但绝不能跨行粘连。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_orphan_id_in_table(self):
        p = os.path.join(self.tmp, "a.docx")
        make_doc(p, [("t", [
            ["姓名", "孙七", "性别", "男"],
            ["", "110101199005121230", "", "13500001234"],
        ])])
        bag = m.extract(p)
        self.assertEqual(bag["身份证号"], "110101199005121230")
        self.assertEqual(bag["联系电话"], "13500001234")

    def test_id_not_glued_to_phone(self):
        """身份证和手机号相邻两行，缝空格的正则不能把它们粘成一个数。

        这是真实踩过的坑：早期版本把 \\n 也缝掉了，29 位数字让
        (?!\\d) 断言直接失效，身份证和电话一个都抓不到。
        """
        p = os.path.join(self.tmp, "b.docx")
        make_doc(p, [("p", "110101199005121230"), ("p", "13500001234")])
        bag = m.extract(p)
        self.assertEqual(bag["身份证号"], "110101199005121230")

    def test_space_folded_id(self):
        """同一行内被空格分开的身份证。"""
        p = os.path.join(self.tmp, "c.docx")
        make_doc(p, [("p", "110101 19900512 1230")])
        bag = m.extract(p)
        self.assertEqual(bag["身份证号"], "110101199005121230")

    def test_bad_checksum_still_kept(self):
        """校验位错的也要抓出来，只标状态，不能丢。"""
        p = os.path.join(self.tmp, "d.docx")
        make_doc(p, [("t", [["身份证号", "110101199005121234"]])])
        bag = m.extract(p)
        self.assertEqual(bag["身份证号"], "110101199005121234")
        self.assertEqual(bag["_id_status"], "校验位错")

    def test_birth_from_id(self):
        """没有出生日期字段时，从身份证里推。"""
        p = os.path.join(self.tmp, "e.docx")
        make_doc(p, [("t", [["身份证号", "110101199005121230"]])])
        self.assertEqual(m.extract(p)["出生日期"], "1990-05-12")


class RosterTest(unittest.TestCase):
    """表头 + 多行数据的名单表要拆成多条记录。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_split_rows(self):
        p = os.path.join(self.tmp, "r.docx")
        make_doc(p, [("t", [
            ["姓名", "性别", "身份证号", "单位"],
            ["孙七", "男", "110101199005121230", "一连"],
            ["周八", "女", "320102198807154529", "二连"],
        ])])
        bags, tips = m.extract_multi(p)
        self.assertEqual(len(bags), 2)
        self.assertEqual(bags[0]["姓名"], "孙七")
        self.assertEqual(bags[1]["姓名"], "周八")
        self.assertEqual(bags[1]["身份证号"], "320102198807154529")
        self.assertTrue(any("拆出" in t for t in tips))

    def test_single_person_not_split(self):
        """一人一表不该被当成名单表拆开。"""
        p = os.path.join(self.tmp, "s.docx")
        make_doc(p, [("t", [
            ["姓名", "郑十", "性别", "男"],
            ["身份证号", "110101199005121230", "单位", "直属队"],
        ])])
        bags, _ = m.extract_multi(p)
        self.assertEqual(len(bags), 1)
        self.assertEqual(bags[0]["姓名"], "郑十")


class RunTaskTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.src = os.path.join(self.tmp, "in")
        self.out = os.path.join(self.tmp, "out")
        os.makedirs(self.src)
        os.makedirs(self.out)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _ctx(self):
        class LG:
            lines = []
        class C:
            cancelled = False
            task = {"log": LG()}

            def info(self, x):
                LG.lines.append(x)

            def warn(self, x):
                LG.lines.append(x)

            def error(self, x):
                LG.lines.append(x)

            def progress(self, a, b, c=""):
                pass
        return C()

    def test_summary_xlsx(self):
        make_doc(os.path.join(self.src, "甲.docx"),
                 [("t", [["姓名", "甲", "身份证号", "110101199005121230"]])])
        make_doc(os.path.join(self.src, "乙.docx"),
                 [("t", [["姓名", "乙", "身份证号", "320102198807154529"]])])
        r = m.run_task([self.src], self.out,
                       {"wanted": "姓名,身份证号", "fmt": "xlsx"}, self._ctx())
        self.assertEqual(r["success_count"], 2)
        self.assertEqual(r["fail_count"], 0)
        dst = r["output_files"][0]
        self.assertTrue(dst.endswith(".xlsx"))
        import openpyxl
        ws = openpyxl.load_workbook(dst).active
        vals = [[c for c in row] for row in ws.iter_rows(min_row=2, values_only=True)]
        names = [row[1] for row in vals]
        self.assertIn("甲", names)
        self.assertIn("乙", names)

    def test_empty_dir_reports(self):
        r = m.run_task([self.src], self.out, {"wanted": "姓名"}, self._ctx())
        self.assertEqual(r["success_count"], 0)
        self.assertEqual(r["output_files"], [])


if __name__ == "__main__":
    unittest.main()
