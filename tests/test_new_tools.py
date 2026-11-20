"""新补工具的测试：易用宝对标功能 + 内置依赖 + 拖拽上传落盘。"""
import csv
import json
import os
import shutil
import tempfile
import unittest

import openpyxl

from tests.test_tools import BASE, FakeCtx, load_mod


def run(group, name, inputs, out, params=None):
    mod = load_mod(group, name)
    ctx = FakeCtx()
    res = mod.run_task(list(inputs), out, params or {}, ctx)
    return res, ctx


def make_book(path, rows, title="数据"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.out = os.path.join(self.tmp, "out")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def book(self, rows, name="a.xlsx"):
        return make_book(os.path.join(self.tmp, name), rows)


class TestRegionAndHierarchy(Base):
    def test_transpose(self):
        f = self.book([["a", "b", "c"], [1, 2, 3]])
        res, _ = run("excel", "transpose_ops", [f], self.out, {"mode": "transpose"})
        self.assertEqual(res["fail_count"], 0)
        wb = openpyxl.load_workbook(res["output_files"][0])
        ws = wb["转置_数据"]
        self.assertEqual(ws.cell(row=1, column=1).value, "a")
        self.assertEqual(ws.cell(row=1, column=2).value, 1)

    def test_one2two(self):
        f = self.book([["行", "列", "值"], ["甲", "一月", 1], ["甲", "二月", 2], ["乙", "一月", 3]])
        res, _ = run("excel", "transpose_ops", [f], self.out,
                     {"mode": "one2two", "key_col": "A", "item_col": "B", "value_col": "C"})
        ws = openpyxl.load_workbook(res["output_files"][0])["二维_数据"]
        self.assertEqual(ws.cell(row=1, column=2).value, "一月")
        self.assertEqual(ws.cell(row=2, column=2).value, 1)
        self.assertEqual(ws.cell(row=2, column=3).value, 2)

    def test_single2multi(self):
        f = self.book([[i] for i in range(1, 8)])
        res, _ = run("excel", "transpose_ops", [f], self.out,
                     {"mode": "single2multi", "cols_per_row": 3})
        ws = openpyxl.load_workbook(res["output_files"][0])["多行多列_数据"]
        self.assertEqual([ws.cell(row=1, column=c).value for c in (1, 2, 3)], [1, 2, 3])

    def test_hierarchy_number(self):
        f = self.book([["层级"], [1], [2], [2], [1], [2]])
        res, _ = run("excel", "hierarchy_ops", [f], self.out,
                     {"mode": "number", "level_col": "A"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        vals = [ws.cell(row=r, column=2).value for r in range(2, 7)]
        self.assertEqual(vals, ["1", "1.1", "1.2", "2", "2.1"])

    def test_hierarchy_split_merge(self):
        f = self.book([["层级"], ["广东省/深圳市/南山区"]])
        res, _ = run("excel", "hierarchy_ops", [f], self.out,
                     {"mode": "split", "level_col": "A", "sep": "/"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=3).value, "深圳市")

    def test_swap(self):
        f = self.book([["a", "b", "c", "d"]])
        res, _ = run("excel", "transpose_ops", [f], self.out,
                     {"mode": "swap", "range1": "A1:B1", "range2": "C1:D1"})
        ws = openpyxl.load_workbook(res["output_files"][0]).active
        self.assertEqual([ws.cell(row=1, column=c).value for c in range(1, 5)], ["c", "d", "a", "b"])


class TestDateAndNumber(Base):
    def test_parts(self):
        f = self.book([["日期"], ["2024年3月5日"]])
        res, _ = run("excel", "date_ops", [f], self.out,
                     {"mode": "parts", "col": "A", "parts": "年,月,日,星期"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=2).value, 2024)
        self.assertEqual(ws.cell(row=2, column=5).value, "星期二")

    def test_normalize_and_clear(self):
        f = self.book([["日期"], ["2024/3/5"], ["不是日期"]])
        res, _ = run("excel", "date_ops", [f], self.out, {"mode": "normalize", "col": "A"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=1).value.strftime("%Y-%m-%d"), "2024-03-05")
        self.assertEqual(ws.cell(row=3, column=1).value, "不是日期")

        res, _ = run("excel", "date_ops", [f], self.out, {"mode": "clear", "col": "A"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertIsNone(ws.cell(row=3, column=1).value)

    def test_timestamp_roundtrip(self):
        f = self.book([["ts"], [1700000000]])
        res, _ = run("excel", "date_ops", [f], self.out,
                     {"mode": "timestamp", "col": "A", "direction": "to_date"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=2).value.year, 2023)

    def test_workday(self):
        f = self.book([["开始", "结束"], ["2024-10-01", "2024-10-07"]])
        res, _ = run("excel", "date_ops", [f], self.out,
                     {"mode": "workday", "col": "A", "col2": "B",
                      "holidays": "2024-10-01,2024-10-02,2024-10-03"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        # 10/1~10/7 共 5 个工作日（周二、三、四、五、次周一），
        # 扣除 1~3 号假期后剩 2 天（4 号与 7 号）
        self.assertEqual(ws.cell(row=2, column=3).value, 2)

    def test_round_half_even(self):
        f = self.book([["n"], [2.5], [3.5], [1.005]])
        res, _ = run("excel", "round_ops", [f], self.out,
                     {"mode": "half_even", "col": "A", "digits": 0})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=1).value, 2)
        self.assertEqual(ws.cell(row=3, column=1).value, 4)

    def test_round_level(self):
        f = self.book([["n"], [95], [75], [10]])
        res, _ = run("excel", "round_ops", [f], self.out,
                     {"mode": "level", "col": "A", "rules": "0-60:不及格,60-90:及格,90-100:优秀"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual([ws.cell(row=r, column=1).value for r in (2, 3, 4)], ["优秀", "及格", "不及格"])

    def test_sum_total_and_merged(self):
        f = self.book([["k", "v"], ["a", 1], ["a", 2], ["b", 3]])
        res, _ = run("excel", "sum_ops", [f], self.out, {"mode": "total", "value_col": "B"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=5, column=1).value, "合计")
        self.assertEqual(ws.cell(row=5, column=2).value, 6)

        res, _ = run("excel", "sum_ops", [f], self.out,
                     {"mode": "merged", "col": "A", "value_col": "B"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=2).value, 3)
        self.assertEqual(ws.cell(row=4, column=2).value, 3)


class TestText(Base):
    def test_prefix_and_spaces(self):
        f = self.book([["名"], ["张三"], [" 李 四 "]])
        res, _ = run("excel", "text_ops", [f], self.out,
                     {"mode": "prefix", "col": "A", "text": "#"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=1).value, "#张三")

        res, _ = run("excel", "text_ops", [f], self.out, {"mode": "spaces", "col": "A"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=3, column=1).value, "李四")

    def test_extract_phone_and_unique(self):
        f = self.book([["备注"], ["电话13812345678"], ["另一个13900001111"], ["无"]])
        res, _ = run("excel", "text_extract", [f], self.out, {"mode": "phone", "col": "A"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=2).value, "13812345678")

        res, _ = run("excel", "text_extract", [f], self.out, {"mode": "unique", "col": "A"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertIn("唯一值_数据", wb.sheetnames)

    def test_pinyin_and_english(self):
        f = self.book([["名", "数"], ["张三", 123]])
        res, _ = run("excel", "text_convert", [f], self.out,
                     {"mode": "pinyin", "col": "A", "sep": "-"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=3).value, "zhang-san")

        res, _ = run("excel", "text_convert", [f], self.out, {"mode": "num2en", "col": "B"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=3).value, "one hundred and twenty-three")

    def test_rmb_reverse_and_base(self):
        from tools.excel.text_convert import rmb2num, to_base, safe_eval
        self.assertEqual(rmb2num("壹佰贰拾叁元肆角伍分"), 123.45)
        self.assertEqual(rmb2num("零元整"), 0.0)
        self.assertEqual(to_base("FF", 16, 10), "255")
        self.assertEqual(safe_eval("1+2*3"), 7)
        self.assertEqual(safe_eval("=8/2"), 4.0)


class TestSheetsAndFormula(Base):
    def test_catalog_and_tab_color(self):
        f = self.book([["a"], [1]])
        res, _ = run("excel", "sheet_ops", [f], self.out, {"mode": "catalog"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertIn("目录", wb.sheetnames)

        res, _ = run("excel", "sheet_ops", [f], self.out,
                     {"mode": "tab_color", "color": "ff0000"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertIn("FF0000", str(wb["数据"].sheet_properties.tabColor.rgb))

    def test_protect_and_freeze(self):
        f = self.book([["a"], [1]])
        res, _ = run("excel", "sheet_ops", [f], self.out,
                     {"mode": "protect", "password": "123"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertTrue(wb["数据"].protection.sheet)

        res, _ = run("excel", "sheet_ops", [f], self.out, {"mode": "freeze", "freeze": "row"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertEqual(wb["数据"].freeze_panes, "A2")

    def test_new_blank(self):
        f = self.book([["a"]])
        res, _ = run("excel", "sheet_ops", [f], self.out,
                     {"mode": "new_blank", "names": "一月,二月"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertIn("一月", wb.sheetnames)
        self.assertIn("二月", wb.sheetnames)

    def test_formula_fill(self):
        f = self.book([["a", "b"], [2, 3], [4, 5]])
        res, _ = run("excel", "formula_ops", [f], self.out,
                     {"mode": "fill", "col": "C", "template": "=A{r}*B{r}"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=3).value, "=A2*B2")
        self.assertEqual(ws.cell(row=3, column=3).value, "=A3*B3")

    def test_formula_ref_absolute(self):
        from tools.excel.formula_ops import to_absolute, to_relative
        self.assertEqual(to_absolute("=A1+B2"), "=$A$1+$B$2")
        self.assertEqual(to_relative("=$A$1+$B$2"), "=A1+B2")

    def test_split_by_keyword(self):
        f = self.book([["部门", "人"], ["销售部", "甲"], ["财务部", "乙"], ["人事部", "丙"]])
        res, _ = run("excel", "split_by_keyword", [f], self.out,
                     {"col": "A", "keywords": "销售,财务"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        self.assertIn("销售", wb.sheetnames)
        self.assertIn("财务", wb.sheetnames)
        self.assertIn("其他", wb.sheetnames)
        self.assertEqual(wb["销售"].max_row, 2)

    def test_compare_columns(self):
        f = self.book([["左", "右"], ["a", "a"], ["b", "c"]])
        res, _ = run("excel", "compare_ops", [f], self.out,
                     {"mode": "columns", "col_a": "A", "col_b": "B"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        ws = wb["对比_数据"]
        kinds = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
        self.assertIn("两边都有", kinds)
        self.assertIn("只在左列", kinds)

    def test_subtotal(self):
        f = self.book([["部门", "额"], ["销售", 100], ["销售", 200], ["财务", 50]])
        res, _ = run("excel", "subtotal_ops", [f], self.out,
                     {"group_col": "A", "value_cols": "B", "func": "sum"})
        wb = openpyxl.load_workbook(res["output_files"][0])
        ws = wb["汇总_数据"]
        self.assertEqual(ws.cell(row=2, column=2).value, 300)
        self.assertEqual(ws.cell(row=4, column=1).value, "合计")


class TestCommentBarcode(Base):
    def test_comment_extract_delete(self):
        from openpyxl.comments import Comment
        f = self.book([["a"], [1]])
        wb = openpyxl.load_workbook(f)
        wb["数据"]["A1"].comment = Comment("待核实", "管理员")
        wb.save(f)

        res, _ = run("excel", "comment_ops", [f], self.out, {"mode": "extract"})
        wb2 = openpyxl.load_workbook(res["output_files"][0])
        ws = wb2["批注清单"]
        self.assertEqual(ws.cell(row=2, column=4).value, "待核实")

        res, _ = run("excel", "comment_ops", [f], self.out, {"mode": "delete"})
        wb2 = openpyxl.load_workbook(res["output_files"][0])
        self.assertIsNone(wb2["数据"]["A1"].comment)

    def test_barcode_code128_and_ean13(self):
        f = self.book([["码"], ["ABC-123"], ["6901234567892"]])
        res, _ = run("excel", "barcode_gen", [f], self.out,
                     {"type": "code128", "col": "A", "insert": False})
        self.assertEqual(res["fail_count"], 0)
        pngs = [p for p in res["output_files"] if p.endswith(".png")] or \
               [os.path.join(res["output_files"][0], n)
                for n in os.listdir(res["output_files"][0]) if n.endswith(".png")]
        self.assertTrue(pngs)

        res, _ = run("excel", "barcode_gen", [f], self.out,
                     {"type": "ean13", "col": "A", "insert": True})
        self.assertEqual(res["fail_count"], 0)
        self.assertTrue(res["output_files"][0].endswith(".xlsx"))


class TestConvertFormats(Base):
    def test_json_both_ways(self):
        jp = os.path.join(self.tmp, "x.json")
        with open(jp, "w", encoding="utf-8") as fh:
            json.dump([{"名称": "甲", "值": 1}, {"名称": "乙", "值": 2}], fh, ensure_ascii=False)
        res, _ = run("excel", "json_ops", [jp], self.out, {"mode": "json2excel"})
        ws = openpyxl.load_workbook(res["output_files"][0]).active
        self.assertEqual(ws.cell(row=2, column=1).value, "甲")

        res, _ = run("excel", "json_ops", [res["output_files"][0]], self.out,
                     {"mode": "excel2json"})
        data = json.loads(open(res["output_files"][0], encoding="utf-8-sig").read())
        self.assertEqual(data[0]["名称"], "甲")

    def test_textfile_roundtrip(self):
        tp = os.path.join(self.tmp, "t.txt")
        with open(tp, "w", encoding="utf-8") as fh:
            fh.write("姓名,值\n甲,1\n乙,2\n")
        res, _ = run("excel", "textfile_ops", [tp], self.out, {"mode": "txt2excel"})
        ws = openpyxl.load_workbook(res["output_files"][0]).active
        self.assertEqual(ws.cell(row=1, column=1).value, "姓名")
        self.assertEqual(ws.cell(row=3, column=2).value, "2")

        res, _ = run("excel", "textfile_ops", [res["output_files"][0]], self.out,
                     {"mode": "excel2txt"})
        text = open(res["output_files"][0], encoding="utf-8-sig").read()
        self.assertIn("甲", text)

    def test_split_text(self):
        tp = os.path.join(self.tmp, "big.txt")
        with open(tp, "w", encoding="utf-8") as fh:
            fh.write("\n".join(str(i) for i in range(25)))
        res, _ = run("excel", "textfile_ops", [tp], self.out,
                     {"mode": "split", "lines": 10})
        self.assertEqual(len(res["output_files"]), 3)

    def test_vcf(self):
        f = self.book([["姓名", "手机", "邮箱"], ["甲", "13800000000", "a@b.com"]])
        res, _ = run("excel", "vcf_dbf", [f], self.out,
                     {"mode": "vcf", "name_col": "A", "phone_col": "B", "email_col": "C"})
        text = open(res["output_files"][0], encoding="utf-8").read()
        self.assertIn("BEGIN:VCARD", text)
        self.assertIn("TEL;TYPE=CELL:13800000000", text)
        self.assertIn("FN:甲", text)

    def test_dbf_readable(self):
        f = self.book([["姓名", "金额"], ["甲", 12.5], ["乙", 30]])
        res, _ = run("excel", "vcf_dbf", [f], self.out, {"mode": "dbf"})
        raw = open(res["output_files"][0], "rb").read()
        self.assertEqual(raw[0], 0x03)
        nrec = int.from_bytes(raw[4:8], "little")
        hlen = int.from_bytes(raw[8:10], "little")
        rlen = int.from_bytes(raw[10:12], "little")
        self.assertEqual(nrec, 2)
        self.assertEqual(raw[hlen - 1], 0x0D)
        self.assertEqual(len(raw), hlen + nrec * rlen + 1)
        first = raw[hlen:hlen + rlen]
        self.assertEqual(first[0:1], b" ")
        self.assertIn("甲".encode("gb18030"), first)

    def test_pdf_split_merge(self):
        from PIL import Image
        imgs = []
        for i, color in enumerate([(255, 0, 0), (0, 0, 255)]):
            p = os.path.join(self.tmp, "p%d.png" % i)
            Image.new("RGB", (60, 40), color).save(p)
            imgs.append(p)
        res, _ = run("excel", "img_ops", imgs, self.out, {"mode": "to_pdf"})
        pdf = res["output_files"][0]

        res, _ = run("excel", "pdf_ops", [pdf], self.out, {"mode": "split", "pages": 1})
        self.assertEqual(len(res["output_files"]), 2)

        res, _ = run("excel", "pdf_ops", res["output_files"], self.out, {"mode": "merge"})
        self.assertTrue(os.path.getsize(res["output_files"][0]) > 100)

    def test_img_resize_and_dominant(self):
        from PIL import Image
        p = os.path.join(self.tmp, "a.png")
        Image.new("RGB", (200, 100), (10, 20, 30)).save(p)
        res, _ = run("excel", "img_ops", [p], self.out, {"mode": "resize", "width": 100})
        self.assertEqual(Image.open(res["output_files"][0]).size[0], 100)

        res, _ = run("excel", "img_ops", [p], self.out, {"mode": "dominant"})
        text = open(res["output_files"][0], encoding="utf-8-sig").read()
        self.assertIn("#0A141E", text)


class TestIdsAndCommon(Base):
    def test_idcard_and_mask(self):
        f = self.book([["身份证", "手机"], ["11010519491231002X", "13812345678"]])
        res, _ = run("excel", "name_data_ops", [f], self.out,
                     {"mode": "idcard", "col": "A"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=3).value, "1949-12-31")
        self.assertEqual(ws.cell(row=2, column=4).value, "女")
        self.assertEqual(ws.cell(row=2, column=6).value, "正确")

        res, _ = run("excel", "name_data_ops", [f], self.out,
                     {"mode": "mask", "col": "B", "kind": "phone"})
        ws = openpyxl.load_workbook(res["output_files"][0])["数据"]
        self.assertEqual(ws.cell(row=2, column=3).value, "138****5678")

    def test_credit_code(self):
        from tools.excel.name_data_ops import credit_check
        self.assertTrue(credit_check("91350100M000100Y43"))
        self.assertFalse(credit_check("91350100M000100Y44"))

    def test_dup_file(self):
        a = os.path.join(self.tmp, "a.txt")
        b = os.path.join(self.tmp, "b.txt")
        with open(a, "w", encoding="utf-8") as fh:
            fh.write("same")
        shutil.copy(a, b)
        res, _ = run("common", "dup_file", [self.tmp], self.out, {"by": "hash"})
        text = open(res["output_files"][0], encoding="utf-8-sig").read()
        self.assertEqual(res["success_count"], 1)
        self.assertIn("b.txt", text)

    def test_file_list(self):
        with open(os.path.join(self.tmp, "z.txt"), "w", encoding="utf-8") as fh:
            fh.write("x")
        res, _ = run("common", "file_list", [self.tmp], self.out, {})
        rows = list(csv.reader(open(res["output_files"][0], encoding="utf-8-sig")))
        self.assertEqual(rows[0][0], "序号")
        self.assertTrue(any("z.txt" in r for r in rows[1:]))

    def test_file_list_rejects_empty_result(self):
        """输入不可用时不能静默产出只有表头的空 CSV。"""
        from tools.common.file_list import run_task

        class Log:
            lines = []

        class Ctx:
            task = {"log": Log()}

            def info(self, m):
                self.task["log"].lines.append(m)

            def warn(self, m):
                self.task["log"].lines.append(m)

        missing = os.path.join(self.tmp, "根本没有这个目录")
        with self.assertRaises(ValueError) as cm:
            run_task([missing], self.out, {}, Ctx())
        self.assertIn("没有可列出的内容", str(cm.exception))
        with self.assertRaises(ValueError):
            run_task([], self.out, {}, Ctx())

    def test_mkdir_batch(self):
        res, _ = run("common", "mkdir_batch", [], self.out,
                     {"names": "甲,乙,丙", "numbered": True})
        self.assertEqual(res["success_count"], 3)
        self.assertTrue(all(os.path.isdir(p) for p in res["output_files"]))

    def test_encoding_convert(self):
        p = os.path.join(self.tmp, "g.txt")
        with open(p, "w", encoding="gb18030") as fh:
            fh.write("中文内容")
        res, _ = run("common", "encoding_ops", [p], self.out, {"to": "utf-8"})
        self.assertEqual(res["success_count"], 1)
        with open(res["output_files"][0], encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "中文内容")


class TestWordNew(Base):
    def doc(self):
        import docx
        d = docx.Document()
        d.add_heading("第一章", 1)
        d.add_paragraph("正文内容")
        p = os.path.join(self.tmp, "w.docx")
        d.save(p)
        return p

    def test_watermark(self):
        f = self.doc()
        res, _ = run("word", "doc_watermark", [f], self.out, {"text": "保密"})
        import docx
        d = docx.Document(res["output_files"][0])
        self.assertIn("textpath", d.sections[0].header._element.xml)
        self.assertIn("保密", d.sections[0].header._element.xml)

    def test_toc(self):
        f = self.doc()
        res, _ = run("word", "doc_toc", [f], self.out, {"title": "目录"})
        import docx
        d = docx.Document(res["output_files"][0])
        self.assertEqual(d.paragraphs[0].text, "目录")
        self.assertIn("TOC", d.element.body.xml)

    def test_stat(self):
        f = self.doc()
        res, _ = run("word", "doc_stat", [f], self.out, {})
        rows = list(csv.reader(open(res["output_files"][0], encoding="utf-8-sig")))
        self.assertEqual(rows[0][0], "文件名")
        self.assertGreater(int(rows[1][1]), 0)


class TestCatalogAndDeps(unittest.TestCase):
    def test_every_tool_has_category(self):
        from service import catalog, task_service
        tl = task_service.tool_list()
        self.assertGreater(len(tl), 85)
        no_cat = [t["id"] for t in tl if not t.get("cat") or t["cat"] == "其他"]
        self.assertEqual(no_cat, [])

    def test_deps_report(self):
        from utils import deps
        rep = {d["module"]: d["ok"] for d in deps.report()}
        self.assertTrue(rep["openpyxl"])
        self.assertTrue(rep["PIL"])
        self.assertTrue(rep["qrcode"])


if __name__ == "__main__":
    unittest.main()
