"""生成测试样例文件。运行：python samples/make_samples.py"""
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import openpyxl
import docx
from docx.shared import Pt

OUT = os.path.join(BASE, "samples", "data")
os.makedirs(OUT, exist_ok=True)


def make_xlsx(path, rows, merged=False, sheets=("数据",)):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for si, title in enumerate(sheets):
        ws = wb.create_sheet(title)
        ws.append(["编号", "部门", "姓名", "金额"])
        for r in rows:
            ws.append(list(r))
        if merged and si == 0:
            ws.merge_cells("A2:A4")
            ws["A2"] = "A组"
            ws.merge_cells("B2:B3")
            ws["B2"] = "财务"
    wb.save(path)
    return path


def make_docx(path, paras):
    d = docx.Document()
    for p in paras:
        para = d.add_paragraph()
        run = para.add_run(p)
        run.font.size = Pt(12)
    cp = d.core_properties
    cp.author = "张三"
    cp.last_modified_by = "张三"
    cp.title = "内部资料"
    d.save(path)
    return path


if __name__ == "__main__":
    make_xlsx(os.path.join(OUT, "销售表1.xlsx"),
              [(1, "一部", "甲", 100), (2, "一部", "乙", 200), (3, "二部", "丙", 300)],
              merged=True)
    make_xlsx(os.path.join(OUT, "销售表2.xlsx"),
              [(4, "二部", "丁", 400), (5, "三部", "戊", 500)],
              sheets=("Sheet1", "Sheet2"))
    make_xlsx(os.path.join(OUT, "大表.xlsx"),
              [(i, f"部门{i % 3}", f"员工{i}", i * 10) for i in range(1, 51)])
    make_docx(os.path.join(OUT, "文档1.docx"),
              ["第一段内容。", "", "", "第二段需要替换旧词。", "结尾。"])
    make_docx(os.path.join(OUT, "文档2.docx"), ["另一份文档的正文。", "第二行。"])
    print("样例已生成:", OUT)
