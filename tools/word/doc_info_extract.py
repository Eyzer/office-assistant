"""批量从人员信息文档里抓取字段，汇总成一张表。

适配四种常见排版：
  1) 单元格内「姓名：张三」
  2) 横向网格 [姓名][张三][性别][男]
  3) 纵向网格 第一行标签、第二行值
  4) 正文段落「姓名：张三  性别：男」

身份证号 / 手机号 走正则兜底，标签写得再怪也能抓到。
"""
import csv
import os
import re

from utils import doc_utils as du
from utils.file_utils import WORD_EXT, collect_files
from utils.path_utils import safe_name

META = {
    "id": "doc_info_extract",
    "cat": "提取与转换",
    "name": "人员信息提取",
    "desc": "输入 docx（表格填写的人员信息），输出一张 Excel 台账，一人一行。"
            "自动识别四种常见排版，身份证号没标签也能抓。",
    "fields": [
        {"key": "wanted", "label": "① 要提取的字段（逗号分隔）", "type": "textarea", "rows": 3,
         "default": "姓名,性别,民族,出生日期,身份证号,籍贯,政治面貌,文化程度,入伍时间,"
                    "军衔,职务,单位,联系电话,家庭住址",
         "help": "清空 = 提取全部 25 个内置字段。只用 docx，输出 xlsx。"},
        {"key": "extra", "label": "② 还想加的字段（可选，逗号分隔）", "type": "text", "default": "",
         "wide": True,
         "placeholder": "文档里有、上面没列出的标签，例：军官证号,入伍地"},
        {"key": "fmt", "label": "输出格式", "type": "select", "default": "xlsx",
         "options": [{"value": "xlsx", "label": "Excel 表格（推荐）"},
                     {"value": "csv", "label": "CSV 文本"}]},
        {"key": "per_file", "label": "另外给每份文档存一份明细", "type": "checkbox", "default": False},
    ],
}

# ---------------------------------------------------------------- 字段词典
# 标准字段 -> 可识别的标签写法（比较时会去掉空白、冒号、括号）
ALIAS = {
    "姓名": ["姓名", "名字", "姓 名", "人员姓名", "姓"],
    "性别": ["性别", "性 别"],
    "民族": ["民族", "民 族"],
    "出生日期": ["出生日期", "出生年月", "出生时间", "生日", "出生年月日"],
    "身份证号": ["身份证号", "身份证号码", "公民身份号码", "公民身份证号", "身份证",
                 "证件号码", "证件号", "身份证件号码"],
    "籍贯": ["籍贯", "籍 贯", "祖籍"],
    "政治面貌": ["政治面貌", "政治面貌情况", "党派"],
    "文化程度": ["文化程度", "学历", "文化水平", "最高学历"],
    "入伍时间": ["入伍时间", "入伍年月", "入伍日期", "参加工作时间", "参工时间", "工作时间"],
    "军衔": ["军衔", "警衔", "衔级", "军衔级别"],
    "职务": ["职务", "岗位", "职位", "担任职务", "职务名称"],
    "单位": ["单位", "工作单位", "所属单位", "所在单位", "部职别", "现单位"],
    "联系电话": ["联系电话", "电话", "手机", "手机号", "手机号码", "联系方式",
                 "电话号码", "本人电话"],
    "家庭住址": ["家庭住址", "住址", "现住址", "现居住地址", "家庭地址", "通讯地址", "地址"],
    "婚姻状况": ["婚姻状况", "婚否", "婚姻"],
    "健康状况": ["健康状况", "健康情况", "身体情况"],
    "身高": ["身高"],
    "体重": ["体重"],
    "血型": ["血型"],
    "毕业院校": ["毕业院校", "毕业学校", "院校", "毕业院校及专业"],
    "专业": ["专业", "所学专业"],
    "紧急联系人": ["紧急联系人", "联系人", "家属联系人"],
    "紧急联系电话": ["紧急联系电话", "紧急电话", "家属电话", "联系人电话"],
    "银行卡号": ["银行卡号", "工资卡号", "银行账号", "卡号"],
    "士兵证号": ["士兵证号", "军官证号", "证件编号"],
}

# 归一化后的标签 -> 标准字段
_LABEL_MAP = {}
for _std, _alist in ALIAS.items():
    for _a in _alist:
        _LABEL_MAP[re.sub(r"[\s\u3000:：()（）]", "", _a)] = _std

# ---------------------------------------------------------------- 正则兜底
RE_ID18 = re.compile(r"(?<![0-9])[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])"
                     r"(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?![0-9])")
RE_ID15 = re.compile(r"(?<![0-9])[1-9]\d{7}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])"
                     r"\d{3}(?![0-9])")
RE_PHONE = re.compile(r"(?<![0-9])1[3-9]\d{9}(?![0-9])")
RE_DATE = re.compile(r"(?<![0-9])((?:19|20)\d{2})\s*[-年./]\s*(\d{1,2})"
                     r"(?:\s*[-月./]\s*(\d{1,2}))?\s*日?(?![0-9])")


def id_checksum_ok(s):
    """18 位身份证校验位验证，防止把别的长数字当成身份证。"""
    if len(s) != 18:
        return False
    w = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
    m = "10X98765432"
    try:
        total = sum(int(s[i]) * w[i] for i in range(17))
    except ValueError:
        return False
    return m[total % 11] == s[17].upper()


def norm_label(s):
    return re.sub(r"[\s\u3000:：()（）]", "", s or "")


def clean_value(s):
    """去掉值两端的分隔符和多余空白，压掉单元格内换行。"""
    v = (s or "").replace("\r", "\n")
    v = re.sub(r"\n+", " ", v)
    v = v.strip(" \t\u3000:：,，、;；|")
    return v.strip()


def _is_label(s):
    return norm_label(s) in _LABEL_MAP


def _put(bag, label, value):
    """写入字段。同一字段已有非空值时不覆盖（先到的优先，通常是主表）。"""
    std = _LABEL_MAP.get(norm_label(label))
    if not std:
        return False
    v = clean_value(value)
    if not v or _is_label(v):
        return False
    if not bag.get(std):
        bag[std] = v
        return True
    return False


# ---------------------------------------------------------------- 抽取
SPLIT_RE = re.compile(r"[：:]")


def _scan_text(bag, text):
    """从一段文本里抓「标签：值」。支持一行里多组用空格/制表符分隔。"""
    if not text:
        return
    for seg in re.split(r"[\t\u3000]+|\s{2,}", text):
        seg = seg.strip()
        if not seg:
            continue
        if not SPLIT_RE.search(seg):
            continue
        # 「姓名：张三 性别：男」这种半角空格分隔的，先按已知标签二次切
        parts = _split_by_label(seg)
        for lab, val in parts:
            _put(bag, lab, val)


def _split_by_label(seg):
    """把「姓名：张三 性别：男」切成 [(姓名,张三),(性别,男)]。"""
    hits = []
    for key in _LABEL_MAP:
        for m in re.finditer(re.escape(key), seg):
            hits.append((m.start(), m.end(), key))
    if not hits:
        return []
    hits.sort()
    # 只保留不重叠的、出现在冒号前面的标签位置
    picked = []
    for st, en, key in hits:
        if picked and st < picked[-1][1]:
            continue
        tail = seg[en:en + 3]
        if not re.match(r"^[\s\u3000]*[：:]", tail):
            continue
        picked.append((st, en, key))
    out = []
    for i, (st, en, key) in enumerate(picked):
        nxt = picked[i + 1][0] if i + 1 < len(picked) else len(seg)
        chunk = seg[en:nxt]
        chunk = SPLIT_RE.sub("", chunk, count=1)
        out.append((key, chunk))
    return out


def _grid_of(table):
    """表格转二维网格，合并单元格只保留一次。

    keep 必须持有 tc 引用：lxml 的元素代理被回收后 id 会被复用，
    同行内就可能把两个不同单元格误判成同一个而漏掉。
    """
    rows, keep = [], []
    for row in table.rows:
        cells, seen = [], set()
        for c in row.cells:
            tc = c._tc
            keep.append(tc)
            k = id(tc)
            if k in seen:
                continue
            seen.add(k)
            cells.append(c.text)
        rows.append(cells)
    return rows


def _scan_grid(bag, rows):
    for row in rows:
        cells = [clean_value(c) for c in row]
        if not any(cells):
            continue
        # 横向：[姓名][张三][性别][男]
        i = 0
        while i < len(cells):
            if not _is_label(cells[i]):
                i += 1
                continue
            # 向右跳过空格再取值：[姓名][][张三] 这种中间夹空列的很常见
            j = i + 1
            while j < len(cells) and not cells[j]:
                j += 1
            # 右边如果还是标签，说明这行是纯表头，不能往右取值
            if j < len(cells) and not _is_label(cells[j]):
                _put(bag, cells[i], cells[j])
                i = j + 1
            else:
                i += 1
        # 单元格内冒号
        for c in cells:
            if SPLIT_RE.search(c):
                for lab, val in _split_by_label(c):
                    _put(bag, lab, val)


def _scan_vertical(bag, rows):
    """纵向布局：上一行是标签，下一行是值。"""
    for a, b in zip(rows, rows[1:]):
        la = [clean_value(x) for x in a]
        lb = [clean_value(x) for x in b]
        if len(la) != len(lb):
            continue
        # 单列竖排：[姓名]/[张三]/[性别]/[男]，一格标签一格值交替
        if len(la) == 1:
            if _is_label(la[0]) and lb[0] and not _is_label(lb[0]):
                _put(bag, la[0], lb[0])
            continue
        if len(la) < 2:
            continue
        labels = [x for x in la if _is_label(x)]
        if len(labels) < max(2, len(la) * 0.6):
            continue
        for lab, val in zip(la, lb):
            if _is_label(lab) and val and not _is_label(val):
                _put(bag, lab, val)


def _fallback_regex(bag, blob):
    """标签没命中时靠正则补。

    身份证：格式对就抓，校验位只用来标状态——真实表格里手输错一位很常见，
    严格校验会把该抓的人漏掉，宁可抓出来标「校验位错」交给人核。
    """
    # 同一行内被空格分开的数字先缝上（如「410105 19920308 8877」）。
    # 注意不能跨行缝：相邻两行的数字（身份证 + 手机号）一粘就变成 29 位，
    # 后面的边界断言 (?!\d) 直接失效，什么也抓不到。
    blob = re.sub(r"(?<=\d)[ \t\u3000]{1,3}(?=\d)", "", blob)
    if not bag.get("身份证号"):
        cands = [m.group(0).upper() for m in RE_ID18.finditer(blob)]
        good = [c for c in cands if id_checksum_ok(c)]
        if good:
            bag["身份证号"] = good[0]
        elif cands:
            bag["身份证号"] = cands[0]
        else:
            m = RE_ID15.search(blob)
            if m:
                bag["身份证号"] = m.group(0)
    if not bag.get("联系电话"):
        m = RE_PHONE.search(blob)
        if m:
            bag["联系电话"] = m.group(0)
    if not bag.get("出生日期"):
        # 身份证里的出生日期最可靠，优先用它
        idv = bag.get("身份证号") or ""
        if len(idv) >= 14 and idv[6:14].isdigit():
            y, mo, dd = idv[6:10], idv[10:12], idv[12:14]
            if 1900 < int(y) < 2100 and 1 <= int(mo) <= 12 and 1 <= int(dd) <= 31:
                bag["出生日期"] = "%s-%s-%s" % (y, mo, dd)
                return bag
        for m in RE_DATE.finditer(blob):
            y, mo, dd = m.group(1), m.group(2), m.group(3)
            if 1900 < int(y) < 2100 and 1 <= int(mo) <= 12:
                bag["出生日期"] = m.group(0).replace(" ", "")
                break
    return bag


def _id_status(v):
    if not v:
        return ""
    if len(v) == 18:
        return "校验通过" if id_checksum_ok(v) else "校验位错"
    if len(v) == 15:
        return "旧版15位"
    return "位数不对"


def _records_by_header(rows):
    """表格带表头时，按列把每行数据拆成一条记录。

    命中条件严格：一行里 ≥2 个已知标签，且标签占该行非空单元格 70% 以上。
    这样「姓名|张三|性别|男」这种值行不会被误当表头。
    """
    for hi, row in enumerate(rows):
        filled = [c for c in row if clean_value(c)]
        if len(filled) < 2:
            continue
        cols = []
        for i, c in enumerate(row):
            std = _LABEL_MAP.get(norm_label(c))
            if std:
                cols.append((i, std))
        if len(cols) < 2 or len(cols) < len(filled) * 0.7:
            continue
        recs = []
        for data_row in rows[hi + 1:]:
            if not any(clean_value(c) for c in data_row):
                continue
            bag = {}
            for i, std in cols:
                if i >= len(data_row):
                    continue
                v = clean_value(data_row[i])
                if v and not _is_label(v) and not bag.get(std):
                    bag[std] = v
            if bag:
                recs.append(bag)
        if recs:
            return recs
    return []


def extract_multi(path):
    """返回 ([{字段:值}, ...], 提示列表)。带表头的名单表会拆成多条。"""
    d = du.load(path)
    bag = {}
    blob_parts = []

    # 正文段落
    for p in d.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        blob_parts.append(t)
        _scan_text(bag, t)

    # 页眉页脚也扫一遍（有的表单位置很怪）
    for sec in d.sections:
        for hf in (sec.header, sec.footer):
            for p in hf.paragraphs:
                t = p.text.strip()
                if t:
                    blob_parts.append(t)
                    _scan_text(bag, t)

    tips = []
    for table in d.tables:
        rows = _grid_of(table)
        for row in rows:
            for c in row:
                if c and c.strip():
                    blob_parts.append(c.strip())
        # 先看是不是名单表（表头 + 多行数据）
        recs = _records_by_header(rows)
        if recs:
            for r in recs:
                # 只拿这一行自己的内容兜底，绝不能拿整篇——否则第一个人的
                # 身份证号会被正则补到所有人身上
                _fallback_regex(r, "\n".join(str(v) for v in r.values()))
                r["_id_status"] = _id_status(r.get("身份证号"))
            if len(recs) > 1:
                tips.append("按名单表头拆出 %d 人" % len(recs))
            return recs, tips
        _scan_grid(bag, rows)
        _scan_vertical(bag, rows)

    blob = "\n".join(blob_parts)
    _fallback_regex(bag, blob)

    # 一份文档里塞了多个人但没表头：按身份证号数量提醒，只取第一组（避免串行错位）
    flat = re.sub(r"(?<=\d)[ \t\u3000]{1,3}(?=\d)", "", blob)
    ids_seen = list(dict.fromkeys(m.group(0).upper() for m in RE_ID18.finditer(flat)))
    if len(ids_seen) > 1:
        tips.append("文档内有 %d 个身份证号但没找到表头，只取了第一组（%s）"
                    % (len(ids_seen), ids_seen[0]))
    if bag.get("身份证号"):
        bag["身份证号"] = re.sub(r"\s+", "", bag["身份证号"]).upper()
    bag["_id_status"] = _id_status(bag.get("身份证号"))
    return [bag], tips


def extract(path):
    """返回 {字段: 值}。第一条记录。"""
    return extract_multi(path)[0][0]


# ---------------------------------------------------------------- 主流程
def _fields_of(params):
    want = []
    for raw in (params.get("wanted") or "").split(",") + (params.get("extra") or "").split(","):
        k = norm_label(raw)
        if not k:
            continue
        std = _LABEL_MAP.get(k) or raw.strip()
        if std and std not in want:
            want.append(std)
    return want or list(ALIAS.keys())


def _write_csv(dst, header, rows):
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    with open(dst, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)
    return dst


def _write_xlsx(dst, header, rows):
    import openpyxl
    from utils import excel_utils as xu
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "人员信息"
    ws.append(header)
    for r in rows:
        ws.append(r)
    # 身份证号/电话强制文本，别被 Excel 转成科学计数
    text_cols = [i for i, h in enumerate(header, 1)
                 if h in ("身份证号", "联系电话", "银行卡号", "士兵证号")]
    for ci in text_cols:
        for ri in range(2, len(rows) + 2):
            ws.cell(row=ri, column=ci).number_format = "@"
    try:
        from openpyxl.styles import Alignment, Font, PatternFill
        head_fill = PatternFill("solid", fgColor="DCE6F1")
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = head_fill
            c.alignment = Alignment(horizontal="center", vertical="center")
        widths = {"文件名": 26, "身份证号": 22, "家庭住址": 34, "单位": 20,
                  "联系电话": 15, "出生日期": 14, "姓名": 12,
                  "身份证核验": 12, "源路径": 40}
        for i, h in enumerate(header, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = widths.get(h, 12)
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        # 核验状态着色：校验位错/位数不对标红，空标黄
        if "身份证核验" in header:
            ci = header.index("身份证核验") + 1
            for ri in range(2, len(rows) + 2):
                c = ws.cell(row=ri, column=ci)
                v = str(c.value or "")
                if v in ("校验通过", "旧版15位"):
                    c.fill = PatternFill("solid", fgColor="E2EFDA")
                elif v == "":
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
                else:
                    c.fill = PatternFill("solid", fgColor="FCE4E4")
                    c.font = Font(color="C00000")
    except Exception:
        pass
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    xu.save(wb, dst)
    return dst


def _uniq(dst):
    if not os.path.exists(dst):
        return dst
    base, ext = os.path.splitext(dst)
    i = 2
    while os.path.exists("%s_%d%s" % (base, i, ext)):
        i += 1
    return "%s_%d%s" % (base, i, ext)


def run_task(inputs, output_dir, params, ctx):
    files = collect_files(inputs, WORD_EXT)
    if not files:
        ctx.warn("没有找到 docx 文件")
        return {"success_count": 0, "fail_count": 0, "output_files": [], "log_messages": []}

    fields = _fields_of(params)
    fmt = (params.get("fmt") or "xlsx").lower()
    per_file = bool(params.get("per_file", False))

    header = ["文件名"] + fields + ["身份证核验", "源路径"]
    rows, outs = [], []
    ok, fail, empty, suspect = 0, 0, 0, 0
    total = len(files)

    for i, f in enumerate(files, 1):
        if ctx.cancelled:
            ctx.warn("已取消")
            break
        name = os.path.basename(f)
        try:
            bags, tips = extract_multi(f)
            for t in tips:
                ctx.warn("%s：%s" % (name, t))
            file_rows = []
            for bi, bag in enumerate(bags, 1):
                got = [bag.get(k, "") for k in fields]
                status = bag.get("_id_status", "")
                if not any(got):
                    empty += 1
                    ctx.warn("%s：没抓到任何字段，检查是不是扫描件/图片表格" % name)
                    continue
                if status and status not in ("校验通过", "旧版15位"):
                    suspect += 1
                    ctx.warn("%s：身份证号%s，需要人工核"
                             % (name if len(bags) == 1 else "%s第%d条" % (name, bi), status))
                # 一人一份文档时用文件名；一份文档多人时补序号，方便回溯
                label = safe_name(name) if len(bags) == 1 else "%s#%d" % (safe_name(name), bi)
                file_rows.append([label] + got + [status, f])
            rows.extend(file_rows)
            ok += len(file_rows)
            hit = [k for k in fields if bags[0].get(k)]
            ctx.info("%s：%d 条记录，命中 %d/%d 项"
                     % (name, len(file_rows), len(hit), len(fields)))
            if per_file and file_rows:
                d_dst = _uniq(os.path.join(output_dir, safe_name(os.path.splitext(name)[0]) + "_信息.csv"))
                _write_csv(d_dst, header, file_rows)
                outs.append(d_dst)
        except Exception as e:
            fail += 1
            ctx.error("%s：%s" % (name, e))
        ctx.progress(i, total, name)

    if not rows:
        ctx.error("全部失败，没有可输出内容")
        return {"success_count": 0, "fail_count": fail, "output_files": [], "log_messages": []}

    os.makedirs(output_dir, exist_ok=True)
    if fmt == "csv":
        dst = _uniq(os.path.join(output_dir, "人员信息汇总.csv"))
        _write_csv(dst, header, rows)
    else:
        dst = _uniq(os.path.join(output_dir, "人员信息汇总.xlsx"))
        _write_xlsx(dst, header, rows)
    outs.insert(0, dst)

    ctx.info("汇总 %d 份文档，失败 %d，空表 %d，身份证待核 %d → %s"
             % (ok, fail, empty, suspect, os.path.basename(dst)))
    return {"success_count": ok, "fail_count": fail,
            "log_messages": ctx.task["log"].lines[-50:], "output_files": outs}
