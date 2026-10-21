#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Designed by TiantianYe
# 适配指南：读取 plan/slides.json + extract/，产出 pptx。
# 环境变量：PPT_OUT（输出路径）、PPT_FONT（交付字体，默认微软雅黑；渲染验证可用本机有中文字体）、
#           PPT_TITLE（封面大标题，默认取 Word 文档标题块）、PPT_FOOTER（每页页脚署名，默认 Oli教室专用）、
#           PPT_DECK_NAME（默认输出文件名）。
"""构建教学 PPT：
- 65 页：封面 / 8 章节页 / 内容页（含表格、题图）
- 全部文本 run 原样搬运；477 个公式以 OMML 原生嵌入（a14:m），字号/颜色与所在段落一致
- 每页 fade 250ms 过渡；内容按拍点击渐显（题干一拍，答案/解析逐段一拍）
- 每页页脚 Oli教室专用 + 页码
"""
import copy, json, os, re, sys
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))
import plan_slides as PL  # 复用估算函数与内容（import 时会重跑计划，幂等）

BLOCKS = PL.BLOCKS
PLAN = PL.FINAL_SLIDES if hasattr(PL, "FINAL_SLIDES") else json.load(
    open(os.path.join(ROOT, "plan", "slides.json"), encoding="utf-8"))

EMU_IN = 914400
SLIDE_W, SLIDE_H = 12192000, 6858000
CX, CW = Inches(0.5), Inches(12.33)          # 内容区
CTOP = Inches(1.14)
C_BOTTOM_IN = 7.00
GAP = Emu(int(PL.GAP * EMU_IN))

C_TITLE = "1A237E"
C_BODY = "1A2230"
C_BLUE = "3E6FC9"
C_RED = "D9534F"
C_GRAY = "8A93A6"
C_CAP = "5A6472"
C_BORDER = "D6DCE5"
C_HEAD_BG = "F5F7FA"
C_ALT = "F7F9FC"

FONT = os.environ.get("PPT_FONT", "微软雅黑")

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
def qn(t):
    pre, local = t.split(":")
    return "{%s}%s" % (NS[pre], local)

M_T = re.compile(r"<m:t(?:\s[^>]*)?>([^<]*)</m:t>")

# ---------- 段落样式判定 ----------
ANS_WORDS = ("答案",)
RED_LABELS = ("易错点", "易错位置")
BLUE_LABELS = ("解析", "逐步解析", "要点解析", "解题策略", "识别方法", "方法总结",
               "关键突破口", "备考含义", "通用三步法", "能力要求", "逐项拆解",
               "干扰项设计意图", "命题提示", "考查能力", "变式方向", "同类真题",
               "典型原题", "原题", "原题结构", "例", "方法要点")

def para_style(b):
    """返回 (整体色, label 运行着色函数)"""
    txt = "".join(r.get("t", "") for r in b.get("runs", []) if "t" in r).strip()
    stripped = re.sub(r"^[ \t]+", "", txt)
    if stripped.startswith("答案"):
        return "answer"
    if stripped.startswith("易错"):
        return "warn"
    return "normal"

def is_label_run(b, ri, r):
    """段首/冒号前的粗体引导词"""
    if not r.get("b") or "t" not in r:
        return False
    t = r["t"]
    if any(lb for lb in BLUE_LABELS if t.strip().startswith(lb)):
        return True
    if ri == 0 and r.get("b") and len(t.strip()) <= 20 and (
            t.strip().endswith("：") or t.strip().endswith(":")):
        return True
    return False

# ---------- OMML 处理 ----------
def math_element(xml_str, color, pt):
    """解析 docx OMML 片段 → 设置字号/颜色 → 返回元素（m:oMath 或 m:oMathPara）"""
    frag = xml_str.encode("utf-8") if isinstance(xml_str, str) else xml_str
    el = etree.fromstring(frag)
    half = str(int(round(pt * 2)))
    ns_w = NS["w"]
    for tag in ("sz", "szCs"):
        for e in el.iter("{%s}%s" % (ns_w, tag)):
            e.set("{%s}val" % ns_w, half)
    for e in el.iter("{%s}color" % ns_w):
        e.set("{%s}val" % ns_w, color)
    # 为每个 m:r 补 w:rPr（字号+颜色），m:ctrlPr 内的 w:rPr 补色
    for mr in el.iter("{%s}r" % NS["m"]):
        rpr = mr.find("{%s}rPr" % ns_w)
        if rpr is None:
            rpr = etree.SubElement(mr, "{%s}rPr" % ns_w)
            m_rpr = mr.find("{%s}rPr" % NS["m"])
            mr.remove(rpr)
            if m_rpr is not None:
                m_rpr.addnext(rpr)
            else:
                mr.insert(0, rpr)
        for tag, val in (("sz", half), ("szCs", half), ("color", color)):
            e = rpr.find("{%s}%s" % (ns_w, tag))
            if e is None:
                e = etree.SubElement(rpr, "{%s}%s" % (ns_w, tag))
            e.set("{%s}val" % ns_w, val)
    return el

def linearize_math(om):
    """更精细的线性化：分数加括号、上下标用 ^/_"""
    tag = etree.QName(om).localname if isinstance(om, etree._Element) else ""

    def lin(e):
        t = etree.QName(e).localname
        if t == "t":
            return e.text or ""
        if t in ("rPr", "ctrlPr", "fPr", "sSupPr", "sSubPr", "sSubSupPr", "radPr",
                 "dPr", "naryPr", "oMathParaPr", "argPr", "limLowPr", "limUppPr",
                 "mPr", "barPr", "boxPr", "groupChrPr", "funcPr"):
            return ""
        if t == "f":
            num = e.find("{%s}num" % NS["m"])
            den = e.find("{%s}den" % NS["m"])
            sn, sd = lin_all(num), lin_all(den)
            sn = sn if len(sn) <= 1 else "(" + sn + ")"
            sd = sd if len(sd) <= 1 else "(" + sd + ")"
            return sn + "/" + sd
        if t == "sSup":
            base = lin_all(e.find("{%s}e" % NS["m"]))
            sup = lin_all(e.find("{%s}sup" % NS["m"]))
            return base + "^" + (sup if len(sup) <= 1 else "(" + sup + ")")
        if t == "sSub":
            base = lin_all(e.find("{%s}e" % NS["m"]))
            sub = lin_all(e.find("{%s}sub" % NS["m"]))
            return base + "_" + (sub if len(sub) <= 1 else "(" + sub + ")")
        if t == "sSubSup":
            base = lin_all(e.find("{%s}e" % NS["m"]))
            sub = lin_all(e.find("{%s}sub" % NS["m"]))
            sup = lin_all(e.find("{%s}sup" % NS["m"]))
            return base + "_" + (sub if len(sub) <= 1 else "(" + sub + ")") + "^" + \
                (sup if len(sup) <= 1 else "(" + sup + ")")
        if t == "rad":
            e_ = e.find("{%s}e" % NS["m"])
            return "√(" + lin_all(e_) + ")"
        if t == "d":
            dpr = e.find("{%s}dPr" % NS["m"])
            beg, end = "(", ")"
            if dpr is not None:
                b = dpr.find("{%s}begChr" % NS["m"])
                c = dpr.find("{%s}endChr" % NS["m"])
                if b is not None:
                    beg = b.get("{%s}val" % NS["m"], "(")
                if c is not None:
                    end = c.get("{%s}val" % NS["m"], ")")
            inner = "".join(lin_all(x) for x in e.findall("{%s}e" % NS["m"]))
            return beg + inner + end
        if t == "nary":
            chr_el = e.find(".//{%s}chr" % NS["m"])
            sym = chr_el.get("{%s}val" % NS["m"], "∑") if chr_el is not None else "∫"
            sub = e.find("{%s}sub" % NS["m"])
            sup = e.find("{%s}sup" % NS["m"])
            e_ = e.find("{%s}e" % NS["m"])
            s = sym
            if sub is not None:
                s += "_" + lin_all(sub)
            if sup is not None:
                s += "^" + lin_all(sup)
            return s + lin_all(e_)
        if t in ("e", "num", "den", "sup", "sub", "lim", "func", "oMath",
                 "oMathPara", "mr"):
            return "".join(lin(x) for x in e)
        return "".join(lin(x) for x in e)

    def lin_all(e):
        if e is None:
            return ""
        return "".join(lin(x) for x in e)

    return lin(om)

def append_run_xml(p_el, runs, pt, color_default, bold_default=False,
                   allow_color=True):
    """把 docx run 列表写入 a:p（文本/换行/公式，公式=AC(OMML+线性兜底)）"""
    for ri, r in enumerate(runs):
        if "br" in r:
            br = etree.SubElement(p_el, qn("a:br"))
            continue
        if "math" in r:
            el = math_element(r["math"], color_default if allow_color else C_BODY, pt)
            om = el
            if etree.QName(el).localname == "oMathPara":
                found = el.find("{%s}oMath" % NS["m"])
                om = found if found is not None else el[0]
            ac = etree.SubElement(p_el, qn("mc:AlternateContent"),
                                  nsmap={"mc": NS["mc"]})
            # a14/m 命名空间必须声明在 mc:Choice 上（MCE 规范要求 Requires 前缀
            # 在 Choice 作用域内可解析），否则严格解析器会回落到 Fallback 文本
            choice = etree.SubElement(ac, qn("mc:Choice"),
                                      nsmap={"a14": NS["a14"], "m": NS["m"]})
            choice.set("Requires", "a14")
            a14m = etree.SubElement(choice, qn("a14:m"))
            if r.get("disp"):
                omp = etree.Element("{%s}oMathPara" % NS["m"])
                omp.append(copy.deepcopy(om))
                a14m.append(omp)
            else:
                a14m.append(copy.deepcopy(om))
            fb = etree.SubElement(ac, qn("mc:Fallback"))
            fb_r = etree.SubElement(fb, qn("a:r"))
            fb_rpr = etree.SubElement(fb_r, qn("a:rPr"))
            fb_rpr.set("lang", "zh-CN")
            fb_rpr.set("sz", str(int(pt * 100)))
            if bold_default:
                fb_rpr.set("b", "1")
            fill = etree.SubElement(fb_rpr, qn("a:solidFill"))
            clr = etree.SubElement(fill, qn("a:srgbClr"))
            clr.set("val", color_default)
            latin = etree.SubElement(fb_rpr, qn("a:latin"))
            latin.set("typeface", FONT)
            ea = etree.SubElement(fb_rpr, qn("a:ea"))
            ea.set("typeface", FONT)
            ft = etree.SubElement(fb_r, qn("a:t"))
            ft.text = linearize_math(om)
            continue
        if "img" in r:
            continue  # 图片由布局层处理
        t = r.get("t", "")
        if not t:
            continue
        t = t.replace("\t", "    ")
        color = color_default
        bold = bold_default or bool(r.get("b"))
        r_el = etree.SubElement(p_el, qn("a:r"))
        rpr = etree.SubElement(r_el, qn("a:rPr"))
        rpr.set("lang", "zh-CN")
        rpr.set("sz", str(int(pt * 100)))
        if bold:
            rpr.set("b", "1")
        fill = etree.SubElement(rpr, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", color)
        latin = etree.SubElement(rpr, qn("a:latin"))
        latin.set("typeface", FONT)
        ea = etree.SubElement(rpr, qn("a:ea"))
        ea.set("typeface", FONT)
        t_el = etree.SubElement(r_el, qn("a:t"))
        t_el.text = t

def set_ppr(p_el, pt, algn=None, bullet=False, spc_aft=3):
    ppr = p_el.find(qn("a:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("a:pPr"))
        p_el.insert(0, ppr)
    if algn:
        ppr.set("algn", algn)
    if bullet:
        ppr.set("marL", "342900")
        ppr.set("indent", "-342900")
    # 子元素顺序（schema）：lnSpc, spcBef, spcAft, buFont, buChar
    ln = etree.Element(qn("a:lnSpc"))
    pct = etree.SubElement(ln, qn("a:spcPct"))
    pct.set("val", "115000")
    ppr.append(ln)
    spc = etree.Element(qn("a:spcAft"))
    pts = etree.SubElement(spc, qn("a:spcPts"))
    pts.set("val", str(int(spc_aft * 100)))
    ppr.append(spc)
    if bullet:
        bf = etree.Element(qn("a:buFont"))
        bf.set("typeface", "Arial")
        ppr.append(bf)
        bc = etree.Element(qn("a:buChar"))
        bc.set("char", "•")
        ppr.append(bc)

def add_paragraph(txBody, b, pt, style_mode="normal"):
    """在 txBody 中添加一个段落"""
    p_el = etree.SubElement(txBody, qn("a:p"))
    st = para_style(b)
    bullet = b.get("num") is not None
    math_only = (any("math" in r for r in b.get("runs", []))
                 and not any("t" in r and r["t"].strip() for r in b.get("runs", [])))
    if st == "answer":
        set_ppr(p_el, pt, algn="ctr" if math_only else None, bullet=bullet)
        append_run_xml(p_el, b["runs"], pt, C_BLUE, bold_default=True)
        return p_el
    if math_only:
        set_ppr(p_el, pt, algn="ctr")
        append_run_xml(p_el, b["runs"], pt, C_BODY)
        return p_el
    if st == "warn":
        set_ppr(p_el, pt, bullet=bullet)
        _append_runs_mixed(p_el, b["runs"], pt, C_BODY, red_first_label=True)
        return p_el
    # normal：引导词蓝、正文深色；板块头粗体段蓝、其余正文黑
    txt = "".join(r.get("t", "") for r in b.get("runs", []) if "t" in r).strip()
    if txt.startswith("变式"):
        set_ppr(p_el, pt, bullet=bullet)
        append_run_xml(p_el, b["runs"], pt, C_BLUE, bold_default=True)
        return p_el
    if txt.startswith("板块"):
        # 仅粗体引导段（"板块X XXX"）用蓝，括号说明等回归正文黑
        set_ppr(p_el, pt, bullet=bullet)
        for r in b["runs"]:
            if "br" in r:
                etree.SubElement(p_el, qn("a:br"))
                continue
            if "math" in r:
                append_run_xml(p_el, [r], pt, C_BODY)
                continue
            if "t" not in r or not r.get("t"):
                continue
            if r.get("b"):
                append_run_xml(p_el, [r], pt, C_BLUE)
            else:
                append_run_xml(p_el, [{"t": r["t"]}], pt, C_BODY)
        return p_el
    set_ppr(p_el, pt, bullet=bullet)
    _append_runs_mixed(p_el, b["runs"], pt, C_BODY, red_first_label=False)
    return p_el

def _append_runs_mixed(p_el, runs, pt, body_color, red_first_label=False):
    """引导词（粗体标签到冒号）着蓝/红，其余正文色，粗体保留"""
    phase = "label" if red_first_label or True else "body"
    label_color = C_RED if red_first_label else C_BLUE
    seen_text = False
    in_label = True
    for r in runs:
        if "br" in r:
            etree.SubElement(p_el, qn("a:br"))
            continue
        if "math" in r:
            color = label_color if (in_label and not seen_text) else body_color
            append_run_xml(p_el, [r], pt, color)
            seen_text = True
            continue
        if "img" in r or "t" not in r:
            continue
        t = r["t"]
        if not t.strip():
            if not seen_text and in_label:
                append_run_xml(p_el, [r], pt, label_color)
            else:
                append_run_xml(p_el, [r], pt, body_color)
            continue
        if in_label and r.get("b"):
            # 引导词 run；找到冒号结束
            idx = max(t.find("："), t.find(":"))
            if idx >= 0:
                head, tail = t[:idx + 1], t[idx + 1:]
                append_run_xml(p_el, [{"t": head, "b": True}], pt, label_color)
                if tail:
                    append_run_xml(p_el, [{"t": tail, "b": r.get("b")}], pt, body_color)
                in_label = False
                seen_text = True
            else:
                # 整 run 都是标签的一部分？若 run 很长且不以冒号结尾，视为正文粗体
                if len(t.strip()) <= 24:
                    append_run_xml(p_el, [r], pt, label_color)
                else:
                    append_run_xml(p_el, [r], pt, body_color)
                seen_text = True
        else:
            in_label = False
            append_run_xml(p_el, [r], pt, body_color)
            seen_text = True

# ---------- 形状工具 ----------
def add_textbox(slide, x, y, w, h):
    from pptx.enum.text import MSO_AUTO_SIZE
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    for k, v in (("lIns", "54864"), ("rIns", "54864"), ("tIns", "27432"), ("bIns", "27432")):
        bodyPr.set(k, v)
    # 移除默认空段落（避免渲染时多占一行）
    for pel in tf._txBody.findall(qn("a:p")):
        tf._txBody.remove(pel)
    return tb

def shape_id(shape):
    return shape._element.nvSpPr.cNvPr.get("id") if hasattr(shape._element, "nvSpPr") \
        else shape._element.find(".//" + qn("p:cNvPr")).get("id")

def all_shape_ids(shape):
    el = shape._element
    return [e.get("id") for e in el.iter(qn("p:cNvPr"))][:1]

def set_cell_border(cell, color=C_BORDER, w_emu="9525"):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = tcPr.find(qn(tag))
        if ln is not None:
            tcPr.remove(ln)
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, qn(tag))
        ln.set("w", w_emu)
        ln.set("cap", "flat")
        fill = etree.SubElement(ln, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", color)
    # ln 元素顺序要求：lnL,lnR,lnT,lnB 需在 tcPr 子元素顺序最前 —— python-pptx 的
    # get_or_add_tcPr 保证插入位；此处 SubElement 追加在末尾，需移动到最前
    lns = [tcPr.find(qn(t)) for t in ("a:lnL", "a:lnR", "a:lnT", "a:lnB")]
    for ln in reversed(lns):
        if ln is not None:
            tcPr.insert(0, ln)

def fill_cell(cell, paras, pt, bold_force=None, align=None):
    txBody = cell._tc.find(qn("a:txBody"))
    # 清空默认段落
    for pel in txBody.findall(qn("a:p")):
        txBody.remove(pel)
    for b in paras:
        runs = [r for r in b.get("runs", [])]
        p_el = etree.SubElement(txBody, qn("a:p"))
        ppr = etree.Element(qn("a:pPr"))
        if align:
            ppr.set("algn", align)
        p_el.insert(0, ppr)
        if not any(("t" in r) or ("math" in r) for r in runs):
            r_el = etree.SubElement(p_el, qn("a:r"))
            rpr = etree.SubElement(r_el, qn("a:rPr"))
            rpr.set("lang", "zh-CN")
            rpr.set("sz", str(int(pt * 100)))
            t_el = etree.SubElement(r_el, qn("a:t"))
            t_el.text = " "
            continue
        for r in runs:
            if "math" in r:
                append_run_xml(p_el, [r], pt, C_BODY)
                continue
            if "t" not in r:
                continue
            t = r.get("t", "").replace("\t", "    ")
            if not t:
                continue
            bold = bool(r.get("b")) if bold_force is None else bold_force
            append_run_xml(p_el, [{"t": t, "b": bold}], pt, C_BODY)

def build_table(slide, tbl_block, top, font_pt, widths=None, rows_subset=None):
    tbl = tbl_block if rows_subset is None else {"rows": rows_subset}
    if widths is None:
        widths = PL.table_col_widths(tbl_block, font_pt)
    nrow = len(tbl["rows"])
    ncol = len(tbl["rows"][0]["cells"])
    gf = slide.shapes.add_table(nrow, ncol, CX, top, Inches(12.33), Emu(1))
    table = gf.table
    # 列宽
    for ci, wd in enumerate(widths):
        table.columns[ci].width = Emu(int(wd * EMU_IN))
    # 行高（最小值）
    for ri, row in enumerate(tbl["rows"]):
        lines = 1
        for ci, c in enumerate(row["cells"]):
            t = "".join(r.get("t", "") for p in c["paras"] for r in p.get("runs", []) if "t" in r)
            ml = sum(len("".join(M_T.findall(r["math"]))) * 0.62 + 1.8
                     for p in c["paras"] for r in p.get("runs", []) if "math" in r)
            wlen = sum(1.0 if ord(ch) > 0x2E80 else 0.56 for ch in t) + ml
            if wlen:
                cpl = max(4.0, (widths[ci] - 0.12) * 72 / (font_pt * 1.04))
                lines = max(lines, -(-int(wlen * 100) // int(cpl * 100)))
        table.rows[ri].height = Emu(int((lines * font_pt * 1.30 / 72 + 0.09) * EMU_IN))
    header = rows_subset is not None
    for ri, row in enumerate(tbl["rows"]):
        for ci, c in enumerate(row["cells"]):
            cell = table.cell(ri, ci)
            fill_color = C_HEAD_BG if ri == 0 else (C_ALT if ri % 2 == 0 else "FFFFFF")
            cell.fill.solid()
            from pptx.dml.color import RGBColor
            cell.fill.fore_color.rgb = RGBColor.from_string(fill_color)
            cell.vertical_anchor = 3  # MSO_ANCHOR.MIDDLE? via enum below
            from pptx.enum.text import MSO_ANCHOR
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Emu(int(0.06 * EMU_IN))
            cell.margin_right = Emu(int(0.06 * EMU_IN))
            cell.margin_top = Emu(int(0.02 * EMU_IN))
            cell.margin_bottom = Emu(int(0.02 * EMU_IN))
            set_cell_border(cell)
            fill_cell(cell, c["paras"], font_pt, bold_force=True if ri == 0 else None)
    gf.height = Emu(sum(r.height for r in table.rows))
    return gf

def add_picture_block(slide, rid, x, y):
    w_in, h_in, target = PL.img_display_size(rid)
    path = os.path.join(ROOT, "extract", target)
    pic = slide.shapes.add_picture(path, Emu(int(x)), Emu(int(y)),
                                   Emu(int(w_in * EMU_IN)), Emu(int(h_in * EMU_IN)))
    return pic, w_in, h_in

# ---------- 过渡与时序 ----------
TRANS_XML = (
    '<mc:AlternateContent xmlns:mc="%(mc)s" xmlns:p="%(p)s" xmlns:p14="%(p14)s">'
    '<mc:Choice Requires="p14">'
    '<p:transition spd="med" p14:dur="250"><p:fade/></p:transition>'
    '</mc:Choice><mc:Fallback>'
    '<p:transition spd="med"><p:fade/></p:transition>'
    '</mc:Fallback></mc:AlternateContent>'
) % NS

def add_transition(slide):
    frag = etree.fromstring(TRANS_XML)
    slide._element.append(frag)

EFFECT = """<p:par xmlns:p="{p}"><p:cTn id="{i1}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{i3}" dur="300"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>"""

def build_timing_xml(beats):
    """beats: [[spid,...], ...] 每拍一次点击"""
    nid = [3]
    parts = []
    for spids in beats:
        effects = []
        for k, spid in enumerate(spids):
            node = "clickEffect" if k == 0 else "withEffect"
            effects.append(EFFECT.format(p=NS["p"], i1=nid[0], i2=nid[0] + 1,
                                         i3=nid[0] + 2, spid=spid, node=node))
            nid[0] += 3
        parts.append(
            '<p:par xmlns:p="{p}"><p:cTn id="{i1}" fill="hold">'
            '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
            '<p:par xmlns:p="{p}"><p:cTn id="{i2}" fill="hold">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            '<p:childTnLst>{effects}</p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn></p:par>'.format(
                p=NS["p"], i1=nid[0], i2=nid[0] + 1, effects="".join(effects)))
        nid[0] += 2
    xml = (
        '<p:timing xmlns:p="{p}" xmlns:a="{a}"><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        '<p:childTnLst>{beats}</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    ).format(p=NS["p"], a=NS["a"], beats="".join(parts))
    return xml

def add_timing(slide, beats):
    if not beats:
        return
    frag = etree.fromstring(build_timing_xml(beats))
    slide._element.append(frag)

# ---------- 页面构建 ----------
def add_footer(slide, page_no, total):
    tb = add_textbox(slide, Inches(9.53), Inches(7.06), Inches(3.3), Inches(0.32))
    txBody = tb.text_frame._txBody
    p_el = etree.SubElement(txBody, qn("a:p"))
    ppr = etree.Element(qn("a:pPr"))
    ppr.set("algn", "r")
    p_el.insert(0, ppr)
    footer = os.environ.get("PPT_FOOTER", "Oli教室专用")
    append_run_xml(p_el, [{"t": "%s · %d/%d" % (footer, page_no, total)}], 10, C_GRAY)

def title_pt_for(title_text):
    wlen = sum(1.0 if ord(c) > 0x2E80 else 0.56 for c in title_text)
    if wlen * 22 <= 830:
        return 22
    if wlen * 20 <= 830:
        return 20
    return 18

def add_title(slide, title_text, pt=None):
    if pt is None:
        pt = title_pt_for(title_text)
    bar = slide.shapes.add_shape(1, Inches(0.5), Inches(0.38), Inches(0.09), Inches(0.5))
    bar.fill.solid()
    from pptx.dml.color import RGBColor
    bar.fill.fore_color.rgb = RGBColor.from_string(C_BLUE)
    bar.line.fill.background()
    bar.shadow.inherit = False
    tb = add_textbox(slide, Inches(0.72), Inches(0.26), Inches(12.1), Inches(0.74))
    txBody = tb.text_frame._txBody
    p_el = etree.SubElement(txBody, qn("a:p"))
    set_ppr(p_el, pt, spc_aft=0)
    append_run_xml(p_el, [{"t": title_text, "b": True}], pt, C_TITLE)
    ln = slide.shapes.add_shape(1, Inches(0.5), Inches(1.02), Inches(12.33), Pt(1))
    ln.fill.solid()
    ln.fill.fore_color.rgb = RGBColor.from_string(C_BORDER)
    ln.line.fill.background()
    ln.shadow.inherit = False

def add_block_shape(slide, b, x, y, w, pt):
    """普通文本块 → textbox；含图片的块 → (pic) 处理后返回高度"""
    has_img = any("img" in r for r in b.get("runs", []))
    if has_img:
        rid = [r["img"]["rid"] for r in b["runs"] if "img" in r][0]
        w_in, h_in, target = PL.img_display_size(rid)
        px = x + (Inches(12.33) - int(w_in * EMU_IN)) // 2
        pic = slide.shapes.add_picture(os.path.join(ROOT, "extract", target),
                                       Emu(int(px)), Emu(int(y)),
                                       Emu(int(w_in * EMU_IN)), Emu(int(h_in * EMU_IN)))
        return pic, h_in
    h = PL.est_block_h(b, pt, pt)
    tb = add_textbox(slide, x, Emu(int(y)), w, Emu(int(max(h, 0.3) * EMU_IN)))
    txBody = tb.text_frame._txBody
    add_paragraph(txBody, b, pt)
    return tb, h

def build_content_slide(slide, s, page_no, total):
    sc = s.get("scale", 1.0)
    pt = PL.BODY_FONT * sc
    tpt = PL.TABLE_FONT * sc
    add_title(slide, s["title"])
    add_footer(slide, page_no, total)
    y = CTOP
    beats = []
    group_gap = Emu(int(PL.GROUP_GAP * EMU_IN))
    unit_gap = Emu(int(PL.GAP * EMU_IN))
    first_group = True
    for g in s["groups"]:
        if not first_group:
            y += group_gap - unit_gap   # 同页不同题/组之间的空行
        first_group = False
        for u in g["units"]:
            shapes = []
            if u["kind"] == "stem":
                for bi in u["blocks"]:
                    b = BLOCKS[bi]
                    sh, h = add_block_shape(slide, b, CX, y, Inches(12.33), pt)
                    shapes.append(sh)
                    y += int(h * EMU_IN) + int(0.06 * EMU_IN)
                y -= int(0.06 * EMU_IN)
            elif u["kind"] == "imgpair":
                n = len(u["imgs"])
                slot_w = 12.33 / n
                box_w, box_h = slot_w - 0.35, 2.2
                placements = []
                row_h = 0.0
                for im in u["imgs"]:
                    w_in, h_in, target = PL.img_display_size_in(im["rid"], box_w, box_h)
                    placements.append((w_in, h_in, target))
                    row_h = max(row_h, h_in)
                for k, (w_in, h_in, target) in enumerate(placements):
                    slot_x = CX + Inches(slot_w * k)
                    px = slot_x + (Inches(slot_w) - int(w_in * EMU_IN)) // 2
                    pic = slide.shapes.add_picture(
                        os.path.join(ROOT, "extract", target),
                        Emu(int(px)), y,
                        Emu(int(w_in * EMU_IN)), Emu(int(h_in * EMU_IN)))
                    shapes.append(pic)
                y += int(row_h * EMU_IN) + int(0.04 * EMU_IN)
                for k, im in enumerate(u["imgs"]):
                    if im.get("caption"):
                        slot_x = CX + Inches(slot_w * k)
                        cap = add_textbox(slide, slot_x, y, Inches(slot_w), Inches(0.3))
                        txBody = cap.text_frame._txBody
                        p_el = etree.SubElement(txBody, qn("a:p"))
                        set_ppr(p_el, 11, algn="ctr", spc_aft=0)
                        append_run_xml(p_el, [{"t": im["caption"]}], 11, C_CAP)
                        shapes.append(cap)
                y += int(0.32 * EMU_IN) + int(0.08 * EMU_IN)
            elif u["kind"] == "image":
                rid = u["img_rid"]
                w_in, h_in, target = PL.img_display_size(rid)
                px = CX + (Inches(12.33) - int(w_in * EMU_IN)) // 2
                pic = slide.shapes.add_picture(os.path.join(ROOT, "extract", target),
                                               Emu(int(px)), y,
                                               Emu(int(w_in * EMU_IN)), Emu(int(h_in * EMU_IN)))
                shapes.append(pic)
                y += int(h_in * EMU_IN) + int(0.05 * EMU_IN)
                if u.get("caption"):
                    cap = add_textbox(slide, CX, y, Inches(12.33), Inches(0.34))
                    txBody = cap.text_frame._txBody
                    p_el = etree.SubElement(txBody, qn("a:p"))
                    set_ppr(p_el, 12, algn="ctr", spc_aft=0)
                    append_run_xml(p_el, [{"t": u["caption"]}], 12, C_CAP)
                    shapes.append(cap)
                    y += int(0.36 * EMU_IN)
                y += int(0.05 * EMU_IN)
            elif u["kind"] == "table":
                gf = build_table(slide, BLOCKS[u["block"]], y, tpt)
                shapes.append(gf)
                y += int(PL.est_table_h(BLOCKS[u["block"]], tpt) * EMU_IN)
            elif u["kind"] == "tablepart":
                gf = build_table(slide, BLOCKS[u["block"]], y, tpt,
                                 widths=u.get("widths"), rows_subset=u["rows"])
                shapes.append(gf)
                y += int(PL.est_table_h({"rows": u["rows"]}, tpt, u.get("widths")) * EMU_IN)
            else:
                b = BLOCKS[u["block"]]
                sh, h = add_block_shape(slide, b, CX, y, Inches(12.33), pt)
                shapes.append(sh)
                y += int(h * EMU_IN)
            beats.append([sid for sh in shapes for sid in all_shape_ids(sh)])
            y += GAP
    add_transition(slide)
    add_timing(slide, beats)

def build_cover(slide, s, page_no, total):
    add_footer(slide, page_no, total)
    blocks = s["blocks"]
    tb = add_textbox(slide, Inches(0.9), Inches(2.42), Inches(11.53), Inches(1.5))
    p_el = etree.SubElement(tb.text_frame._txBody, qn("a:p"))
    set_ppr(p_el, 40, algn="ctr", spc_aft=0)
    main_title = os.environ.get(
        "PPT_TITLE",
        "".join(r.get("t", "") for r in BLOCKS[blocks[0]]["runs"] if "t" in r).strip())
    append_run_xml(p_el, [{"t": main_title, "b": True}], 40, C_TITLE)
    tb2 = add_textbox(slide, Inches(0.9), Inches(4.05), Inches(11.53), Inches(0.6))
    p_el = etree.SubElement(tb2.text_frame._txBody, qn("a:p"))
    set_ppr(p_el, 16, algn="ctr", spc_aft=0)
    runs = [r for r in BLOCKS[blocks[1]]["runs"] if "t" in r]
    append_run_xml(p_el, runs, 16, C_BLUE, bold_default=True)
    tb3 = add_textbox(slide, Inches(0.9), Inches(4.85), Inches(11.53), Inches(0.6))
    p_el = etree.SubElement(tb3.text_frame._txBody, qn("a:p"))
    set_ppr(p_el, 15, algn="ctr", spc_aft=0)
    runs = [r for r in BLOCKS[blocks[2]]["runs"] if "t" in r]
    append_run_xml(p_el, runs, 15, C_BODY)
    add_transition(slide)

def build_divider(slide, s, page_no, total):
    from pptx.dml.color import RGBColor
    add_footer(slide, page_no, total)
    bar = slide.shapes.add_shape(1, Inches(0.86), Inches(3.02), Inches(0.12), Inches(0.72))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor.from_string(C_BLUE)
    bar.line.fill.background()
    bar.shadow.inherit = False
    tb = add_textbox(slide, Inches(1.14), Inches(2.86), Inches(11.2), Inches(1.1))
    p_el = etree.SubElement(tb.text_frame._txBody, qn("a:p"))
    set_ppr(p_el, 30, spc_aft=0)
    append_run_xml(p_el, [{"t": s["title"], "b": True}], 30, C_TITLE)
    add_transition(slide)

def main():
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    blank = prs.slide_layouts[6]
    total = len(PLAN)
    out = os.environ.get("PPT_OUT", os.path.join(
        ROOT, "out", os.environ.get("PPT_DECK_NAME", "teaching_ppt.pptx")))
    os.makedirs(os.path.dirname(out), exist_ok=True)
    for i, s in enumerate(PLAN):
        slide = prs.slides.add_slide(blank)
        if s["type"] == "cover":
            build_cover(slide, s, i + 1, total)
        elif s["type"] == "divider":
            build_divider(slide, s, i + 1, total)
        else:
            build_content_slide(slide, s, i + 1, total)
    prs.save(out)
    print("saved:", out, "slides:", total)

if __name__ == "__main__":
    main()
