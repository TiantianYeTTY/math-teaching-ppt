#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Designed by TiantianYe
# 适配指南：build 之后运行，8 项硬校验全绿才允许交付。
# 环境变量与 build_pptx.py 一致（PPT_DECK_NAME/PPT_TITLE/PPT_FOOTER），确保比对口径相同。
"""全量校验：
A 每页过渡(fade)  B 每页时序且拍数=计划、spid 均存在  C 每页 Oli教室专用
D 文本逐段覆盖（与 Word 原文一致）  E OMML 数量一致  F 图片数量与宽高比
G 字体白名单  H 颜色白名单  I 内容不越界（底边 ≤ 7.03"）
"""
import json, os, re, sys, zipfile
from lxml import etree
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
PPTX = os.path.join(ROOT, "out", os.environ.get("PPT_DECK_NAME", "teaching_ppt.pptx"))

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
def qn(t):
    pre, local = t.split(":")
    return "{%s}%s" % (NS[pre], local)

D = json.load(open(os.path.join(ROOT, "extract", "content.json"), encoding="utf-8"))
BLOCKS = D["blocks"]
PLAN = json.load(open(os.path.join(ROOT, "plan", "slides.json"), encoding="utf-8"))
RID2T = D["rid2target"]

FONTS_OK = {"微软雅黑", "Arial", "Cambria Math", "+mn-lt", "+mj-lt"}
COLORS_OK = {"1A237E", "1A2230", "3E6FC9", "D9534F", "8A93A6", "5A6472",
             "D6DCE5", "F5F7FA", "F7F9FC", "FFFFFF"}

def para_text_of_block(b):
    if b["kind"] != "para":
        return None
    return "".join(r.get("t", "").replace("\t", "    ")
                   for r in b["runs"] if "t" in r).strip()

def expected_content_texts(s):
    """按计划顺序给出该页应有的非空文本段"""
    out = []
    if s["type"] == "cover":
        cover_title = os.environ.get(
            "PPT_TITLE",
            para_text_of_block(BLOCKS[s["blocks"][0]]))  # 封面大标题（默认=Word 文档标题块）
        out.append(cover_title)
        for bi in s["blocks"][1:]:
            t = para_text_of_block(BLOCKS[bi])
            if t:
                out.append(t)
        return out
    if s["type"] == "divider":
        return [s["title"]]
    for u in s["units"]:
        if u["kind"] == "stem":
            for bi in u["blocks"]:
                t = para_text_of_block(BLOCKS[bi])
                if t:
                    out.append(t)
        elif u["kind"] == "image":
            if u.get("caption"):
                out.append(u["caption"])
        elif u["kind"] == "figure_entry":
            t = para_text_of_block(BLOCKS[u["text_unit"]["block"]])
            if t:
                out.append(t)
            if u["image_unit"].get("caption"):
                out.append(u["image_unit"]["caption"])
        elif u["kind"] == "imgpair":
            for im in u["imgs"]:
                if im.get("caption"):
                    out.append(im["caption"])
        elif u["kind"] == "imgpair":
            for im in u["imgs"]:
                if im.get("caption"):
                    out.append(im["caption"])
        elif u["kind"] == "table":
            for row in BLOCKS[u["block"]]["rows"]:
                for c in row["cells"]:
                    for p in c["paras"]:
                        t = para_text_of_block(p)
                        if t:
                            out.append(t)
        elif u["kind"] == "tablepart":
            for row in u["rows"]:
                for c in row["cells"]:
                    for p in c["paras"]:
                        t = para_text_of_block(p)
                        if t:
                            out.append(t)
        else:
            t = para_text_of_block(BLOCKS[u["block"]])
            if t:
                out.append(t)
    return out

def actual_slide_texts(sld):
    """按文档顺序取内容区文本（排除标题/页脚/装饰）；按段落聚合 a:t"""
    out = []
    spTree = sld.find(".//" + qn("p:cSld") + "/" + qn("p:spTree"))
    for sp in spTree:
        tag = etree.QName(sp).localname
        off = sp.find(".//" + qn("a:off"))
        y = int(off.get("y")) if off is not None else 0
        x = int(off.get("x")) if off is not None else 0
        if y < 1000000 or (y > 6400000 and x > 8500000):
            continue  # 标题区 / 页脚
        if tag in ("sp", "graphicFrame", "pic"):
            for p_el in sp.iter(qn("a:p")):
                # 排除公式 Fallback 线性文本（Choice 分支的 OMML 才是正文）
                txts = []
                for t in p_el.iter(qn("a:t")):
                    anc = t.getparent()
                    in_fb = False
                    while anc is not None:
                        if anc.tag == qn("mc:Fallback"):
                            in_fb = True
                            break
                        anc = anc.getparent()
                    if not in_fb:
                        txts.append(t.text or "")
                txt = "".join(txts).strip()
                if txt:
                    out.append(txt)
    return out

zf = zipfile.ZipFile(PPTX)
slide_names = sorted([n for n in zf.namelist()
                      if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                     key=lambda n: int(re.search(r"\d+", n).group()))
print("slide count:", len(slide_names), "plan:", len(PLAN))
assert len(slide_names) == len(PLAN)

errors, warns = [], []

# rels per slide for image ratio check
def slide_rels(idx):
    try:
        data = zf.read("ppt/slides/_rels/slide%d.xml.rels" % idx)
    except KeyError:
        return {}
    root = etree.fromstring(data)
    return {rel.get("Id"): rel.get("Target") for rel in root}

total_math = 0
total_pics = 0
for i, name in enumerate(slide_names):
    sld = etree.fromstring(zf.read(name))
    s = PLAN[i]
    n = i + 1
    # A 过渡
    trans = sld.findall(".//" + qn("p:transition"))
    if not trans:
        errors.append(f"p{n}: 缺少过渡动画")
    elif trans[0].find(qn("p:fade")) is None:
        errors.append(f"p{n}: 过渡不是 fade")
    # B 时序
    timing = sld.find(qn("p:timing"))
    if s["type"] == "content":
        if timing is None:
            errors.append(f"p{n}: 缺少动画时序")
        else:
            beats = timing.findall(".//" + qn("p:seq") + "/" + qn("p:cTn") +
                                   "/" + qn("p:childTnLst") + "/" + qn("p:par"))
            if len(beats) != len(s["units"]):
                errors.append(f"p{n}: 拍数 {len(beats)} != 计划 {len(s['units'])}")
            spids_used = {e.get("spid") for e in timing.iter(qn("p:spTgt"))}
            ids_in_slide = {e.get("id") for e in sld.iter(qn("p:cNvPr"))}
            missing = spids_used - ids_in_slide
            if missing:
                errors.append(f"p{n}: 时序引用不存在的形状 {missing}")
    else:
        if timing is not None:
            warns.append(f"p{n}: 非内容页含时序（可忽略）")
    # C 页脚
    all_text = "".join(t.text or "" for t in sld.iter(qn("a:t")))
    footer_txt = os.environ.get("PPT_FOOTER", "Oli教室专用")
    if footer_txt not in all_text:
        errors.append(f"p{n}: 缺少页脚署名 {footer_txt}")
    # D 覆盖
    exp = expected_content_texts(s)
    act = actual_slide_texts(sld)
    if exp != act:
        # 找出差异
        for k in range(max(len(exp), len(act))):
            e = exp[k] if k < len(exp) else "<无>"
            a = act[k] if k < len(act) else "<无>"
            if e != a:
                errors.append(f"p{n}: 第{k+1}段不一致\n  期望: {e[:60]}\n  实际: {a[:60]}")
                break
        if len(exp) != len(act):
            errors.append(f"p{n}: 段数不匹配 exp={len(exp)} act={len(act)}")
    # E/F 公式与图片
    n_math = len(list(sld.iter(qn("a14:m"))))
    total_math += n_math
    pics = sld.findall(".//" + qn("p:pic"))
    total_pics += len(pics)
    rels = slide_rels(n)
    for pic in pics:
        blip = pic.find(".//" + qn("a:blip"))
        rid = blip.get(qn("r:embed"))
        target = rels.get(rid, "")
        ext = pic.find(".//" + qn("a:ext"))
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        base = os.path.basename(target)
        info = D["img_info"].get(base)
        if info:
            native = info["w"] / info["h"]
            shown = cx / cy
            if abs(native - shown) / native > 0.01:
                errors.append(f"p{n}: 图片 {base} 宽高比失真 native={native:.3f} shown={shown:.3f}")
    # G/H 字体与颜色
    for e in sld.iter(qn("a:latin")):
        if e.get("typeface") not in FONTS_OK:
            errors.append(f"p{n}: 非白名单字体 {e.get('typeface')}")
    for e in sld.iter(qn("a:ea")):
        if e.get("typeface") not in FONTS_OK:
            errors.append(f"p{n}: 非白名单字体(ea) {e.get('typeface')}")
    for e in sld.iter(qn("a:srgbClr")):
        v = (e.get("val") or "").upper()
        if v not in COLORS_OK:
            errors.append(f"p{n}: 非白名单颜色 {v}")
    # I 越界（页脚区域除外）
    spTree = sld.find(".//" + qn("p:cSld") + "/" + qn("p:spTree"))
    for sp in spTree:
        off = sp.find(".//" + qn("a:off"))
        ext2 = sp.find(".//" + qn("a:ext"))
        if off is None or ext2 is None:
            continue
        y = int(off.get("y")); cy2 = int(ext2.get("cy"))
        if y > 6400000:
            continue  # 页脚
        if y + cy2 > int(7.03 * 914400):
            tag = etree.QName(sp).localname
            errors.append(f"p{n}: 形状越界 y+h={((y+cy2)/914400):.2f}in ({tag})")

# 数学总数 / 图片期望数（从 Word 提取结果推导，勿写死——各文档图数不同）
doc_math = 0
doc_pics = 0
def count_math(b):
    global doc_math, doc_pics
    if b["kind"] == "para":
        doc_math += sum(1 for r in b["runs"] if "math" in r)
        doc_pics += sum(1 for r in b["runs"] if "img" in r)
    else:
        for row in b["rows"]:
            for c in row["cells"]:
                for p in c["paras"]:
                    doc_math += sum(1 for r in p["runs"] if "math" in r)
                    doc_pics += sum(1 for r in p["runs"] if "img" in r)
for b in BLOCKS:
    count_math(b)
print(f"math: docx={doc_math} pptx={total_math}")
print(f"pics: pptx={total_pics} (expect {doc_pics})")
if doc_math != total_math:
    errors.append(f"公式数不一致 docx={doc_math} pptx={total_math}")
if total_pics != doc_pics:
    errors.append(f"图片数不一致 expect={doc_pics} actual={total_pics}")

print("\n==== 结果 ====")
print("errors:", len(errors))
for e in errors[:40]:
    print(" -", e)
print("warns:", len(warns))
for w in warns[:10]:
    print(" -", w)
sys.exit(1 if errors else 0)
