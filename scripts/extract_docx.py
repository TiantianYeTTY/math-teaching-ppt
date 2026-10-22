#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Designed by TiantianYe
# 适配指南：用法 `python extract_docx.py <word稿路径>`。
# 输出 <工程根>/extract/content.json 与 extract/media/。工程根=脚本所在目录的上一级。
# 新年级/新文档直接换 word 路径即可；页眉文字会随 Word 原文提取（页脚署名在 build_pptx.py 用 PPT_FOOTER 配置）。
"""从七年级期中考点梳理报告.docx 按文档顺序提取全部内容：
文本 run（含粗体/斜体/上下标）、OMML 公式原始 XML、图片（rId+EMU 尺寸）、表格。
输出 extract/content.json 与 extract/media/。
"""
import json, os, shutil, sys, zipfile
from lxml import etree
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
if len(sys.argv) < 2:
    sys.exit("用法: python extract_docx.py <word稿路径>")
DOCX = sys.argv[1]
OUT = os.path.join(os.path.dirname(BASE), "extract")  # 工程根/extract（与 plan/build 的读取约定一致）
MEDIA_OUT = os.path.join(OUT, "media")

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
ANS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPNS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
PRNS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

def w(tag): return "{%s}%s" % (WNS, tag)
def m(tag): return "{%s}%s" % (MNS, tag)
def a(tag): return "{%s}%s" % (ANS, tag)
def wp(tag): return "{%s}%s" % (WPNS, tag)
def r(tag): return "{%s}%s" % (RNS, tag)

def run_props(r_el):
    """提取 run 级格式：b/i/u/vertAlign"""
    p = {}
    rpr = r_el.find(w("rPr"))
    if rpr is not None:
        b = rpr.find(w("b"))
        if b is not None and b.get(w("val"), "1") != "0":
            p["b"] = True
        i = rpr.find(w("i"))
        if i is not None and i.get(w("val"), "1") != "0":
            p["i"] = True
        u = rpr.find(w("u"))
        if u is not None and u.get(w("val"), "single") != "none":
            p["u"] = True
        va = rpr.find(w("vertAlign"))
        if va is not None:
            p["va"] = va.get(w("val"))
    return p

def parse_image(drawing):
    """从 w:drawing 提取 blip rId 与 extent(EMU)"""
    blip = drawing.find(".//" + a("blip"))
    if blip is None:
        return None
    rid = blip.get(r("embed"))
    if not rid:
        return None
    ext = drawing.find(".//" + wp("extent"))
    cx = cy = None
    if ext is not None:
        cx = int(ext.get("cx")); cy = int(ext.get("cy"))
    return {"rid": rid, "cx": cx, "cy": cy}

def parse_runs(container):
    """按文档顺序解析容器(w:p / w:hyperlink)内的 run 序列"""
    runs = []
    for child in container:
        tag = etree.QName(child).localname
        ns = etree.QName(child).namespace
        if ns == WNS and tag == "r":
            props = run_props(child)
            for rc in child:
                rtag = etree.QName(rc).localname
                if rtag == "t":
                    txt = rc.text or ""
                    runs.append({"t": txt, **props})
                elif rtag == "br":
                    runs.append({"br": True})
                elif rtag == "tab":
                    runs.append({"t": "\t"})
                elif rtag == "drawing":
                    img = parse_image(rc)
                    if img:
                        runs.append({"img": img})
                elif rtag in ("fldChar", "instrText", "rPr", "pPr"):
                    continue
                # 其余忽略
        elif ns == WNS and tag == "hyperlink":
            runs.extend(parse_runs(child))
        elif ns == MNS and tag in ("oMath", "oMathPara"):
            xml = etree.tostring(child, encoding="unicode")
            runs.append({"math": xml, "disp": tag == "oMathPara"})
        elif ns == WNS and tag in ("proofErr", "bookmarkStart", "bookmarkEnd",
                                    "pPr", "sectPr", "commentRangeStart",
                                    "commentRangeEnd", "del", "ins"):
            continue
        else:
            SKIP_TAGS[tag] = SKIP_TAGS.get(tag, 0) + 1
    return runs

def parse_paragraph(p):
    ppr = p.find(w("pPr"))
    style = None
    num = None
    if ppr is not None:
        ps = ppr.find(w("pStyle"))
        if ps is not None:
            style = ps.get(w("val"))
        npr = ppr.find(w("numPr"))
        if npr is not None:
            ilvl_el = npr.find(w("ilvl"))
            num = int(ilvl_el.get(w("val"), "0")) if ilvl_el is not None else 0
    runs = parse_runs(p)
    return {"kind": "para", "style": style, "num": num, "runs": runs}

def parse_table(tbl):
    rows = []
    for tr in tbl.findall(w("tr")):
        cells = []
        for tc in tr.findall(w("tc")):
            tcpr = tc.find(w("tcPr"))
            span = 1
            vmerge = None
            if tcpr is not None:
                gs = tcpr.find(w("gridSpan"))
                if gs is not None:
                    span = int(gs.get(w("val"), "1"))
                vm = tcpr.find(w("vMerge"))
                if vm is not None:
                    vmerge = vm.get(w("val"), "continue")
            paras = []
            for p in tc.findall(w("p")):
                paras.append(parse_paragraph(p))
            cells.append({"paras": paras, "span": span, "vmerge": vmerge})
        rows.append({"cells": cells})
    return {"kind": "table", "rows": rows}

SKIP_TAGS = {}

def main():
    os.makedirs(MEDIA_OUT, exist_ok=True)
    zf = zipfile.ZipFile(DOCX)
    doc = etree.fromstring(zf.read("word/document.xml"))
    body = doc.find(w("body"))

    # rels
    rels_root = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
    rid2target = {}
    for rel in rels_root:
        rid2target[rel.get("Id")] = rel.get("Target")

    blocks = []
    for child in body:
        tag = etree.QName(child).localname
        ns = etree.QName(child).namespace
        if ns == WNS and tag == "p":
            blocks.append(parse_paragraph(child))
        elif ns == WNS and tag == "tbl":
            blocks.append(parse_table(child))

    # 拆分混合段：同一段落内"文字/公式 + 图片"拆为 文字段(前) + 图片段(后)，
    # 否则图片单元渲染路径只出图，描述文字与内嵌公式整段丢失（EA-31）
    split_blocks = []
    for b in blocks:
        if b.get("kind") == "para":
            runs = b.get("runs", [])
            has_img = any("img" in r for r in runs)
            has_content = any(("t" in r and r.get("t", "").strip()) or "math" in r
                              for r in runs)
            if has_img and has_content:
                text_runs = [r for r in runs if "img" not in r]
                img_runs = [r for r in runs if "img" in r]
                if any(("t" in r and r.get("t", "").strip()) or "math" in r
                       for r in text_runs):
                    nb = dict(b)
                    nb["runs"] = text_runs
                    split_blocks.append(nb)
                for ir in img_runs:
                    split_blocks.append({"kind": "para", "style": b.get("style"),
                                         "num": None, "runs": [ir]})
                continue
        split_blocks.append(b)
    blocks = split_blocks

    # 提取 media
    media_map = {}
    for name in zf.namelist():
        if name.startswith("word/media/"):
            base = os.path.basename(name)
            data = zf.read(name)
            with open(os.path.join(MEDIA_OUT, base), "wb") as f:
                f.write(data)
            media_map[base] = len(data)

    # 图片原生像素尺寸
    img_info = {}
    for base in media_map:
        path = os.path.join(MEDIA_OUT, base)
        try:
            im = Image.open(path)
            img_info[base] = {"w": im.width, "h": im.height, "fmt": im.format}
        except Exception as e:
            img_info[base] = {"err": str(e)}

    out = {
        "docx": DOCX,
        "blocks": blocks,
        "rid2target": rid2target,
        "media_bytes": media_map,
        "img_info": img_info,
        "skip_tags": SKIP_TAGS,
    }
    with open(os.path.join(OUT, "content.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)

    # 统计
    n_para = sum(1 for b in blocks if b["kind"] == "para")
    n_tbl = sum(1 for b in blocks if b["kind"] == "table")
    n_math = 0
    n_img = 0
    n_chars = 0
    styles = {}
    for b in blocks:
        if b["kind"] == "para":
            st = b["style"] or "(none)"
            styles[st] = styles.get(st, 0) + 1
            for r_ in b["runs"]:
                if "math" in r_:
                    n_math += 1
                elif "img" in r_:
                    n_img += 1
                elif "t" in r_:
                    n_chars += len(r_["t"])
        else:
            for row in b["rows"]:
                for c in row["cells"]:
                    for p in c["paras"]:
                        for r_ in p["runs"]:
                            if "math" in r_:
                                n_math += 1
                            elif "img" in r_:
                                n_img += 1
                            elif "t" in r_:
                                n_chars += len(r_["t"])
    print("blocks:", len(blocks), "paras:", n_para, "tables:", n_tbl)
    print("math:", n_math, "images:", n_img, "chars:", n_chars)
    print("styles:", json.dumps(styles, ensure_ascii=False))
    print("skip_tags:", SKIP_TAGS)
    print("media files:", len(media_map))
    for base, info in list(img_info.items())[:50]:
        print("  img", base, info)

if __name__ == "__main__":
    main()
