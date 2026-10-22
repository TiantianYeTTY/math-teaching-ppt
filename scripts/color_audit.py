#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run 级颜色审计（pipeline.md §3）：列出全片蓝色(3E6FC9)/红色(D9534F) run，
人工核对语义白名单。zipfile+lxml，无第三方依赖。"""
import os, re, sys, zipfile
from lxml import etree

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
PPTX = os.path.join(ROOT, "out", os.environ.get("PPT_DECK_NAME", "teaching_ppt.pptx"))

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
A = "{%s}" % NS["a"]

zf = zipfile.ZipFile(PPTX)
slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                key=lambda n: int(re.search(r"\d+", n).group()))
blue, red = [], []
for i, name in enumerate(slides):
    sld = etree.fromstring(zf.read(name))
    for p_el in sld.iter(A + "p"):
        ptxt = "".join(t.text or "" for t in p_el.iter(A + "t"))
        for r in p_el.iter(A + "r"):
            c = r.find(A + "rPr/" + A + "solidFill/" + A + "srgbClr")
            if c is None:
                continue
            v = (c.get("val") or "").upper()
            rt = "".join(t.text or "" for t in r.iter(A + "t"))
            if v == "3E6FC9":
                blue.append((i + 1, ptxt[:42], rt[:48]))
            elif v == "D9534F":
                red.append((i + 1, ptxt[:42], rt[:48]))

print(f"== 蓝色 run 共 {len(blue)} ==")
for pg, pt, rt in blue:
    print(f"  p{pg:02d} | {pt} || run: {rt}")
print(f"== 红色 run 共 {len(red)} ==")
for pg, pt, rt in red:
    print(f"  p{pg:02d} | {pt} || run: {rt}")
