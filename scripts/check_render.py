#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Designed by TiantianYe
# 适配指南：`python check_render.py <soffice 转出的 pdf>`。
# 用 PDF 文本 bbox 检测行间重叠与越底（兜底文本形态；真公式形态靠 plan 估算余量保证）。
"""渲染后自动检查：文本行 两两重叠（排除 <2 字符的公式点符号）+ 内容越底。
用法：check_render.py <pdf路径> ；发现问题 exit 1。"""
import os
import sys
import pymupdf

PDF = sys.argv[1] if len(sys.argv) > 1 else "render/test_hira.pdf"
BOTTOM = 7.02 * 72
doc = pymupdf.open(PDF)
issues = []

def lines(page):
    d = page.get_text("dict")
    out = []
    for blk in d.get("blocks", []):
        if blk.get("type") != 0:
            continue
        for ln in blk["lines"]:
            txt = "".join(s["text"] for s in ln["spans"]).strip()
            if txt:
                out.append((pymupdf.Rect(ln["bbox"]), txt))
    return out

def inter_area(a, b):
    r = a & b
    return r.get_area() if not r.is_empty else 0

for pno in range(len(doc)):
    page = doc[pno]
    L = lines(page)
    for r, t in L:
        FOOTER_KEY = os.environ.get("PPT_FOOTER", "Oli教室专用")[:4]
        if r.y1 > BOTTOM and FOOTER_KEY not in t.replace(" ", ""):
            issues.append(f"p{pno+1} 越底 y1={r.y1/72:.2f}in: {t[:30]}")
    n = len(L)
    for i in range(n):
        for j in range(i + 1, n):
            a, ta = L[i]
            b, tb = L[j]
            import re as _re
            if len(_re.sub(r"[\s⋅•·\.\-—∙]", "", ta)) < 2 or len(_re.sub(r"[\s⋅•·\.\-—∙]", "", tb)) < 2:
                continue
            ia = inter_area(a, b)
            if ia > 0.3 * min(a.get_area(), b.get_area()) and ia > 60:
                issues.append(f"p{pno+1} 重叠: [{ta[:20]}]×[{tb[:20]}] "
                              f"y {a.y0/72:.2f}-{a.y1/72:.2f}/{b.y0/72:.2f}-{b.y1/72:.2f}")

if issues:
    print(f"渲染检查：{len(issues)} 个问题")
    for x in issues[:30]:
        print(" -", x)
    sys.exit(1)
print(f"渲染检查：{len(doc)} 页全部干净（无重叠/无越底）")
