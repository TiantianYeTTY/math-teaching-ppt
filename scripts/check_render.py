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
    frags = []
    for bi, blk in enumerate(d.get("blocks", [])):
        if blk.get("type") != 0:
            continue
        for ln in blk["lines"]:
            txt = "".join(s["text"] for s in ln["spans"]).strip()
            if txt:
                frags.append([bi, pymupdf.Rect(ln["bbox"]), txt])
    # 同一文本块内 y 向重叠 >50% 且 x 向相交/相邻的片段 = 同一视觉行（LO 兜底在字体
    # 切换处会把一行拆成多个 line 对象，bbox 接缝交叠造成伪重叠报警）。表格同一行
    # 的各单元格虽同 y 带但 x 向分离（列间有真实间隙），不会被合并。跨文本块（跨
    # shape）的真实遮挡不受影响——那是不同 block 之间的比较。
    frags.sort(key=lambda f: (f[0], f[1].y0, f[1].x0))
    out = []
    for bi, r, t in frags:
        hit = None
        for o in out:
            if o[0] != bi:
                continue
            inter_y = min(o[1].y1, r.y1) - max(o[1].y0, r.y0)
            minh = min(o[1].y1 - o[1].y0, r.y1 - r.y0)
            if minh <= 0 or inter_y <= 0.5 * minh:
                continue
            inter_x = min(o[1].x1, r.x1) - max(o[1].x0, r.x0)
            gap_x = max(o[1].x0, r.x0) - min(o[1].x1, r.x1)
            if inter_x > 0 or gap_x < 0.02 * 72:
                hit = o
                break
        if hit is not None:
            hit[1] = hit[1] | r
            hit[2] = hit[2] + t
        else:
            out.append([bi, r, t])
    return [(r, t) for _, r, t in out]

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
