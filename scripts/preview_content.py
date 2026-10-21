#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Designed by TiantianYe
# 适配指南：extract 之后运行，打印 content.json 的样式映射、标题大纲与全文渲染预览（extract/preview.txt），
# 用于人工核对 Word 解析是否完整（题目/答案标记、公式、表格、图注的形态）。无参数。
"""输出 content.json 的可读预览：样式名映射、标题大纲、全文段落渲染（公式计为⟨M⟩、图片计为⟨IMG⟩）。"""
import json, os, sys, zipfile
from lxml import etree

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
# DOCX 可选传入：仅用于读取 styles.xml 的样式名映射；不传则跳过映射段
DOCX = sys.argv[1] if len(sys.argv) > 1 else None
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def w(t): return "{%s}%s" % (WNS, t)

with open(os.path.join(ROOT, "extract", "content.json"), encoding="utf-8") as f:
    data = json.load(f)

# 样式 ID → 名称（需读取 word/styles.xml；未传 DOCX 时跳过）
id2name = {}
if DOCX:
    sroot = etree.fromstring(zipfile.ZipFile(DOCX).read("word/styles.xml"))
    for st in sroot.findall(w("style")):
        sid = st.get(w("styleId"))
        nm = st.find(w("name"))
        id2name[sid] = nm.get(w("val")) if nm is not None else sid

def runs_text(runs):
    out = []
    for r in runs:
        if "t" in r:
            t = r["t"]
            if r.get("b"): t = "**" + t + "**"
            out.append(t)
        elif "math" in r: out.append("⟨M⟩")
        elif "img" in r: out.append("⟨IMG:%s⟩" % r["img"]["rid"])
        elif "br" in r: out.append("⏎")
    return "".join(out)

lines = []
lines.append("=== 样式映射 ===")
for sid, nm in sorted(id2name.items()):
    lines.append(f"  {sid!r:20s} -> {nm}")

data["id2name"] = id2name
with open(os.path.join(ROOT, "extract", "content.json"), "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False)

lines.append("\n=== 大纲（标题块）===")
for i, b in enumerate(data["blocks"]):
    if b["kind"] == "para" and b["style"] in ("1", "2", "3", "4", "Title"):
        lines.append(f"  [{i:3d}] H{b['style']} {runs_text(b['runs'])}")

lines.append("\n=== 全文渲染 ===")
num_count = 0
for i, b in enumerate(data["blocks"]):
    if b["kind"] == "para":
        if b.get("num") is not None:
            num_count += 1
        prefix = f"[{i:3d}|{b['style'] or '-'}{'|N' + str(b['num']) if b.get('num') is not None else ''}] "
        txt = runs_text(b["runs"])
        lines.append(prefix + (txt if txt else "(空)"))
    else:
        lines.append(f"[{i:3d}|TABLE] {len(b['rows'])}行")
        for ri, row in enumerate(b["rows"]):
            cells = []
            for c in row["cells"]:
                ct = " / ".join(runs_text(p["runs"]) for p in c["paras"] if runs_text(p["runs"]))
                cells.append(ct if ct else "·")
            lines.append(f"    r{ri}: " + " ┃ ".join(cells))

lines.append(f"\n=== numPr 列表段落数: {num_count} ===")
out_path = os.path.join(ROOT, "extract", "preview.txt")
with open(out_path, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("written", out_path, "lines:", len(lines))
