#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Designed by TiantianYe
# 适配指南：plan_slides.py 之后运行。校验 Word 非空块 100% 被分配、无重复；打印计划预览。无参数。
"""计划预览与覆盖校验"""
import json, os, sys

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
D = json.load(open(os.path.join(ROOT, "extract", "content.json"), encoding="utf-8"))
BLOCKS = D["blocks"]
PLAN = json.load(open(os.path.join(ROOT, "plan", "slides.json"), encoding="utf-8"))

def btext(i):
    b = BLOCKS[i]
    if b["kind"] == "para":
        t = "".join(r.get("t", "") for r in b["runs"] if "t" in r)
        m = sum(1 for r in b["runs"] if "math" in r)
        img = sum(1 for r in b["runs"] if "img" in r)
        tag = f"⟨M{m}⟩" if m else ""
        tag += f"⟨I{img}⟩" if img else ""
        return (t[:34] + tag).strip() or "(空)"
    return f"TABLE {len(b['rows'])}r"

# 覆盖检查
seen = {}
lines = []
for si, s in enumerate(PLAN):
    sid = f"s{si:02d}"
    if s["type"] in ("cover", "divider"):
        lines.append(f"{sid} [{s['type']}] {s['title']}")
        continue
    sc = s.get("scale", 1.0)
    lines.append(f"{sid} [content] {s['title']}{' (续%d)' % s['cont'] if s.get('cont') else ''} scale={sc:.2f}")
    for u in s["units"]:
        if u["kind"] == "stem":
            lines.append(f"   BEAT-stem: " + " | ".join(f"[{i}]{btext(i)}" for i in u["blocks"]))
            for i in u["blocks"]:
                seen[i] = seen.get(i, 0) + 1
        elif u["kind"] == "image":
            lines.append(f"   BEAT-image: [{u['block']}]{btext(u['block'])} cap={u.get('caption','')[:20]}")
            seen[u["block"]] = seen.get(u["block"], 0) + 1
            if "caption_block" in u:
                seen[u["caption_block"]] = seen.get(u["caption_block"], 0) + 1
        elif u["kind"] == "figure_entry":
            tu, iu = u["text_unit"], u["image_unit"]
            lines.append(f"   BEAT-figure: [{tu['block']}]{btext(tu['block'])} cap={iu.get('caption','')[:20]}")
            seen[tu["block"]] = seen.get(tu["block"], 0) + 1
            seen[iu["block"]] = seen.get(iu["block"], 0) + 1
            if "caption_block" in iu:
                seen[iu["caption_block"]] = seen.get(iu["caption_block"], 0) + 1
        elif u["kind"] == "imgpair":
            caps = " ┃ ".join(f"[{im['block']}]{im.get('caption','')[:16]}" for im in u["imgs"])
            lines.append(f"   BEAT-imgpair: {caps}")
            for im in u["imgs"]:
                seen[im["block"]] = seen.get(im["block"], 0) + 1
        elif u["kind"] == "tablepart":
            lines.append(f"   BEAT-tablepart: block[{u['block']}] rows={len(u['rows'])}")
        elif u["kind"] == "table":
            lines.append(f"   BEAT-table: [{u['block']}]{btext(u['block'])}")
            seen[u["block"]] = seen.get(u("block"), 0) + 1 if False else seen.get(u["block"], 0) + 1
        else:
            lines.append(f"   BEAT: [{u['block']}]{btext(u['block'])}")
            seen[u["block"]] = seen.get(u["block"], 0) + 1

missing = [i for i, b in enumerate(BLOCKS)
           if b["kind"] == "para" and i > 2
           and ("".join(r.get("t", "") for r in b["runs"]).strip()
                or any("math" in r or "img" in r for r in b["runs"]))
           and i not in seen and b.get("style") not in ("1", "2", "4")]
dup = [i for i, c in seen.items() if c > 1]
tables_in_plan = sum(1 for s in PLAN if s["type"] == "content"
                     for u in s["units"] if u["kind"] in ("table", "tablepart"))
print("missing blocks:", missing)
print("dup blocks:", dup)
print("tables/tables-parts in plan:", tables_in_plan)
out = "\n".join(lines)
open(os.path.join(ROOT, "plan", "plan_preview.txt"), "w", encoding="utf-8").write(out)
print("plan_preview written, lines:", len(lines))
