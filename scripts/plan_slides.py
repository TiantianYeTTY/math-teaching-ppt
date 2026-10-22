#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Designed by TiantianYe
# 适配指南：读取 <工程根>/extract/content.json，产出 <工程根>/plan/slides.json 与 plan_preview.txt。
# 版式常量（CONTENT_W/H、BODY_FONT、TABLE_FONT、SCALES）在文件头部按需调整；
# PPT_CALIB 环境变量为全局高度校准系数（渲染反馈迭代用）。
"""分页计划 v3（2026-10-21 晚，按用户第二次反馈重构）：
- 原子内容组：板块/整道题/变式 = 不可拆组；**页级再平衡只整组搬移，绝不拆组**
- 组超一整页才拆（题目|解析边界优先），拆块只在该组自己的连续页内微调
- 图片限高 2.4in（放大≤1.6 倍），使 图+题 同页、双图同页
- 题干合并规则升级：题图（可隔图注）紧邻题干一律并入题干拍
- 估算器按 OMML 渲染加安全余量（公式 0.80em/字符+2.4em、含式段落行高 1.58 倍、
  表行 1.44 倍、规划容量 5.60in 留 0.26in 底部余量吸收公式渲染高度差）
- 字号：正文 16 / 表格 12；缩放档 [1.0, 0.94, 0.9, 0.86]；页内追加 ≥0.9，整组放置/拆块 ≥0.86
输出 plan/slides.json（每页含 units 扁平列表 + groups 分组结构）
"""
import json, math, os, re

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
with open(os.path.join(ROOT, "extract", "content.json"), encoding="utf-8") as f:
    D = json.load(f)
BLOCKS = D["blocks"]
RID2T = D["rid2target"]
IMG_INFO = D["img_info"]

CALIB = float(os.environ.get("PPT_CALIB", "1.0"))  # 渲染反馈迭代用全局校准系数

CONTENT_W = 12.33
CONTENT_H = 5.60        # 规划容量（构建放置区到 7.00，余量吸收公式渲染高度差）
BODY_FONT = 16.0
TABLE_FONT = 12.0
GAP = 0.12
SCALES = [1.0, 0.94, 0.9, 0.86, 0.82]
MIN_PACK = 0.9          # 页内追加组的最低档
MIN_SC = 0.86           # 整组放置/拆块的最低档

# 用户指定（2026-10-22，硬规则 UG）：同一道题尽量题干与图同页，可"适当缩小图片"顺利合并；
# 实在放不下才允许例外拆两页（拆分点在"题目|解析"边界）。
# 机制：按文件名限高等比缩小，est 与 build 同走 img_display_size，口径一致。
# ⚠️ 文件名以 extract/content.json 的 rid2target 实测为准（块号→文件映射勿凭清单顺序猜测，
# 九年级 v2 首轮按清单顺序猜全部错位一格）。取值方法：估算器网格搜索 (scale × cap) 找
# 图片最大可保留值；九年级 v2 定稿取值供量级参考：image3=1.6 / image5=1.0 /
# image6=0.8(配0.78档) / image13=1.0
IMG_MAX_H = {}   # 按项目填写；空 = 全部走默认限高（2.2in）

# 用户指定（2026-10-21）：组尾两页合并为一页，值为允许的最低档位。
# "5.3 双直角斜边中线与函数综合（立达中学 第 25 题）"原 46/47 两页合并
MANUAL_MERGE_TAIL = {
    "5.3 双直角斜边中线与函数综合（立达中学 第 25 题）": 0.82,
}

M_T = re.compile(r"<m:t(?:\s[^>]*)?>([^<]*)</m:t>")
HAS_TALL = re.compile(r"<m:(f|rad|nary)[ >]")

def para_text(b):
    if b["kind"] != "para":
        return ""
    return "".join(r.get("t", "") for r in b["runs"]).strip()

def first_run(b):
    if b["kind"] != "para":
        return None
    for r in b["runs"]:
        if "t" in r and r["t"].strip():
            return r
    return None

def is_bold_start(b):
    r = first_run(b)
    return r is not None and r.get("b")

def est_para_h(b, font):
    has_math = any("math" in r for r in b.get("runs", []))
    # 分层余量：段落含分式/根号/求和结构（真渲染撑行）才用 1.95；
    # 仅上标/线性公式不撑行，用 1.58 保持 v4 布局密度
    has_tall_math = has_math and any(
        HAS_TALL.search(r["math"]) for r in b.get("runs", []) if "math" in r)
    lf = (1.95 if has_tall_math else (1.58 if has_math else 1.45)) * CALIB
    wpt = 0.0
    for r in b.get("runs", []):
        if "t" in r:
            t = r["t"].replace("\t", "    ")
            wpt += sum(1.0 if ord(c) > 0x2E80 else 0.60 for c in t) * font
        elif "math" in r:
            wpt += (len("".join(M_T.findall(r["math"]))) * 0.80 + 2.4) * font
    if wpt == 0:
        return 0.0
    # 行宽余量 0.85in：bullet 缩进+标点禁则+真机字体度量会使贴线段提前折行，
    # 0.45 不足（六年级 EA-39：估 1 行实 2 行与下一条重叠）
    line_w = (CONTENT_W - 0.85) * 72
    lines = max(1, math.ceil(wpt / line_w))
    return lines * font * lf / 72 + 0.07 * CALIB

def est_math_h(b, font):
    xml = "".join(r["math"] for r in b["runs"] if "math" in r)
    wide = sum(len("".join(M_T.findall(x))) for x in [xml]) * 0.80 * font
    if wide > (CONTENT_W - 0.5) * 72:
        return 3.0 * font * 1.6 / 72 * CALIB + 0.1
    tall = 2.8 if HAS_TALL.search(xml) else 2.0
    return tall * font * 1.6 / 72 * CALIB + 0.1

def table_col_widths(tbl, font):
    ncol = max(len(row["cells"]) for row in tbl["rows"])
    weights = [0.0] * ncol
    for row in tbl["rows"]:
        for ci, c in enumerate(row["cells"]):
            t = "".join(p2text(p) for p in c["paras"])
            wlen = sum(1.0 if ord(ch) > 0x2E80 else 0.60 for ch in t) + sum(math_len(p) for p in c["paras"])
            weights[ci] = max(weights[ci], wlen)
    weights = [max(w, 2.5) for w in weights]
    s = sum(weights)
    widths = [max(0.5, min(6.2, w / s * CONTENT_W)) for w in weights]
    k = CONTENT_W / sum(widths)
    return [w * k for w in widths]

def p2text(p):
    return "".join(r.get("t", "") for r in p.get("runs", []) if "t" in r)

def math_len(b):
    n = 0.0
    if b["kind"] != "para":
        return 0.0
    for r in b["runs"]:
        if "math" in r:
            n += len("".join(M_T.findall(r["math"]))) * 0.80 + 2.4
    return n

def est_table_h(tbl, font, widths=None):
    if widths is None:
        widths = table_col_widths(tbl, font)
    total = 0.12
    for ri, row in enumerate(tbl["rows"]):
        lines = 1
        for ci, c in enumerate(row["cells"]):
            t = "".join(p2text(p) for p in c["paras"])
            ml = sum(math_len(p) for p in c["paras"])
            wlen = sum(1.0 if ord(ch) > 0x2E80 else 0.60 for ch in t) + ml
            if wlen == 0:
                continue
            cpl = max(4.0, (widths[ci] - 0.14) * 72 / (font * 1.02))
            lines = max(lines, math.ceil(wlen / cpl))
        rf = 1.44 * CALIB
        total += lines * font * rf / 72 + 0.13
    return total

def est_block_h(b, font, tfont):
    if b["kind"] == "table":
        return est_table_h(b, tfont)
    if b["kind"] == "para":
        for r in b["runs"]:
            if "img" in r:
                _, h_in, _ = img_display_size(r["img"]["rid"])
                return h_in + 0.06
    has_math_only = (b["kind"] == "para" and any("math" in r for r in b["runs"])
                     and not any("t" in r and r["t"].strip() for r in b["runs"]))
    if has_math_only:
        return est_math_h(b, font)
    return est_para_h(b, font)

def fix_extent_ratio(w0, h0, info):
    """Word extent 非等比（原稿手工拉伸）时，在 extent 框内 contain 等比修正，
    保证展示比例=原图像素比（validate F 项 ±1% 口径）。EA-37。"""
    nat = info["w"] / info["h"]
    if h0 > 0 and abs(w0 / h0 - nat) / nat > 0.01:
        if w0 / h0 > nat:
            w0 = h0 * nat
        else:
            h0 = w0 / nat
    return w0, h0

def img_display_size(rid):
    """限高 2.4in；缩小不受限，放大 ≤1.6 倍；窄小图可放大到宽 3.5in。
    IMG_MAX_H：用户指定合并页"适当缩小图片"时按文件名限高（等比缩小，est/build 同口径）"""
    target = RID2T[rid]
    base = os.path.basename(target)
    info = IMG_INFO[base]
    for b in BLOCKS:
        if b["kind"] == "para":
            for r in b["runs"]:
                if "img" in r and r["img"]["rid"] == rid:
                    cx, cy = r["img"]["cx"], r["img"]["cy"]
                    w0, h0 = cx / 914400, cy / 914400
                    break
            else:
                continue
            break
    else:
        w0, h0 = info["w"] / 96, info["h"] / 96
    w0, h0 = fix_extent_ratio(w0, h0, info)
    s = min(8.8 / w0, 2.2 / h0, 1.6)
    if w0 * s < 3.5 and h0 * s < 2.0:
        s = min(3.5 / w0, 2.2 / h0, 2.2)
    w_in, h_in = w0 * s, h0 * s
    cap = IMG_MAX_H.get(base)
    if cap and h_in > cap:
        w_in, h_in = w_in * cap / h_in, cap
    return w_in, h_in, target

def img_display_size_in(rid, max_w, max_h):
    """容器限宽/限高的等比尺寸；缩小不受限，放大 ≤1.6 倍，窄小图可到宽 3.5。
    IMG_MAX_H 同样生效（est/build 同口径）"""
    target = RID2T[rid]
    base = os.path.basename(target)
    info = IMG_INFO[base]
    for b in BLOCKS:
        if b["kind"] == "para":
            for r in b["runs"]:
                if "img" in r and r["img"]["rid"] == rid:
                    cx, cy = r["img"]["cx"], r["img"]["cy"]
                    w0, h0 = cx / 914400, cy / 914400
                    break
            else:
                continue
            break
    else:
        w0, h0 = info["w"] / 96, info["h"] / 96
    w0, h0 = fix_extent_ratio(w0, h0, info)
    s = min(max_w / w0, max_h / h0, 1.6)
    if w0 * s < 3.5 and h0 * s < 2.0:
        s = min(3.5 / w0, max_h / h0, 2.2)
    w_in, h_in = w0 * s, h0 * s
    cap = IMG_MAX_H.get(base)
    if cap and h_in > cap:
        w_in, h_in = w_in * cap / h_in, cap
    return w_in, h_in, target

def img_unit_h(u, font):
    _, h, _ = img_display_size(u["img_rid"])
    cap = 0.32 if u.get("caption") else 0.0
    return h + cap + 0.08

def imgpair_unit_h(u, font):
    hs = []
    for im in u["imgs"]:
        _, h, _ = img_display_size_in(im["rid"], 5.9, 2.2)
        hs.append(h)
    return max(hs) + 0.32 + 0.08

def figure_entry_unit_h(u, font, tfont):
    """题图对（描述块+图+图注）：文段 + 图单元 + 块间 0.06（与 build 放置一致）"""
    tu, iu = u["text_unit"], u["image_unit"]
    return est_block_h(BLOCKS[tu["block"]], font, tfont) + img_unit_h(iu, font) + 0.06

def unit_h(u, font, tfont):
    if u["kind"] == "imgpair":
        return imgpair_unit_h(u, font)
    if u["kind"] == "figure_entry":
        return figure_entry_unit_h(u, font, tfont)
    if u["kind"] == "image":
        return img_unit_h(u, font)
    if u["kind"] == "tablepart":
        return est_table_h({"rows": u["rows"]}, tfont, u.get("widths"))
    if u["kind"] == "stem":
        return sum(est_block_h(BLOCKS[i], font, tfont) for i in u["blocks"]) + 0.05 * (len(u["blocks"]) - 1)
    return est_block_h(BLOCKS[u["block"]], font, tfont)

def fit_scale(unit_list, min_sc=0.0):
    for sc in SCALES:
        if sc < min_sc:
            continue
        font = BODY_FONT * sc
        tfont = TABLE_FONT * sc
        total = sum(unit_h(u, font, tfont) for u in unit_list) + GAP * (len(unit_list) - 1)
        if total <= CONTENT_H:
            return sc
    return None

GROUP_GAP = 0.42   # 同页不同题/组之间的空行间距

def units_fill(units, sc):
    font = BODY_FONT * sc
    tfont = TABLE_FONT * sc
    return sum(unit_h(u, font, tfont) for u in units) + GAP * (len(units) - 1)

def groups_fill(gs, sc):
    base = units_fill([u for g in gs for u in g["units"]], sc)
    return base + GROUP_GAP * (len(gs) - 1)

# ---------- 组切分 ----------
STEM_LABELS = ("典型原题", "原题", "例")
ANS_TRIG = ("逐步解析", "要点解析", "答案", "解析")
NUM_BOLD = re.compile(r"^（[0-9０-９]+）")

groups = []
gunits = []
g_answered = False
cur_h1 = None
cur_h2 = None
in_stem = False

def close_group():
    global gunits, g_answered, in_stem
    if gunits:
        t = cur_h2 or cur_h1
        for u in gunits:
            u["gtitle"] = t
        groups.append({"title": t, "h1": cur_h1, "units": gunits})
    gunits = []
    g_answered = False
    in_stem = False

for idx, b in enumerate(BLOCKS):
    if idx in (1, 2):
        continue
    if b["kind"] == "para":
        style = b.get("style")
        if style == "1":
            if idx == 0:
                cur_h1 = para_text(b)
                continue
            close_group()
            cur_h1 = para_text(b)
            cur_h2 = None
            continue
        if style == "2":
            close_group()
            cur_h2 = para_text(b)
            continue
        if style == "4":
            close_group()
            in_stem = True
            gunits.append({"kind": "stem", "blocks": [idx]})
            continue
        txt = para_text(b)
        has_img = any("img" in r for r in b["runs"])
        has_math = any("math" in r for r in b["runs"])
        if not txt and not has_img and not has_math:
            continue
        if in_stem and any(txt.startswith(t) for t in ANS_TRIG):
            in_stem = False
            g_answered = True
            gunits.append({"kind": "block", "block": idx})
            continue
        starts_stem = False
        if is_bold_start(b):
            if any(txt.startswith(lb) for lb in STEM_LABELS):
                starts_stem = True
            elif NUM_BOLD.match(txt) and style in ("a0", "FirstParagraph"):
                starts_stem = True
            elif style == "aa":
                starts_stem = True
        if starts_stem:
            if g_answered:
                close_group()
            merged = False
            if gunits and not g_answered:
                last = gunits[-1]
                if last["kind"] == "image" and last.get("last_block") == idx - 1:
                    gunits.pop()
                    gunits.append({"kind": "stem", "blocks": [last["block"], idx]})
                    merged = True
                elif (last["kind"] == "block"
                      and BLOCKS[last["block"]].get("style") == "ImageCaption"
                      and len(gunits) >= 2 and gunits[-2]["kind"] == "image"
                      and sum(1 for x in gunits[-4:] if x["kind"] == "image") < 2):
                    # 仅吸收单张题图；多张连续题图保持独立（让双图同页）
                    img_u = gunits[-2]
                    cap_u = last
                    del gunits[-2:]
                    gunits.append({"kind": "stem",
                                   "blocks": [img_u["block"], cap_u["block"], idx]})
                    merged = True
            if not merged:
                gunits.append({"kind": "stem", "blocks": [idx]})
            in_stem = True
            continue
        if in_stem:
            gunits[-1]["blocks"].append(idx)
            continue
        if txt.startswith("板块") and is_bold_start(b):
            close_group()
            gunits.append({"kind": "block", "block": idx})
            continue
        if has_img:
            gunits.append({"kind": "image", "block": idx,
                           "img_rid": [r["img"]["rid"] for r in b["runs"] if "img" in r][0],
                           "last_block": idx})
            continue
        gunits.append({"kind": "block", "block": idx})
    else:
        in_stem = False
        gunits.append({"kind": "table", "block": idx})
close_group()

# 独立图单元吸收紧随其后的图注
for g in groups:
    us = g["units"]
    for j, u in enumerate(us):
        if u["kind"] == "image":
            nxt = us[j + 1] if j + 1 < len(us) else None
            if nxt and nxt["kind"] == "block":
                nb = BLOCKS[nxt["block"]]
                if nb.get("style") == "ImageCaption":
                    u["caption"] = para_text(nb)
                    u["caption_runs"] = nb.get("runs", [])
                    u["caption_block"] = nxt["block"]
                    nxt["consumed"] = True
    g["units"] = [u for u in us if not u.get("consumed")]
    # 相邻两张题图 → 左右并排（imgpair）
    merged_units = []
    j = 0
    while j < len(g["units"]):
        u = g["units"][j]
        nxt = g["units"][j + 1] if j + 1 < len(g["units"]) else None
        if u["kind"] == "image" and nxt is not None and nxt["kind"] == "image":
            merged_units.append({"kind": "imgpair", "gtitle": u.get("gtitle"),
                                 "imgs": [
                                     {"rid": u["img_rid"], "caption": u.get("caption", ""),
                                      "caption_runs": u.get("caption_runs"),
                                      "block": u["block"]},
                                     {"rid": nxt["img_rid"], "caption": nxt.get("caption", ""),
                                      "caption_runs": nxt.get("caption_runs"),
                                      "block": nxt["block"]}]})
            j += 2
            continue
        merged_units.append(u)
        j += 1
    g["units"] = merged_units
    # 用户指定（2026-10-21）：题图集"一页一题一图"——把紧邻图片单元之前的非空
    # 文字块与该图绑定成原子单元（题在前、图在后，装箱不拆），避免描述留在
    # 上一页页尾、图落到下一页的题图错位。空段不绑；两图相邻时前图无前置
    # 文字块也不绑，imgpair 并排逻辑不受影响
    bound_units = []
    for u in g["units"]:
        if (u["kind"] == "image" and bound_units
                and bound_units[-1]["kind"] == "block"
                and not bound_units[-1].get("bound_to_image")):
            pb = BLOCKS[bound_units[-1]["block"]]
            t = "".join(r.get("t", "") for r in pb.get("runs", []) if "t" in r)
            if t.strip():
                bound_units[-1] = {"kind": "figure_entry", "gtitle": u.get("gtitle"),
                                   "text_unit": bound_units[-1], "image_unit": u}
                continue
        bound_units.append(u)
    g["units"] = bound_units

# ---------- 表格均衡拆行 ----------
def split_table_rows(u):
    tbl = BLOCKS[u["block"]]
    header = tbl["rows"][0]
    rows = tbl["rows"][1:]
    n = len(rows)
    total = est_table_h({"rows": [header] + rows}, TABLE_FONT * MIN_SC)
    n_parts = max(1, math.ceil(total / (CONTENT_H - 0.25)))
    per = math.ceil(n / n_parts)
    parts = []
    k = 0
    while k < n:
        pr = [header] + rows[k:k + per]
        parts.append({"kind": "tablepart", "block": u["block"], "rows": pr,
                      "gtitle": u.get("gtitle"),
                      "widths": table_col_widths({"rows": pr}, TABLE_FONT * MIN_SC)})
        k += per
    return parts

# ---------- 组内拆块（仅当组超一整页） ----------
def chunk_group(g, pre_groups=None):
    units = []
    if pre_groups:
        for gg in pre_groups:
            units.extend(gg["units"])
    units.extend(g["units"])
    chunks = None
    stems = [i for i, u in enumerate(units) if u["kind"] == "stem"]
    if stems:
        k = stems[-1] + 1
        if 0 < k < len(units):
            a, bch = units[:k], units[k:]
            if fit_scale(a, MIN_SC) is not None and fit_scale(bch, MIN_SC) is not None:
                chunks = [a, bch]
    if chunks is None:
        chunks = []
        buf = []
        for u in units:
            cand = buf + [u]
            if fit_scale(cand, MIN_SC) is not None:
                buf = cand
                continue
            # 题干单元本身超过任何档位 → 按块（图/图注/原题/各小问）拆流入
            if u["kind"] == "stem" and fit_scale([u], 0.0) is None and len(u["blocks"]) > 1:
                for bi in u["blocks"]:
                    bu = {"kind": "block", "block": bi, "gtitle": u.get("gtitle")}
                    cand = buf + [bu]
                    if fit_scale(cand, MIN_SC) is not None:
                        buf = cand
                    else:
                        if buf:
                            chunks.append(buf)
                            buf = []
                        buf = [bu]
                continue
            if u["kind"] == "table" and fit_scale([u], MIN_SC) is None:
                for part in split_table_rows(u):
                    part["gtitle"] = u.get("gtitle")
                    cand = buf + [part]
                    if fit_scale(cand, MIN_SC) is not None:
                        buf = cand
                    else:
                        if buf:
                            chunks.append(buf)
                            buf = []
                        buf = [part]
                continue
            if buf:
                chunks.append(buf)
                buf = []
            buf = [u]
        if buf:
            chunks.append(buf)
    gl = [{"type": "content", "title": g["title"], "h1": g["h1"],
           "groups": [{"units": ch, "gtitle": g["title"]}], "scale": 1.0, "cont": 0}
          for ch in chunks]
    group_rebalance(gl)
    group_absorb(gl)
    group_rebalance(gl)
    tail_balance(gl)
    # 用户指定：组尾两页合并（MANUAL_MERGE_TAIL），失败则保持拆分并告警
    if g["title"] in MANUAL_MERGE_TAIL and len(gl) >= 2:
        mn = MANUAL_MERGE_TAIL[g["title"]]
        lu = [u for gg in gl[-1]["groups"] for u in gg["units"]]
        pu = [u for gg in gl[-2]["groups"] for u in gg["units"]]
        sc = fit_scale(pu + lu, mn)
        if sc is not None:
            gl[-2:] = [{"type": "content", "title": g["title"], "h1": g["h1"],
                        "groups": [{"units": pu + lu, "gtitle": g["title"]}],
                        "scale": sc, "cont": 0}]
            print(f"[manual] 组尾两页已合并为 1 页（scale={sc}）: {g['title']}")
        else:
            print(f"[warn] MANUAL_MERGE_TAIL 未通过容量校验(最低档{mn})，保持拆分: {g['title']}")
    return gl

def tail_balance(gl):
    """组已跨多页时，在各块之间移动整单元填实尾页（不产生新拆分）"""
    if len(gl) < 2:
        return
    guard = 0
    while guard < 12:
        guard += 1
        last, prev = gl[-1], gl[-2]
        last_units = [u for gg in last["groups"] for u in gg["units"]]
        prev_units = [u for gg in prev["groups"] for u in gg["units"]]
        if not prev_units or not last_units:
            break
        if units_fill(last_units, last.get("scale", 1.0)) > 0.45 * CONTENT_H:
            break
        u = prev_units[-1]
        new_prev = prev_units[:-1]
        new_last = [u] + last_units
        sp = fit_scale(new_prev, MIN_SC)
        sl = fit_scale(new_last, MIN_SC)
        if sp is None or sl is None:
            break
        if units_fill(new_prev, sp) < 0.5 * CONTENT_H:
            break
        gt = prev["groups"][0].get("gtitle")
        prev["groups"] = [{"units": new_prev, "gtitle": gt}]
        last["groups"] = [{"units": new_last, "gtitle": gt}]
        prev["scale"], last["scale"] = sp, sl

def group_rebalance(gl):
    """整组搬移：末页 <35% 时把前一页最后一组下移"""
    if len(gl) < 2:
        return
    guard = 0
    while guard < 20:
        guard += 1
        last, prev = gl[-1], gl[-2]
        if len(prev["groups"]) < 2 or not last["groups"]:
            break
        if groups_fill(last["groups"], last.get("scale", 1.0)) > 0.45 * CONTENT_H:
            break
        u = prev["groups"][-1]
        new_prev = prev["groups"][:-1]
        new_last = last["groups"] + [u]
        sp = fit_scale([x for gg in new_prev for x in gg["units"]], MIN_SC)
        sl = fit_scale([x for gg in new_last for x in gg["units"]], MIN_SC)
        if sp is None or sl is None:
            break
        if groups_fill(new_prev, sp) < 0.5 * CONTENT_H:
            break
        prev["groups"], prev["scale"] = new_prev, sp
        last["groups"], last["scale"] = new_last, sl

def group_absorb(gl):
    """整组搬移：稀疏页 <35% 时把后一页第一组上提"""
    i = 0
    guard = 0
    while i < len(gl) - 1 and guard < 60:
        guard += 1
        curp, nxt = gl[i], gl[i + 1]
        if not curp["groups"] or not nxt["groups"]:
            i += 1
            continue
        if groups_fill(curp["groups"], curp.get("scale", 1.0)) > 0.45 * CONTENT_H:
            i += 1
            continue
        u = nxt["groups"][0]
        trial = curp["groups"] + [u]
        sc = fit_scale([x for gg in trial for x in gg["units"]], MIN_SC)
        if sc is None:
            i += 1
            continue
        curp["groups"] = trial
        curp["scale"] = sc
        nxt["groups"] = nxt["groups"][1:]
        if not nxt["groups"]:
            del gl[i + 1]
            continue
        ns = fit_scale([x for gg in nxt["groups"] for x in gg["units"]], MIN_SC)
        if ns is not None:
            nxt["scale"] = ns

def page_fit(gs, min_sc):
    sc = fit_scale([u for gg in gs for u in gg["units"]], min_sc)
    if sc is None:
        return None
    if groups_fill(gs, sc) > CONTENT_H:
        return None
    return sc

def retitle(pg):
    ts = {u.get("gtitle") for gg in pg["groups"] for u in gg["units"]}
    base = ts.pop() if len(ts) == 1 else pg["h1"]
    pg["title"] = base + ("（续）" if pg.get("cont") else "")

# ---------- 打包（整组为最小搬运单位；小页可并入后续大组的拆块首页） ----------
pages = []
cur = []
cur_h1_run = None
held = None
held_h1 = None

def emit(pg):
    if not pg.get("groups"):
        return
    retitle(pg)
    pages.append(pg)

def flush_cur():
    global cur, cur_h1_run, held, held_h1
    if not cur:
        return
    sc = fit_scale([u for g in cur for u in g["units"]], MIN_SC) or MIN_SC
    fill = groups_fill(cur, sc)
    if fill < 0.4 * CONTENT_H:
        if held is not None and held_h1 != cur_h1_run:
            # 异部分的小页先发出，再把当前小页挂起
            emit({"type": "content", "h1": held_h1, "groups": held,
                  "scale": fit_scale([u for gg in held for u in gg["units"]], MIN_SC) or MIN_SC,
                  "cont": 0})
            held = None
            held_h1 = None
        held = (held or []) + cur
        held_h1 = cur_h1_run
    else:
        if held is not None:
            emit({"type": "content", "h1": held_h1, "groups": held,
                  "scale": fit_scale([u for gg in held for u in gg["units"]], MIN_SC) or MIN_SC,
                  "cont": 0})
            held = None
            held_h1 = None
        emit({"type": "content", "h1": cur_h1_run, "groups": cur,
              "scale": sc, "cont": 0})
    cur = []
    cur_h1_run = None

for g in groups:
    if cur and g["h1"] == cur_h1_run:
        if page_fit(cur + [g], MIN_PACK) is not None:
            cur.append(g)
            continue
    flush_cur()
    if held is not None and g["h1"] == held_h1:
        if page_fit(held + [g], MIN_SC) is not None:
            cur = held + [g]
            cur_h1_run = g["h1"]
            held = None
            continue
        if page_fit([g], 0.0) is not None:
            # g 非超大组：先发出 held 小页
            emit({"type": "content", "h1": held_h1, "groups": held,
                  "scale": fit_scale([u for gg in held for u in gg["units"]], MIN_SC) or MIN_SC,
                  "cont": 0})
            held = None
        # g 为超大组：held 保留，拆块时前置合并
    pre = None
    if held is not None and g["h1"] == held_h1:
        pre = held
        held = None
        held_h1 = None
    if page_fit([g], MIN_SC) is not None:
        cur = [g]
        cur_h1_run = g["h1"]
        continue
    chs = chunk_group(g, pre_groups=pre)
    pre = None
    for i, cpg in enumerate(chs):
        cpg["cont"] = i
        cpg["scale"] = fit_scale([u for gg in cpg["groups"] for u in gg["units"]], 0.0) or 0.82
        if i == 0 and held is not None and g["h1"] == held_h1:
            trial = held + cpg["groups"]
            sc = fit_scale([u for gg in trial for u in gg["units"]], 0.0)
            if sc is not None:
                cpg["groups"] = trial
                cpg["scale"] = sc
            else:
                emit({"type": "content", "h1": held_h1, "groups": held,
                      "scale": fit_scale([u for gg in held for u in gg["units"]], MIN_SC) or MIN_SC,
                      "cont": 0})
            held = None
        emit(cpg)
    cur = []
    cur_h1_run = None
flush_cur()
if held is not None:
    emit({"type": "content", "h1": held_h1, "groups": held,
          "scale": fit_scale([u for gg in held for u in gg["units"]], MIN_SC) or MIN_SC,
          "cont": 0})

# 页级整组再平衡（同一 H1 连续内容页）
def post_balance(slides):
    runs = []
    run = []
    for s in slides:
        if s["type"] == "content":
            if run and run[0]["h1"] != s["h1"]:
                runs.append(run)
                run = []
            run.append(s)
        else:
            if run:
                runs.append(run)
            run = []
    if run:
        runs.append(run)
    for gl in runs:
        if len(gl) < 2:
            continue
        group_rebalance(gl)
        group_absorb(gl)
        group_rebalance(gl)
        for s in gl:
            retitle(s)
    # 关键修复（2026-10-22）：group_absorb 的 del 只作用于 run 副本列表，
    # 被吸空的页会以 groups=[] 空壳残留在 slides 主列表（此前靠防御过滤掩盖，
    # 但会让后续手工 override 的 merge_next 撞上空页）。重建主列表使删除生效。
    slides[:] = [s for run in runs for s in run]

post_balance(pages)

# ---------- 手工页级 override（九年级会话新增；列表为空 = 纯自动分页，不影响其他页） ----------
# 每项: {"anchor_block": <Word 块 id>, "action": "...", "scale": <可选显式档位>}
#   merge_next: anchor 所在页与下一页（同 H1）合一页，组列表直接拼接（保留组间距）
#   pull_prev:  anchor 所在单元移到前一页末尾（同 gtitle 并入末组，否则新开组）
#   pull_text:  anchor 为 figure_entry 文字块时仅移文字（题图对拆文，用户指定 2026-10-22）
#   new_page:   anchor 单元另起新页（硬规则 UH 大标题/新章节起新页的实现；其后邻页
#               前部组尽量上提填实避免标题独占空页；absorb_any=True 放宽为任意组，
#               吸收环拟合下限用 op 自带 scale 而非 MIN_SC）
#   scale 缺省 = fit_scale 自动下探（0.82 兜底）；给定则强制采用（布局以渲染目检为准）
# ⚠️ 限额/限高会改变自动打包布局：手工 op 可能被自动合并抢先实现或锚点错位，
# 改 IMG_MAX_H/MANUAL_PAGES 后必须重跑并读 plan 日志核对每个 op 的落点。
# 九年级 v2 定稿的 op 用法示例（anchor 为 Word 块 id，按项目各异）：
#   {"anchor_block": 137, "action": "merge_next", "scale": 0.78}          # 题干页并答案页
#   {"anchor_block": 229, "action": "pull_text"}                          # ②文字上移、图留原页
#   {"anchor_block": 243, "action": "new_page", "absorb_any": True, "scale": 0.82}
MANUAL_PAGES = []   # 按项目填写；空 = 纯自动分页

def _page_units(pg):
    return [u for gg in pg["groups"] for u in gg["units"]]

def _unit_blocks(u):
    if u["kind"] == "stem":
        return set(u["blocks"])
    if u["kind"] == "imgpair":
        s = {im["block"] for im in u["imgs"]}
        s |= {im["caption_block"] for im in u["imgs"] if im.get("caption_block") is not None}
        return s
    if u["kind"] == "figure_entry":
        s = {u["text_unit"]["block"], u["image_unit"]["block"]}
        iu = u["image_unit"]
        if iu.get("caption_block") is not None:
            s.add(iu["caption_block"])
        return s
    return {u.get("block")}

def _locate(slides, block_id):
    """返回 (页 idx, 组 idx, 组内单元 idx)；找不到返回 None"""
    for i, pg in enumerate(slides):
        if pg["type"] != "content":
            continue
        for gi, gg in enumerate(pg["groups"]):
            for ui, u in enumerate(gg["units"]):
                if block_id in _unit_blocks(u):
                    return i, gi, ui
    return None

def _rescale(pg, floor=0.82):
    us = _page_units(pg)
    pg["scale"] = fit_scale(us, MIN_SC) or fit_scale(us, floor) or pg.get("scale", 1.0)

def _append_units(pg, moved):
    """moved 单元挂到 pg 末尾：同 gtitle 并入末组（保留组间距），否则新开一组"""
    last = pg["groups"][-1] if pg["groups"] else None
    if last is not None and last.get("gtitle") == moved[0].get("gtitle"):
        last["units"] = last["units"] + moved
    else:
        pg["groups"].append({"units": moved, "gtitle": moved[0].get("gtitle")})

def _fit_or(units, explicit):
    if explicit:
        return explicit
    return fit_scale(units, MIN_SC) or fit_scale(units, 0.82) or 0.82

def manual_overrides(slides):
    for ov in MANUAL_PAGES:
        loc = _locate(slides, ov["anchor_block"])
        if loc is None:
            print("manual override: anchor block %s not found, skip" % ov["anchor_block"])
            continue
        idx, gi, ui = loc
        act = ov["action"]
        if act == "merge_next":
            if idx + 1 >= len(slides) or slides[idx + 1]["type"] != "content" \
               or slides[idx + 1].get("h1") != slides[idx].get("h1"):
                print("manual override: merge_next block %s precondition failed, skip" % ov["anchor_block"])
                continue
            ua, ub = _page_units(slides[idx]), _page_units(slides[idx + 1])
            import os as _os
            if _os.environ.get("PLAN_DEBUG"):
                bl = lambda us: sorted(b for u in us for b in _unit_blocks(u) if b)
                print("  [dbg] merge anchor %s: idx=%d L=%s R=%s" % (
                    ov["anchor_block"], idx, bl(ua), bl(ub)))
            sc = _fit_or(ua + ub, ov.get("scale"))
            if fit_scale(ua + ub, MIN_SC) is None and not ov.get("scale"):
                print("manual override: estimator rejects merge, force %.2f (render verifies)" % sc)
            slides[idx]["groups"] = slides[idx]["groups"] + slides[idx + 1]["groups"]
            slides[idx]["scale"] = sc
            slides[idx]["cont"] = 0
            retitle(slides[idx])
            del slides[idx + 1]
            print("manual override: merged p%d + p%d -> scale %.2f" % (idx, idx + 1, sc))
        elif act == "pull_prev":
            if idx == 0 or slides[idx - 1]["type"] != "content":
                print("manual override: pull_prev block %s precondition failed, skip" % ov["anchor_block"])
                continue
            src = slides[idx]["groups"][gi]
            u = src["units"].pop(ui)
            if not src["units"]:
                del slides[idx]["groups"][gi]
            moved = [u]
            prev = slides[idx - 1]
            sc_a = _fit_or(_page_units(prev) + moved, ov.get("scale"))
            _append_units(prev, moved)
            prev["scale"] = sc_a
            retitle(prev)
            if not slides[idx]["groups"]:
                del slides[idx]
            else:
                _rescale(slides[idx])
                retitle(slides[idx])
            print("manual override: pulled block %s from p%d to p%d (scale %.2f)"
                  % (ov["anchor_block"], idx + 1, idx, sc_a))
        elif act == "pull_text":
            if idx == 0 or slides[idx - 1]["type"] != "content":
                print("manual override: pull_text block %s precondition failed, skip" % ov["anchor_block"])
                continue
            gg = slides[idx]["groups"][gi]
            u = gg["units"][ui]
            if u["kind"] == "figure_entry":
                moved = [dict(u["text_unit"], gtitle=u.get("gtitle"))]
                gg["units"][ui] = u["image_unit"]   # 图（含图注字段）留在原页
            else:
                moved = [gg["units"].pop(ui)]
            prev = slides[idx - 1]
            sc_a = _fit_or(_page_units(prev) + moved, ov.get("scale"))
            _append_units(prev, moved)
            prev["scale"] = sc_a
            retitle(prev)
            _rescale(slides[idx])
            retitle(slides[idx])
            print("manual override: pulled text of block %s from p%d to p%d (scale %.2f)"
                  % (ov["anchor_block"], idx + 1, idx, sc_a))
        elif act == "new_page":
            pg = slides[idx]
            gg = pg["groups"][gi]
            if gi == 0 and ui == 0:
                print("manual override: block %s already at page start, skip" % ov["anchor_block"])
                continue
            head_groups = pg["groups"][:gi]
            if gg["units"][:ui]:
                head_groups = head_groups + [{"units": gg["units"][:ui],
                                              "gtitle": gg.get("gtitle")}]
            tail_groups = []
            if gg["units"][ui:]:
                tail_groups.append({"units": gg["units"][ui:], "gtitle": gg.get("gtitle")})
            tail_groups = tail_groups + pg["groups"][gi + 1:]
            # 从下一页前部吸收同 gtitle 的组，避免新页只剩大标题一片空白；
            # absorb_any=True 时放宽为任意组（用于节标记"同型压轴真题演练："后接条目）
            j = idx + 1
            while (j < len(slides) and slides[j]["type"] == "content"
                   and slides[j].get("h1") == pg.get("h1") and tail_groups):
                ng_first = slides[j]["groups"][0]
                if not ov.get("absorb_any") and \
                   ng_first.get("gtitle") != tail_groups[-1].get("gtitle"):
                    break
                trial = _page_units({"groups": tail_groups}) + ng_first["units"]
                _floor = ov.get("scale") or MIN_SC
                if fit_scale(trial, _floor) is None:
                    break
                tail_groups.append(ng_first)
                rest = slides[j]["groups"][1:]
                if rest:
                    slides[j]["groups"] = rest
                    _rescale(slides[j])
                    retitle(slides[j])
                    break
                del slides[j]
            new_pg = {"type": "content", "h1": pg["h1"], "groups": tail_groups,
                      "scale": _fit_or(_page_units({"groups": tail_groups}), ov.get("scale")),
                      "cont": 0}
            pg["groups"] = head_groups
            # 头部页保持自身自然档位（不吃操作显式 scale——那只用于新页/吸收）
            pg["scale"] = fit_scale(_page_units(pg), MIN_SC) or \
                fit_scale(_page_units(pg), 0.82) or 1.0
            retitle(pg)
            slides.insert(idx + 1, new_pg)
            retitle(new_pg)
            print("manual override: new page before block %s (p%d), scale %.2f"
                  % (ov["anchor_block"], idx + 2, new_pg["scale"]))
        else:
            print("manual override: unknown action %r, skip" % act)
    return slides

manual_overrides(pages)
pages = [pg for pg in pages if pg.get("groups")]  # 防御：剔除再平衡产生的空页

final_slides = [{"type": "cover", "title": para_text(BLOCKS[0]), "blocks": [0, 1, 2]}]
for pg in pages:
    flat = [u for gg in pg["groups"] for u in gg["units"]]
    final_slides.append({"type": "content", "title": pg["title"], "h1": pg["h1"],
                         "scale": pg.get("scale", 1.0), "cont": pg.get("cont", 0),
                         "units": flat, "groups": pg["groups"]})

# 统计
n_content = sum(1 for s in final_slides if s["type"] == "content")
beats = sum(len(s["units"]) for s in final_slides if s["type"] == "content")
print("slides total:", len(final_slides), "| content:", n_content, "| beats:", beats)
low = []
for i, s in enumerate(final_slides):
    if s["type"] != "content":
        continue
    f = groups_fill(s["groups"], s.get("scale", 1.0))
    if f < 0.4 * CONTENT_H:
        low.append((i + 1, round(f / CONTENT_H * 100), s["title"][:20]))
print("pages fill<40%:", low)

os.makedirs(os.path.join(ROOT, "plan"), exist_ok=True)
with open(os.path.join(ROOT, "plan", "slides.json"), "w", encoding="utf-8") as f:
    json.dump(final_slides, f, ensure_ascii=False, indent=1)
print("written plan/slides.json")
