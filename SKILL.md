---
name: math-teaching-ppt
description: 初中数学教学 PPT 全自动生成：输入一份考点梳理 Word 稿（.docx），产出逐字镜像内容、原生可编辑 OMML 公式、逐拍点击动画、版式合规的教学 PPT（.pptx）。触发词：教学PPT、Word转PPT、数学课件、根据word做ppt、期中/期末复习课件、teaching ppt。
---

# math-teaching-ppt —— 从 Word 考点梳理稿生成教学 PPT

## 目标与产出

输入：一份初中数学考点梳理 Word 稿（结构：标题层级 H1/H2/H4、典型原题、答案行、解析段、表格、题图、图注）。
产出：16:9 教学 PPT——内容与 Word **逐字一致**、477 级公式全部原生可编辑 OMML、每页 fade 过渡 + 逐拍点击渐显动画、每页页脚署名、字号/色板/字体统一、无遮挡无大面积空白。

## 硬规则速查（违反任意一条即返工；详细判定与实现见 references/format-rules.md）

1. **逐字镜像**：Word 每个非空段（含纯公式段、纯图片段、表格单元格）必须出现在 PPT 且文本逐字一致。校验：`validate.py` D 项覆盖率 100%。
2. **公式 = 原生 OMML**：从 Word 原样提取，以 `mc:AlternateContent > mc:Choice(Requires=a14, xmlns:a14 声明在本元素) + mc:Fallback` 嵌入。校验：validate E 项公式数一致 + 每个 Choice 带 xmlns:a14。
3. **原子组不拆**：板块/整道题（题干+选项+图+答案+解析+易错点）/变式题是不可拆组，组放不下整组移下一页；仅当组超过一整页时在"题目|解析"边界拆。校验：`check_plan.py` 覆盖 100% 无重复 + 目检。
4. **无遮挡**：含分式/根号/求和结构的公式段行高按 1.95 倍估（真渲染撑行），仅上标/线性段用 1.58 倍；**严禁一刀切加大余量**（会全片稀疏、图题分离）。
5. **动画逐拍**：每页 fade 250ms；题干（含选项/题图）第一拍，答案、解析逐段、拆解表、易错点各一拍。校验：validate B 项拍数=计划。
6. **格式统一**：微软雅黑单一字族；色板白名单（标题 1A237E/正文 1A2230/答案·引导词蓝 3E6FC9/警示红 D9534F/图注 5A6472/页脚灰 8A93A6）。**蓝色语义**仅限：答案行整行、粗体引导词（到冒号）、标题/变式标题行、板块行粗体段、封面标题——正文编号与括号说明必须黑色。交付前跑 run 级颜色审计（见 pipeline.md Phase 5）。
7. **每页页脚署名**：`PPT_FOOTER`（默认"Oli教室专用"）+ 页码/总页数，右下 10pt 灰。校验：validate C 项。
8. **图片**：按原图宽高比等比缩放（校验 ±1%），限高 2.2in 放大≤1.6 倍；**相邻两张题图左右并排**（imgpair，限宽 5.9in），不上下堆叠。
9. **同页多组空行**：同页不同题/组之间留 GROUP_GAP=0.42in 空行。
10. **交付前必须逐页截图目检** + 全部自动化校验 0 错误，二者缺一不可（错误档案 EA-2/EA-8 教训）。
11. **同题同页可缩图合并（UG）**：同一道题尽量**题目和图片在同一张 PPT 上**；放不下时优先等比缩小题图（`IMG_MAX_H` 按文件名限高）直至容纳，用估算器网格搜索取图片最大可保留值；**实在放不下才允许例外**拆两页（题目|解析边界）。⚠️ 限高文件名以 rid2target 实测为准，改后重读 plan 日志（EA-46/47）。
12. **大标题/新章节/新段落新开页（UH）**："压轴二（几何探究型）——代表题：…""同型压轴真题演练："之类的大标题/新章节/新段落**必须新开一页 PPT 写，不要跟在前面内容的后边**；新页从后续内容吸收填实、不得独占空页（UA 的精细化）。

## 六阶段流程（命令均可直接执行；环境变量见 references/pipeline.md）

### Phase 0 建工程目录
中间产物一律进独立工程目录，不污染交付目录：
```
mkdir -p <工程目录>/scripts && cp <skill>/scripts/*.py <工程目录>/scripts/
```

### Phase 1 提取（Word → content.json）
```
python scripts/extract_docx.py <word稿路径>      # 输出 extract/content.json + extract/media/
python scripts/preview_content.py                # 生成 extract/preview.txt
```
**必读** `extract/preview.txt` 开头的样式映射与标题大纲：确认题干/答案/解析/图注/表格形态与本 skill 预期一致（典型原题加粗、答案行"答案："开头、H1/H2/H4 层级）。形态异常先改 extract 的样式识别，再进 Phase 2。

### Phase 2 分页计划
```
python scripts/plan_slides.py     # 产出 plan/slides.json + plan/plan_preview.txt
python scripts/check_plan.py      # 必须输出：dup blocks: []；覆盖率 100%
```
通读 plan_preview.txt 检查：每道题完整成组、超长组拆分点在"题目|解析"边界、表格拆页带表头。

### Phase 3 构建
```
python scripts/build_pptx.py      # 产出 out/<PPT_DECK_NAME>；交付字体默认微软雅黑
```

### Phase 4 自动化校验（全绿才可进入 Phase 5）
```
python scripts/validate.py        # 8 项硬校验必须 0 错误
```
另跑 **run 级颜色审计**（脚本见 references/pipeline.md §颜色审计）：输出全片蓝色 run 清单，逐条核对符合硬规则 6 的蓝色语义；正文编号/括号说明出现蓝色即为 bug。

### Phase 5 视觉验收（不可省略）
```
soffice --headless --convert-to pdf --outdir <dir> <pptx>   # 或 /Applications/LibreOffice.app/Contents/MacOS/soffice
python scripts/check_render.py <pdf>                        # 文本重叠/越底检测，必须 0 问题
```
然后 **90dpi 逐页出 PNG 截图并一页一页目检**（重点：颜色语义、图题同页、双图并排、组间空行、无遮挡、无大面积空白）。两点环境事实：
- 本机 LibreOffice 无微软雅黑且不自动 CJK 回退（渲染成豆腐/替换成无关字体）：把 `Hiragino Sans GB.ttc` 软链进 LibreOffice 字体目录后用 `PPT_FONT="Hiragino Sans GB"` 构建验证版（交付文件字体不变）。
- LO **只渲染公式 Fallback 线性文本**（如 `x^2`），截图中的线性文本是兜底非最终效果；真公式需 PowerPoint/WPS 查看。这决定了 bbox 检测对"真公式撑行遮挡"是盲区——防线是 Phase 4 前的分页估算余量（硬规则 4）。

### Phase 6 交付
拷贝 pptx 到交付目录，`shasum` 双份核对一致；更新记忆日志。

## 参数约定（build 与 validate 必须同口径）

| 环境变量 | 默认 | 说明 |
|---|---|---|
| `PPT_TITLE` | Word 文档标题块 | 封面大标题（40pt 蓝粗，无装饰条） |
| `PPT_FOOTER` | `Oli教室专用` | 每页右下页脚署名 |
| `PPT_DECK_NAME` | `teaching_ppt.pptx` | out/ 下默认输出文件名 |
| `PPT_FONT` | `微软雅黑` | 全片字族；渲染验证时临时换本机有的中文字体 |
| `PPT_OUT` | `out/<PPT_DECK_NAME>` | 输出路径 |
| `PPT_CALIB` | `1.0` | 分页高度全局校准系数（渲染反馈迭代用） |

## 协作约定

- 定稿 pptx 只读：用户认可的成品是模板基准，任何情况下不得修改；改版式=改脚本常量后重建。
- 分页参数调整（字号/容量/档位）集中在 `plan_slides.py` 头部常量；改完必须全链重跑（plan→check→build→validate→render→目检）。
- 版式/结构细节见 `references/design-spec.md`；管线细节与排查手册见 `references/pipeline.md`；所有历史错误与修复见 `references/error-archive.md`（新错误必须追加，保留用户反馈原话）。
- 参考样例：
  - `assets/九年级期中考点梳理教学课件_定稿参考.pptx`（68 页**定稿**，2026-10-22 用户逐页调整后冻结，**只读禁改**）——同题同页缩图合并（UG）与大标题新开页（UH）的样式基准，含 3.3 陷阱题整题页/同型压轴演练/压轴二新开页等全部页级形态；
  - `assets/七年级期中考点梳理教学课件_参考样例.pptx`（52 页定稿）+ `assets/七年级工程设计与结构记录.md` + `assets/七年级分页计划样例.txt`。
