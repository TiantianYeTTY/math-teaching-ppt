# math-teaching-ppt

> 面向初中数学考点梳理场景的 **Word → 教学 PPT 全自动生成技能**（Agent Skill）。
>
> 从解析考点梳理 Word 稿，到**逐字镜像**全部内容、嵌入**原生可编辑 OMML 公式**（非图片）、程序化分页防溢出防空白，再到**逐拍点击动画**与**三层交付自检**——一条管线产出可直接用于课堂教学的高质量课件。
>
> 作者：**TiantianYe**
>
> 作者在 App Store 中的其他代表作品：水课帮-课堂提问助手；识词本-单词笔记；投影与视图3D教具；解构照片

---

## ✨ 这是什么

`math-teaching-ppt` 是一个 Agent Skill（智能体技能包），把一次完整的数学教学课件制作流程沉淀为可复用的工作流与踩坑经验。它解决的问题非常具体：

- 手工做数学课件时，公式往往被**截图或手打文本**替代，既不能后期编辑、也和 Word 稿不一致；
- 课件内容容易对 Word 稿**改写、提炼或遗漏**，排版上则常常出现大面积空白、内容挤压、同一道题被拦腰拆到两页；
- 课堂教学需要**逐拍点击渐显**的动画（先出题目、再逐步出答案），手工设置 200+ 个动画点既繁琐又易漏。

本技能将整个流程标准化：**Word 原样提取 → 高度估算驱动的原子组分页 → OMML 原生嵌入 + 逐拍动画 → 三层交付自检**，并对过程中的 **30 个易错点**给出了根因分析与标准解法（见 [references/error-archive.md](references/error-archive.md)）。

## 🎯 使用场景

当你的任务符合以下画像时，即可使用本技能：

| 要素 | 说明 |
|---|---|
| 年级/学科 | 初中数学（沪教版等；其他年级/版本仅需提供对应 Word 稿） |
| 输入材料 | 一份考点梳理 Word 稿（.docx，含标题层级、典型原题、答案、解析、表格、题图） |
| 典型内容 | 知识体系总览、选择题/填空题/简答题专项、压轴题拆解、跨卷母题、易错点清单、变式题库 |
| 用户诉求 | "公式全部 LaTeX 可编辑、Word 内容逐字照搬严禁改写、排版清晰无空白无遮挡、字号适中、格式统一、每页署名、每页过渡动画 + 每题逐拍渐显、同一内容不中途换页、稀疏页 2~3 题并页" |

典型触发指令示例：

> 这是《X 年级期中考点梳理报告.docx》，请根据这份 word 稿从零制作一份对应的教学 PPT，我需要用作课堂教学。要求：1、所有字母和数学公式都必须使用 LaTeX（可后期编辑，不能是钉死的图片）；2、ppt 所有内容完全来自于 word 稿，全部信息原样照搬、严禁改写或遗漏；3、排版清晰易读、不要大面积空白也不要挤在一起、字号适中；4、图片按比例缩小；5、字体颜色统一；6、每页加上"Oli教室"字样；7、每页加过渡动画，每道题先出题目、再逐步出现答案，严禁一下子显示全部；8、同一内容不要中途换页，一页放不下放第二页；内容少的页可 2~3 题并为一页。

## 📦 安装

将本仓库克隆（或复制 `math-teaching-ppt/` 文件夹）到 Agent 技能目录：

```bash
git clone https://github.com/TiantianYeTTY/math-teaching-ppt.git ~/.workbuddy/skills/math-teaching-ppt
```

重启会话后技能即被识别。之后制作教学课件时，用上面示例中的方式描述任务（附上 Word 稿），智能体会自动加载本技能；也可以显式调用 `/math-teaching-ppt`。

## 🚀 使用方法（工作流概览）

技能按以下六个阶段自动执行，完整细节与全部命令见 [SKILL.md](SKILL.md)：

| 阶段 | 内容 | 关键点 |
|---|---|---|
| 1. 提取 | `extract_docx.py` 按文档顺序提取文本 run / OMML / 图片 / 表格 → content.json | 纯公式段、纯图片段同样入块 |
| 2. 预览核对 | `preview_content.py` 输出样式映射与全文预览，人工核对解析形态 | 形态异常先修提取再进分页 |
| 3. 分页计划 | `plan_slides.py` 原子组切分（板块/整题/变式不拆）→ 高度估算打包 → 再平衡 | 组放不下整组移页；行高按公式结构分层 |
| 4. 计划校验 | `check_plan.py` Word 非空块覆盖率 100%、无重复 | 必须通过才可构建 |
| 5. 构建 | `build_pptx.py` 手写 DrawingML：公式/表格/逐拍 timing/fade 过渡 | OMML 以 mc:AC 原生形态嵌入 |
| 6. 验收交付 | `validate.py` 8 项硬校验 + `check_render.py` 重叠检测 + **逐页截图目检** | 三层全绿才允许交付 |

### 输入 / 产出

```
输入                              产出
─────────────────────────        ─────────────────────────────
考点梳理报告.docx         ───▶    教学 PPT .pptx（逐字镜像 + 可编辑
（标题/题目/答案/解析/表格/ ───▶    OMML 公式 + 逐拍动画 + 每页署名）
 题图）                            工程中间产物（extract/plan/out）
```

参考样例：**52 页七年级期中考点梳理课件**（477 个可编辑公式、8 图、252 拍动画）见 [assets/](assets/)。

## 🔧 环境依赖

- **Python ≥3.11**：`python-pptx`、`lxml`、`pymupdf`、`Pillow`（建议专用 venv）
- **LibreOffice**（soffice，仅用于渲染验收；交付文件不依赖）
- **微软雅黑**（交付字体，Windows 教室原生支持；本机 LO 渲染验收可用 `PPT_FONT` 临时换字体，需把 CJK 字体软链进 LibreOffice 字体目录，见 SKILL.md）
- 公式最终效果以 **PowerPoint / WPS** 打开为准（LibreOffice 只渲染公式兜底文本）

## ⚠️ 重要说明

1. **Word 稿零改动**：内容源只读，课件内容与 Word 稿逐字一致（程序校验覆盖率 100%），严禁改写、提炼或遗漏。
2. **公式可编辑性是硬约束**：全部公式为原生 OMML（PowerPoint/WPS 中双击可编辑），以 `mc:AlternateContent` 原生形态嵌入且 `xmlns:a14` 声明在 `mc:Choice` 上——否则严格解析器会全部回落为线性兜底文本。
3. **原子组不拆**：板块/整道题/变式题是不可拆组，组放不下整组移下一页；防遮挡的行高余量只给含分式/根号/求和结构的段落，**一刀切加大余量会把全片撑稀、图题分离**。
4. **三层自检缺一不可**：程序校验 0 错误 + 渲染重叠/越底检测 + 逐页截图人工目检，任何一层发现问题不得交付。

## 📁 目录结构

```
math-teaching-ppt/
├── README.md                  ← 本文件
├── SKILL.md                   ← 技能正文：六阶段流程 + 10 条硬规则速查 + 参数约定
├── references/
│   ├── format-rules.md        ← 格式硬规则（用户要求逐条展开，可执行可验证）
│   ├── error-archive.md       ← 30 条错误档案（错误 → 用户反馈原话 → 正确做法）
│   ├── pipeline.md            ← 管线详解 + 颜色审计 + 症状排查表 + 交付验收清单
│   └── design-spec.md         ← 版式规格（几何/字号/色板/动画/公式形态/分页常量）
├── scripts/                   ← 7 个参数化管线脚本（含适配指南注释）
│   ├── extract_docx.py        ← Word 按文档序提取（文本/OMML/图/表）
│   ├── preview_content.py     ← 提取结果可读预览
│   ├── plan_slides.py         ← 原子组分页计划（高度估算 + 再平衡）
│   ├── check_plan.py          ← 覆盖率 100% + 无重复校验
│   ├── build_pptx.py          ← pptx 构建（公式/表格/动画/过渡）
│   ├── validate.py            ← 8 项硬校验
│   └── check_render.py        ← 渲染后重叠/越底检测
└── assets/                    ← 参考样例（52 页定稿 PPT + 设计记录 + 分页计划样例）
```

## 🐛 已知踩坑点（节选）

[references/error-archive.md](references/error-archive.md) 中记录了完整的 30 条「错误 → 用户反馈原话 → 正确做法」，包括：

- 公式回落为线性文本的两个致命坑：`xmlns:a14` 必须声明在 `mc:Choice` 上（MCE 作用域规则）；lxml 建 mc 元素必须传 `nsmap`（否则前缀变 `ns0:`）；
- 裸 `a14:m` 会被 LibreOffice 整段丢弃，必须用 `mc:AlternateContent` 包裹；
- python-pptx 文本框自带空段落不删会导致每形状多占一行、全局下坠溢出；
- 页级再平衡若按"单元"搬移会把板块标题与内容拆散——必须整组搬移且不跨标题上下文；
- 相邻双题图左右并排（imgpair）而非上下堆叠，连续多图不吸入题干；
- 颜色语义问题靠 run 级审计而非截图目检（90dpi 下小面积误色肉眼难辨）；
- ……

## 🛠️ 可复用脚本

`scripts/` 内七个脚本均为真实课件全流程实测产物，参数化后可直接运行（详见 [SKILL.md](SKILL.md)）：

```bash
# ① 提取（Word → content.json + media/）
python3 scripts/extract_docx.py /abs/考点梳理报告.docx

# ② 分页计划（覆盖校验必须通过）
python3 scripts/plan_slides.py && python3 scripts/check_plan.py

# ③ 构建（封面标题/页脚署名/输出名可配）
PPT_TITLE="上海市X年级上学期期中数学复习课" python3 scripts/build_pptx.py

# ④ 校验 + 渲染验收（与 build 同口径传 PPT_TITLE 等）
PPT_TITLE="上海市X年级上学期期中数学复习课" python3 scripts/validate.py
soffice --headless --convert-to pdf --outdir . out/teaching_ppt.pptx
python3 scripts/check_render.py teaching_ppt.pdf
```

## 👤 作者

**TiantianYe**

作者在 App Store 中的其他代表作品：水课帮-课堂提问助手；识词本-单词笔记；投影与视图3D教具；解构照片。

欢迎通过 Issue 反馈问题，或提交 PR 完善工作流（例如适配其他学科版式的 Word 稿、补充 Windows 环境的渲染验收管线）。

## 📄 许可

仅供学习交流使用。
