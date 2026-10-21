# math-teaching-ppt

从一份初中数学考点梳理 Word 稿（.docx）全自动生成高质量教学 PPT（.pptx）的完整管线与 Skill。

## 产出特性

- **逐字镜像**：Word 全部内容（文本/表格/题图/图注）按文档顺序进入 PPT，程序化校验覆盖率 100%，无改写无遗漏。
- **原生可编辑公式**：从 Word 提取的 OMML 以 PowerPoint 原生形态嵌入（`mc:AlternateContent` 包裹，`xmlns:a14` 声明在 `mc:Choice` 上），PowerPoint/WPS 中双击可编辑，非图片。
- **逐拍教学动画**：每页 fade 过渡 250ms；题干（含选项/题图）第一拍，答案、解析逐段、拆解表逐项、易错点各一拍点击渐显。
- **原子组分页**：板块/整道题/变式题是不可拆组，组放不下整组移下一页；组超页时只在"题目|解析"边界拆分；同页多组之间留空行；相邻双题图自动左右并排。
- **统一格式**：单一字族、色板白名单、每页页脚署名 + 页码。
- **三层交付自检**：8 项程序校验 + 渲染重叠/越底检测 + 逐页截图目检。

## 快速开始

```bash
# 0) 建工程目录（中间产物不污染交付目录）
mkdir -p my_ppt_work/scripts && cp scripts/*.py my_ppt_work/scripts/

# 1) 提取（Word → content.json）
cd my_ppt_work
python scripts/extract_docx.py /path/to/考点梳理报告.docx
python scripts/preview_content.py          # 人工核对解析形态

# 2) 分页计划（覆盖 100% + 无重复必须通过）
python scripts/plan_slides.py
python scripts/check_plan.py

# 3) 构建（封面标题/页脚署名/输出名可配）
PPT_TITLE="上海市X年级上学期期中数学复习课" python scripts/build_pptx.py

# 4) 校验（必须 0 错误）
PPT_TITLE="..." python scripts/validate.py

# 5) 渲染验收
soffice --headless --convert-to pdf --outdir . out/teaching_ppt.pptx
python scripts/check_render.py teaching_ppt.pdf
# 最后：逐页导出 PNG 截图人工目检（颜色语义/图题同页/空行/遮挡）
```

环境变量：`PPT_TITLE`（封面标题，默认取 Word 文档标题块）、`PPT_FOOTER`（页脚署名，默认 `Oli教室专用`）、`PPT_DECK_NAME`（输出文件名）、`PPT_FONT`（交付字体，默认微软雅黑）、`PPT_OUT`、`PPT_CALIB`（分页高度校准）。

## 仓库结构

```
SKILL.md                     # Skill 主文件：六阶段流程 + 硬规则速查
references/
  format-rules.md            # 格式硬规则（用户要求逐条展开，可执行可验证）
  error-archive.md           # 错误档案：34 条踩坑记录（错误 → 用户反馈原话 → 正确做法）
  pipeline.md                # 管线详解 + 颜色审计方法 + 症状排查表 + 交付验收清单
  design-spec.md             # 版式规格（几何/字号/色板/动画/公式形态/分页常量）
scripts/                     # 七个管线脚本（参数化，含适配指南）
assets/                      # 参考样例：52 页定稿 PPT + 工程设计记录 + 分页计划样例
```

## 依赖

Python ≥3.11：`python-pptx`、`lxml`、`pymupdf`、`Pillow`；LibreOffice（仅渲染验收）。

## 关键工程经验（详见 error-archive.md）

- 公式嵌入的两个致命坑：`xmlns:a14` 必须声明在 `mc:Choice` 上（否则 PowerPoint 全部回落兜底文本）；lxml 建 mc 元素必须传 `nsmap`（否则前缀变 `ns0:`）。
- 防公式撑行遮挡的余量必须**按公式结构分层**：只有含分式/根号/求和的段落加大行高，仅上标/线性段落用常规密度——一刀切会把全片撑稀、图题分离。
- LibreOffice 只渲染公式兜底文本且无中文字体回退，渲染验收前需软链 CJK 字体并理解 QA 截图的局限。

## 参考样例

`assets/七年级期中考点梳理教学课件_参考样例.pptx`：52 页（1 封面 + 51 内容页）、477 个可编辑公式、8 图、252 拍动画，由 13 所学校/区 15 份期中卷整理的 Word 考点梳理稿生成。

## License

MIT
