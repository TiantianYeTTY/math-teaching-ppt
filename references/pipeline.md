# 管线详解与排查手册

## 0. 依赖

- Python ≥3.11 + `python-pptx`、`lxml`、`pymupdf`、`Pillow`（venv 安装，勿污染全局）
- LibreOffice（`soffice`，渲染验收用；交付不依赖）
- 字体验收：见 SKILL.md Phase 5 环境事实

## 1. 数据流

```
Word(.docx)
  │ extract_docx.py <docx>
  ▼
extract/content.json  ←  按文档顺序的 blocks[]：para（runs: 文本/OMML XML/图片 rId）/ table
extract/media/        ←  图片原文件
  │ preview_content.py → extract/preview.txt（人工核对解析形态）
  │ plan_slides.py     → plan/slides.json（单元切分→拍分组→打包→再平衡）
  │ check_plan.py      → 覆盖率 100% + 无重复
  ▼
plan/slides.json
  │ build_pptx.py（python-pptx + lxml 手写 XML：公式/表格/timing/transition）
  ▼
out/<PPT_DECK_NAME>.pptx
  │ validate.py（8 项硬校验）+ 颜色审计（§3）
  │ soffice → pdf → check_render.py + 逐页 PNG 目检
  ▼
交付目录
```

## 2. content.json 关键字段

- `blocks[]`: `kind: para|table`；para 的 `runs[]`：`{"t","b"(粗体),"va"(上下标)}` / `{"math": "<m:oMath>...</m:oMath>", "disp"(显示公式)}` / `{"img":{"rid","cx","cy"(EMU 显示尺寸)}}` / `{"br":true}`
- `blocks[].style`: Word 段落样式 ID（"1"=H1、"2"=H2、"4"=H4、ImageCaption=图注…）
- `blocks[].num/numid`: 项目符号列表
- `rid2target`: rId → media 文件名；`img_info`: 文件 → 像素尺寸
- 排版语义分层（plan_slides）：`板块`开头=板块行、`变式`=变式标题行、`答案`开头=答案行、`易错`/`关键突破口`等粗体到冒号=引导词

## 3. run 级颜色审计（交付前必跑）

```python
# 解包 pptx，逐 slide 逐 a:p 逐 a:r 扫描 rPr/solidFill/srgbClr
# 输出所有 3E6FC9（蓝）run 的 (页码, 段落文本, run 文本)，人工核对语义白名单：
#   合法：答案行整行 / 粗体引导词(≤24字到冒号) / 标题行粗体run / 板块行粗体段 / 变式标题行 / 封面标题
#   违规：正文编号、括号说明、任何非粗体正文出现蓝色
# 同时核对 D9534F（红）只出现在"易错点："类标签。
# 完整脚本见 GitHub 仓库或按此描述现场编写（约 40 行，无第三方依赖，zipfile+lxml 即可）。
```

## 4. 环境变量速查

见 SKILL.md「参数约定」。**build 与 validate 必须同口径**（PPT_TITLE/PPT_FOOTER/PPT_DECK_NAME 三处联动），否则校验误报。

## 5. 常见症状 → 排查

| 症状 | 首查 | 关联错误 |
|---|---|---|
| PowerPoint 打开公式全是 `x^2` 线性文本 | mc:Choice 是否带 `xmlns:a14`；全包 grep `ns0:` | EA-2, EA-3 |
| LO 渲染整页空白 | 是否存在裸 `a14:m`（未用 mc:AC 包裹） | EA-1 |
| 内容互相遮挡（仅 PowerPoint/腾讯文档可见） | 含分式/根号段落的行高档位；规划容量 | EA-21, EA-22 |
| 页面整体下坠溢出 | add_textbox 是否删了默认空 a:p | EA-30 |
| 某板块标题与内容分离 | 再平衡是否整组搬移 | EA-12 |
| 图和题目分离/孤图页 | 题干吸收规则、图片限高、多图不吸收 | EA-13 |
| 页数骤降/单页超载 | page_fit 是否真拟合（不是把档位当缩放传参） | EA-15 |
| validate 覆盖率缺失 | 纯公式/纯图片段是否入块；封面标题是否同源豁免 | EA-9, EA-34 |
| 计划页数异常多（稀疏） | 行高余量是否被一刀切 | EA-22 |
| 颜色"莫名其妙" | run 级颜色审计清单 | EA-23, EA-24 |

## 6. 验收清单（逐项打勾，缺一不交付）

- [ ] check_plan：`dup blocks: []`、覆盖率 100%
- [ ] validate：`errors: 0, warns: 0`（8 项：过渡/时序/页脚/覆盖/公式数/图片比例/字体/色板/越界）
- [ ] 公式 Choice 全部带 xmlns:a14；全包无 `ns0:`
- [ ] 颜色审计：蓝色清单逐条符合语义白名单
- [ ] check_render：`全部干净（无重叠/无越底）`
- [ ] 90dpi 逐页截图逐一目检：颜色语义 / 图题同页 / 双图并排 / 组间空行 / 无遮挡 / 无大面积空白 / 页脚署名
- [ ] shasum 交付副本与 out/ 一致
