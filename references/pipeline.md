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

```bash
PPT_DECK_NAME=<交付文件名> python scripts/color_audit.py > out/color_audit.txt
```

脚本 `scripts/color_audit.py`（zipfile+lxml，无第三方依赖）：解包 pptx，逐 slide 逐 a:p 逐 a:r
扫描 rPr/solidFill/srgbClr，输出所有 3E6FC9（蓝）与 D9534F（红）run 的 (页码, 段落文本, run 文本)。
人工核对语义白名单：
- 合法蓝：答案行整行 / 粗体引导词(≤24字到冒号) / 标题行粗体run / 板块行粗体段 / 变式标题行 / 封面标题（含副题）
- 违规：正文编号、括号说明、任何非粗体正文出现蓝色
- 红：只允许出现在"易错点/易错位置"类标签

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
| 题图描述文字/图注公式丢失 | 是否"文字+公式+图"混合段未拆块；图注是否用 caption_runs | EA-38 |
| check_render 报"重叠"但截图目检干净 | 字体接缝伪影（同块同行被拆）；确认已含同块同行合并规则 | EA-42 |

## 5b. 手工页级 override 与图高限额（用户点名"某两页并一页/某图挪页/大标题起新页"时用）

自动分页无法满足用户点名的个别页布局时，用脚本内配置做**页级手术**，其余页保持自动结果：

- `IMG_MAX_H`（plan_slides.py 头部，硬规则 UG）：`{"<media 文件名>": <限高 in>}`——
  est 与 build 同走 img_display_size，等比缩小口径一致；空 = 全部走默认限高 2.2in。
  ⚠️ 文件名以 extract/content.json 的 `rid2target` 实测为准，勿按清单顺序猜（EA-46）。
- `MANUAL_PAGES`（plan_slides.py 头部，九年级 v2 扩展为四动作；在 post_balance 之后执行，结果为最终态）：
  - `merge_next`：anchor 所在页与下一页（同 H1）合一页，组列表拼接（保留组间距）；
  - `pull_prev`：anchor 单元移到前一页末尾（同 gtitle 并入末组，否则新开组）；
  - `pull_text`：figure_entry 拆文——仅文字上移、图（含图注字段）留原页（用于"②题目并上一页"类需求）；
  - `new_page`：anchor 单元另起新页（硬规则 UH 实现），随后从下一页前部吸收组填实；
    `absorb_any: True` 跳过 gtitle 同组限制（节标记后接异组条目时用）；吸收环拟合下限用 op 自带 `scale` 而非 MIN_SC；
    `absorb_pages: N` 限制最多吸收几页（防吸完目标页后继续吞下一节同 gtitle 标题组，EA-54）；
    `cont` 覆盖新页的（续）标记（缺省 0=无（续）；续解析类新页传 1）。
  - `merge_next` / `new_page` / `pull_prev` 均支持 `cont` 覆盖。
  - `pull_prev` 支持 **`img_cap`**（单元级图限高，est 的 img_unit_h 与 build 的 image 分支同口径）——
    手工挪图要缩图时**必须用它**，勿把键加进 IMG_MAX_H（那是全局打包参数，会偏移后续自动布局、
    令全部既有锚点错位，EA-54）。
  - `scale` 缺省自动下探至 0.82，可显式指定档位（低于常规档属用户知情选择，以渲染目检为准）。
- `MANUAL_MERGE_TAIL`（八年级新增）：按组标题把组尾两页合并（在 chunk_group 内生效）。
- **通用纪律**：
  - 改完必须全链重跑（plan→check→build→validate→render→目检），并 diff 新旧 slides.json
    确认**只有点名页变化、其余页逐块一致**（用块签名+SequenceMatcher 比对）；
  - **合并可行性三重裁决**：估算器网格搜索（scale × 图高）找图片最大可保留值 → validate
    0 错误（形状越界即否决）→ check_render 渲染实测；估算器不放行时勿强行硬塞；
  - ⚠️ 限额/限高会改变自动打包布局：手工 op 可能被自动合并**抢先实现**（此时删掉冗余 op）
    或锚点错位（EA-47），每次改参后必须重跑并读 plan 日志核对每个 op 的落点；
  - 限高使大 stem 不再超页时会熔回单单元（拍数减少是良性——题干+图同拍更符合动画规则），
    内容完整性以 check_plan 覆盖率 100% 为准。
- **⛔ 工作流决定（用户 2026-10-23 定）：一切 PPT 修改统一走管线重建（上述配置化机制），不做交付文件的 XML 精修**。
  依据：跨页结构操作的精修在 OMML 重页面会引发渲染幻影（EA-51）；配置已含用户全部点名修改，重建=确定性复现、无覆盖风险；
  结构类修改精修的调试轮次反超 token 收益。执行纪律：重建前备份交付文件；若用户可能在 PowerPoint 手工改过内容，先说明并 diff 合并；
  纯文本类微调确需精修时仅限 EA-51 例外条款且必须 LO 渲染目检防幻影。

## 6. 验收清单（逐项打勾，缺一不交付）

- [ ] check_plan：`dup blocks: []`、覆盖率 100%
- [ ] validate：`errors: 0, warns: 0`（8 项：过渡/时序/页脚/覆盖/公式数/图片比例/字体/色板/越界）
- [ ] 公式 Choice 全部带 xmlns:a14；全包无 `ns0:`
- [ ] 颜色审计：蓝色清单逐条符合语义白名单
- [ ] check_render：`全部干净（无重叠/无越底）`
- [ ] 90dpi 逐页截图逐一目检：颜色语义 / 图题同页 / 双图并排 / 组间空行 / 无遮挡 / 无大面积空白 / 页脚署名
- [ ] shasum 交付副本与 out/ 一致
