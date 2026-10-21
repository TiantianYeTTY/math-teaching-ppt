# 版式规格（v6 定稿参数 · 与脚本常量一一对应）

## 画布与几何

| 项 | 值 |
|---|---|
| 幻灯片 | 16:9，12192000×6858000 EMU（13.33×7.5in） |
| 内容区 | x=0.5in，宽 12.33in，顶 y=1.14in，底边 ≤7.00in |
| 标题条 | 竖条 x=0.5 y=0.38（0.09×0.5in 蓝）；标题文本 x=0.72 y=0.26（12.1×0.74in）；分隔线 y=1.02 |
| 页脚 | x=9.53 y=7.06（3.3×0.32in），右对齐 10pt 灰 |
| 页边距 | 文本框 lIns/rIns=54864 EMU，tIns/bIns=27432 EMU |
| 组间空行 | GROUP_GAP=0.42in（同页不同题/组之间） |
| 单元间距 | GAP=0.12in |

## 字号与行距

| 项 | 值 |
|---|---|
| 正文 | 16pt（缩放档 [1.0, 0.94, 0.9, 0.86, 0.82]；页内追加组最低 0.9 档，整组/拆块最低 0.86 档） |
| 表格 | 12pt |
| 页标题 | 22pt（宽度超限自动 20/18pt） |
| 封面标题 | 40pt 蓝粗居中（无装饰条）；副题 16pt；引导句 15pt |
| 图注 | 12pt 灰居中 |
| 行距 | 115%；段后距 3pt |
| 行高估算 | 纯文本 1.45 倍 / 含上标·线性公式 1.58 倍 / 含 m:f·m:rad·m:nary 1.95 倍 × CALIB |

## 色板白名单（唯一允许值）

| 语义 | HEX |
|---|---|
| 标题 | 1A237E |
| 正文 | 1A2230 |
| 答案·引导词（蓝） | 3E6FC9 |
| 易错警示（红） | D9534F |
| 图注 | 5A6472 |
| 页脚 | 8A93A6 |
| 边线 | D6DCE5 |
| 表头底 | F5F7FA |
| 隔行底 | F7F9FC |
| 白 | FFFFFF |

蓝色语义白名单：答案行整行 / 粗体引导词（≤24 字到冒号）/ 标题行·变式标题行粗体 run / 板块行粗体段 / 封面标题。

## 公式嵌入形态（PowerPoint 原生写法）

```xml
<a:p>
  <a:r>…正文…</a:r>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
               xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" Requires="a14">
      <a14:m><m:oMath><!-- Word 原样 OMML --></m:oMath></a14:m>
    </mc:Choice>
    <mc:Fallback><a:r>…线性化文本…</a:r></mc:Fallback>
  </mc:AlternateContent>
</a:p>
```

- xmlns:a14 / xmlns:m 必须声明在 mc:Choice 开标签（MCE 作用域规则）；mc 元素创建必须传 nsmap。
- 公式行内/显示两态：disp 段用 m:oMathPara，行内用 m:oMath。
- 渲染矩阵：PowerPoint/WPS 走 Choice（可编辑公式）；LibreOffice 只渲染 Fallback 线性文本。

## 动画

- 每页过渡：`p:transition`（mc:AC 包裹）`p:fade`，p14:dur=250ms，Fallback spd=med。
- 逐拍：`p:timing` 主序列 `p:seq`，每拍一个点击 `p:par`（presetID=10 淡入 300ms + set visibility）。
- 拍划分：题干（含选项/题图/图注/子问）第一拍；答案/解析逐段/拆解表逐项/易错点各一拍；标题与页脚静态。

## 分页引擎（plan_slides.py 常量）

| 常量 | 值 | 说明 |
|---|---|---|
| CONTENT_W | 12.33in | 内容宽 |
| CONTENT_H | 5.60in | 规划容量（放置区到 7.00in，余量吸收公式渲染高度差） |
| BODY_FONT / TABLE_FONT | 16 / 12pt | 基准字号 |
| GAP | 0.12in | 单元间距 |
| GROUP_GAP | 0.42in | 组间空行 |
| SCALES | [1.0, 0.94, 0.9, 0.86, 0.82] | 字号缩放档 |
| MIN_PACK | 0.9 | 页内追加组最低档 |
| MIN_SC | 0.86 | 整组放置/拆块最低档 |

规则：组=原子（板块/整题/变式）；追加组前 page_fit（fit_scale+groups_fill 复核）；题干超档按块拆流入；表格均衡分行块（续页带表头）；再平衡整组搬移、不跨 gtitle；emit 过滤空组页。
