# 学位论文排版与交稿核查：常见工具对比与适用场景

学校发下来的 Word 模板通常已经放好了封面、声明页和校徽。写第一稿时，学生希望把这些内容带进新文档；几轮修改之后，真正头疼的问题变成了：**正文的页码还是从 1 开始吗、中间插入的横向宽表会不会被拉歪、自己手动续写时 Word 顶部样式还能不能直接点选复用、基因斜体/酶名空格/单位格式和 Nature 文献多余 DOI 有没有漏网？**

## 已有工具已经做得很好的部分

下表按真实工作场景梳理常见工具（功能依据各项目公开说明与实测对比）。对能直接用成熟轮子解决的问题，建议直接采用现有工具；本插件专注补齐生科长稿的空白：

| 手头的材料和目标 | 已有成熟方案 | 它们已经做得很好，以及留给本插件的分工 |
| --- | --- | --- |
| 从头用源码写论文，或选定本校现成模板起稿 | `ThuThesis`（LaTeX）、`Quarto`、`Pandoc --reference-doc`、`docxtpl`、校内研究生院模板 | 起稿与样式锁定非常成熟（`docxtpl` 等效生成可达 12/12）。如果你习惯源码写作，直接用它们即可；本插件面向**拿到手还要在 Word 里继续人工改稿**的作者。[Pandoc 手册](https://pandoc.org/MANUAL.html#option--reference-doc) · [docxtpl](https://docxtpl.readthedocs.io/en/latest/) |
| 普通单节 Word 旧稿照模板重排或做常规审计 | `format-thesis-docx`、`DocFormatTool` | 在双方共有的常规页设置与正文样式上判断准确（历史共有规则 16/16 全对）；单节样稿上 `format-thesis-docx` 能扫出更多通用样式项。但在模板未单列第 4、5 节或含生科细节规则时会漏过。[format-thesis-docx](https://github.com/wksudud/format-thesis-docx/tree/c04f786d5b994879f938325630e9ffdf4dde00be) · [DocFormatTool](https://github.com/NeoHotaru/DocFormatTool/blob/40f22ca37f4b856ad6d59a3382c1570e5d7ca67d/README.md) |
| 已有长论文局部修版与分节重置 | `standardize-docx-format`、`format-agent-skill` | 前者支持规则提取、差异预览与勾选修复；后者支持重建前置页与页码，但在文档中明确说明遇到横向宽表等复杂分节时默认停手保护。本插件补充**按角色跨节对齐且保住横向宽表**。[standardize-docx-format](https://github.com/zwt0204/standardize-docx-format/tree/1239c9e24360994879f938325630e9ffdf4dde00be) · [format-agent-skill](https://github.com/KaguraNanaga/format-agent-skill/blob/8dd726caba4395339ba5fda3c209fe1032627de6/references/thesis-mode.md) |
| 特定高校终稿 PDF 视觉复核 | 清华论文检查器（`tsinghua-thesis-format-checker-skill`） | 结合 DOCX 结构与 PDF 版面页码做 48 项自动 + 8 项二阶段复核，适合明确校规的终稿复查。本插件不重写全套 PDF 渲染引擎，而是前置把 **Word 样式底座与 67 项生科体例硬伤**一次清干净。[清华检查器](https://github.com/jacky-rui/tsinghua-thesis-format-checker-skill/blob/main/README.zh-CN.md) |

## 本插件真正拉开差距的三件事

1. **生科毕业论文高频“暗坑”自查表（67 项内置清单 + 自动体检 + 可导出 Word 自查表）**
   - 把生科实验与论文中最容易反复返工的细节做成结构化自查与自动检查：人鼠基因大小写斜体与蛋白正体（*TP53* / *Tp53* vs `TP53`，禁用 `X`/`Y` 指代）、限制酶罗马数字前空格（*EcoR I*、*BamH I*、*Taq I*）、四种线/高能键区分（`-`/`–`/`—`/`~`）、化学式上下标与半角括号 `(NH₄)₂SO₄`、序列极性 `5′-`/`3′-`/`N-`/`C-`、单位规范（`10 µL` 禁 `uL`、`37°C`/`95%` 无空格、整章时间单位统一、离心力 *g* 斜体 vs 质量 `5 g` 正体）、设备厂商与多子图中文逗号（`图 1.1 A，B`）、严禁使用显微相机原带标尺、统计检验与 *P* 值、图注末置 `修改自 Author et al., Year。`、`1.2` 研究目的定位、盲审合作署名泛化、以及 Nature 文献 7 要素（放行 `eLife`/`iScience`，“已正式见刊去 DOI、仅在线未见刊留 DOI”）。
   - **后续路线**：本版专注于版式结构与生科硬性规范；下一步将继续补强生科写作中常见的“似是而非、长而绕口”表达的顺句优化。
2. **Word 顶部自带 19 套可人工复用的「样式」模块（支持单 JSON 跨校切换）**
   - 无论无模板从头起底稿（`< 2` 秒直出 15 页含作者简历标准稿）还是按模板修版，全部章节标题、正文、图注、三线表、文献、F9 目录与简历都绑定为 Word 顶部独立样式（底层绕过 `PAGE`/`NUMBER`/`STYLE`/`DIRECT`/`MEDIA`/`IDENTITY`/`TOC` 7 类 14 项 XML 暗坑，修改一份 JSON 即可切换高校字号、奇偶页眉和引文标准），交到作者手里后还能像正常 Word 文档一样自己点选样式续写、加章节、插图表，不会像一次性排版脚本那样“一改就全散”。
3. **长论文含横向宽表的按语义角色跨节对齐**
   - 在三节模板对应五节含八列横向宽表旧稿的同机对照中，本插件 `repair --template` 实现 `20/20` 四侧边距全对（多出的第 4、5 节 `8/8`）、正文起始页回 `1`、镜像页边距开启且横向宽表完好；对比按序号提取规则的工具为 `14/20`（多出节 `2/8`）。

## 哪些成熟实践进入了本版

| 已有项目解决的问题 | 本插件采用的做法 | 可以怎样验证 |
| --- | --- | --- |
| Pandoc 允许调用方决定警告是否让命令返回失败。[命令行文档](https://pandoc.org/MANUAL.html#option--fail-if-warnings) | `audit --fail-on-warn` 让同一份审计报告既可用于写作中途自查，也可用于交稿前放行。 | 只改学校标题样式时，默认返回 0，严格模式返回 1。两个报告里的发现相同。 |
| 清华论文格式检查项目给自动规则和需要第二阶段人工复核的事项分别留证，并屏蔽私人论文产物。[项目流程](https://github.com/jacky-rui/tsinghua-thesis-format-checker-skill/blob/main/README.zh-CN.md) | JSON 与 Markdown 报告单列自动结果计数和待办项。盲审模式追加图片、签名及嵌入附件身份检查。Git 默认忽略工作稿、PDF 和常用的本地报告。 | 对同一份合成文档分别按终稿和盲审审计，可见三项和四项待办。待办独立于自动 PASS。 |

这两项功能按本项目的 Word 工作流自行实现。校级数值、文档内容和第三方源代码均未从上述项目导入。[中文使用说明](../README.md) · [基准对比](comparison.md) · [无模板对比](no-template-comparison.md)
