# 学位论文排版与交稿核查：常见工具对比与适用场景

学校发下来的 Word 模板通常已经放好了封面、声明页和校徽。写第一稿时，学生希望把这些内容带进新文档。几轮修改之后，问题变成了“正文的页码还是从 1 开始吗”“送盲审的文件属性里还有姓名吗”“换稿时是否误替换了插图”。从模板起稿与交稿前检查，是这个 DSH 插件要接起来的两件事。

## 已有工具能做什么

下表按工作任务选取五个实例。功能描述来自学校、项目维护者或产品提供方于 2026 年 11 月 21 日可查的公开说明。这份初步调研依据项目文档。后续对其中的检查器、Pandoc 和相近的开源工具运行了同一组合成样稿，并记录各自适用的项目。[同稿对照实测](comparison.md)

| 手头的材料和目标 | 已有方案 | 适用情形和依据 |
| --- | --- | --- |
| 从本校认可的版式开始写 | 上海交通大学研究生院提供博士、硕士论文 Word 和 LaTeX 模板 | 该校学生可以直接使用研究生院发布的起稿附件。研究生手册另说明，学院可按学科特点制定模板。选择时要核对学位层次和学院要求。[研究生院下载页](https://www.gs.sjtu.edu.cn/post/detail/Z3M2MjU) · [研究生工作手册](https://www.gs.sjtu.edu.cn/storage/gs/web/cn/article/2025/11/a11b68c3cf6d5138bce86f635d6e15ea.pdf) |
| 正文已经写成 Markdown，想导出 Word | Pandoc 的 `--reference-doc` | 新文档可采用参考 DOCX 的样式、纸张、页边距和页眉页脚。参考文档正文不会被带入，学校封面和声明页若在正文中，需另行处理。[Pandoc 用户手册](https://pandoc.org/MANUAL.html#option--reference-doc) |
| 已有一篇论文，希望根据学校模板重排全文 | DocFormatTool | 项目说明提供从学校模板提取规则、对 DOCX 改写排版、把 Markdown 或纯文本转成 Word，以及排版后质检的命令。[DocFormatTool 项目说明](https://github.com/NeoHotaru/DocFormatTool) |
| 想在浏览器里检查并套用现成校级样式 | FreeFormat | 产品页面提供上传 DOCX、查看评分和问题列表、选择模板、下载格式化文档的操作流程。适合允许上传文稿、希望通过网页操作的用户。[FreeFormat 工具页](https://freeformat.app/zh/tool) |
| 有清华研究生论文的 DOCX 和最终 PDF，要按特定校规审查 | Tsinghua Thesis Format Checker Skill | 该项目结合 DOCX 结构和 PDF 版面，生成结构化报告，并为图表位置、跨页和摘要对应等项目安排人工复核。[检查流程与适用范围](https://github.com/jacky-rui/tsinghua-thesis-format-checker-skill/blob/main/README.zh-CN.md) |

## 哪些做法进入了本版

| 已有项目解决的问题 | 本插件采用的做法 | 可以怎样验证 |
| --- | --- | --- |
| Pandoc 允许调用方决定警告是否让命令返回失败。[命令行文档](https://pandoc.org/MANUAL.html#option--fail-if-warnings) | `audit --fail-on-warn` 让同一份审计报告既可用于写作中途自查，也可用于交稿前放行。 | 只改学校标题样式时，默认返回 0，严格模式返回 1。两个报告里的发现相同。 |
| 清华论文格式检查项目给自动规则和需要第二阶段人工复核的事项分别留证，并屏蔽私人论文产物。[项目流程](https://github.com/jacky-rui/tsinghua-thesis-format-checker-skill/blob/main/README.zh-CN.md) | JSON 与 Markdown 报告单列自动结果计数和待办项。盲审模式追加图片、签名及嵌入附件身份检查。Git 默认忽略工作稿、PDF 和常用的本地报告。 | 对同一份合成文档分别按终稿和盲审审计，可见三项和四项待办。待办独立于自动 PASS。 |

这两项功能按本项目的 Word 工作流自行实现。校级数值、文档内容和第三方源代码均未从上述项目导入。

## 什么时候用本插件

- **拿到了学院模板，准备开写。** 将无宏、无数字签名的 `.docx` 或 `.dotx` 交给 `init --template`，得到一份新的可编辑 DOCX。已有封面、声明、样式和插图随模板保留。如需批量填写中英文题目，可先在模板副本的普通文字中放入 `{{title_zh}}` 和 `{{title_en}}`。
- **导师改过稿，想知道哪些设置偏离模板。** `audit --template` 可对照各节纸张、边距、页码配置和样式定义，并导出 JSON 或 Markdown。给定原稿并启用 `--base-docx` 后，还能按文件路径和 SHA-256 检查内嵌媒体是否变化。正文已写完需要对照学院模板对齐边距、页眉页脚距、奇偶页或分节页码（含横向宽表保护）时，可用 `repair --template` 先预览每节厘米级差异，再另存无损修正版；如需把纯文本或散乱段落从头映射成全新样式结构，也可结合上表中的 DocFormatTool。
- **准备送盲审。** 可以提供姓名、学号等关键词，检查它们在 Word XML 和文档属性中的残留。签名扫描件里的名字、复杂表格跨页、目录更新后的实际页码仍要打开 Word 或导出 PDF 核对。67 项清单同时包含自动检查和人工复核项。

DSH 用户可通过同一技能调用这些操作；也可以在本机直接运行仅依赖 Python 标准库的命令行工具。学校模板决定版面，学院书面要求决定哪些附加规则应当启用。具体命令见[中文使用说明](../README.md)。
