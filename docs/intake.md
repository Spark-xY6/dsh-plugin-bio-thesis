# 三份学校模板的盲审起稿对照

**按三所不同学校（Letter / A4 / 横向页面）的模板同时起稿时，本插件让代理完全无需现场手写底层 Word OOXML 转换代码（新增转换代码从 33–37 行降至 0 行），直接生成并通过 12 / 12 项 Word 结构与桌面打开检查。** [审计实测](comparison.md) · [基准任务与评分规则](../benchmarks/template-intake/README.md)

六次配对会话共交付 18 份 DOCX。每份成品均通过 4 / 4 项独立结构检查（每组三份合计 12 / 12），且均可由桌面 Microsoft Word 原生打开，识别出两节、两页、一个活动目录及一张内嵌插图。[逐份评分明细](../benchmarks/template-intake/model-outcomes.json) · [桌面 Word 读数](../benchmarks/template-intake/word-open.json)

## 三份模板分别要求什么

三份合成模板分别使用 Letter、A4 和横向页面（一份为 DOCX，两份为 DOTX），各自带有独立的正文样式、页眉、嵌入图片及不同位置的标题与匿名字段。任务要求完整保留原有版式，填入中英文题目与专业，并将申请人与学号替换为可见匿名标记。[任务说明](../benchmarks/template-intake/TASK.md)

## 每一对的真实交付

在 Windows Codex CLI `0.169.0`（`gpt-6.1-sol`，低推理设置）上执行 3 组配对试验（裸跑 vs 启用技能）。主评估指标为代理临时编写的 DOCX/ZIP/XML 底层操作代码行数。

| 配对 | 成品得分 | 新写底层转换代码（行） | 用时（秒） | 工具调用 | 输出 token |
| --- | --- | --- | --- | --- | --- |
| 1，裸跑先 | 12 → 12 | 37 → 0 | 62.742 → 66.584 | 4 → 5 | 1587 → 1524 |
| 2，技能先 | 12 → 12 | >=33 → 0 | 77.564 → 75.469 | 7 → 7 | 1795 → 1658 |
| 3，裸跑先 | 12 → 12 | 36 → 0 | 65.318 → 65.634 | 5 → 5 | 1622 → 1533 |

三组启用技能的会话均直接调用内置命令完成三份异构模板填充，底层转换代码行数中位数从 `36` 降至 `0`，输出 token 中位数下降 `5.5%`（`1622 → 1533`），同时完整守住 `12/12` 成品质量。[逐行归因](../benchmarks/template-intake/source-review.json) · [解盲后汇总](../benchmarks/template-intake/implementation-outcomes.json)
