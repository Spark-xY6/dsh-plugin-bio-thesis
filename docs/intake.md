# 三份学校模板的盲审起稿对照

**按三所学校的模板起稿时，插件省去了模型现场编写底层 Word 转换代码的工作。** 三组配对中，裸跑模型分别编写 37 行、至少 33 行和 36 行；加载技能后，三次都直接调用已发布的生成命令，新增转换代码均为 0 行。两侧的成品质量相同。这个任务在首次审计对照打平后设计，判定标准在本轮模型运行前公开。[审计实测](comparison.md) · [事前登记](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/dd5e5ff9e6c3d20d6f9472ea9547bf7f3fd1d535/benchmarks/template-intake/PROTOCOL.md)

六次会话一共交付 18 份 DOCX。每份成品均通过 4/4 项检查，每次三份合计 12/12。桌面 Microsoft Word 也能打开这 18 份文件，每份均识别出两节、两页、一个目录及一张内嵌图片。[逐份评分、时间与 token](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/model-outcomes.json) · [桌面 Word 读数](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/word-open.json)

## 三份模板分别要求什么

三份合成模板从同一个基础包派生，分别使用 Letter、A4 和横向页面。其中一份直接以 DOCX 起稿，两份从 DOTX 转成 DOCX。每校有自己的正文说明、段落样式、页眉、嵌入图片，以及出现位置不同的填写字段。模型需保留三种版式，填入各自的题目和专业，将申请人和注册号替换为可见匿名标记。没有提供真实论文、姓名或学号。[实际任务](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/TASK.md)

每份成品独立检查四项：Word 包可编辑，指定字段和可见标签填写正确，所有提供的身份值均已去除，非填写部分仍保留原模板内容和结构。三份合计 12 分。裁判另用 18 个正例、反例和等价 XML 改写校准评分，只改变压缩或 XML 序列化方式可以取得同分。[评分器与控制样例](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/README.md)

## 怎样判定优势

这项次级实验在首次审计实验打平后设计。新任务的评分规则和判定门槛已在本次模型运行前提交到公开仓库。模型未收到裁判、答案或其他会话记录。三组使用同一模型、同一提示词和 600 秒上限，分别按裸跑先、技能先、裸跑先执行。三个配对可并行，每对内依次运行。宿主为 Windows Codex CLI 0.169.0，模型为低推理设置的 `gpt-6.1-sol`。技能在真实 DSH 的安装和运行另有验收记录。[事前公开的方案](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/dd5e5ff9e6c3d20d6f9472ea9547bf7f3fd1d535/benchmarks/template-intake/PROTOCOL.md) · [冻结输入与代码哈希](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/freeze.json)

主要指标是模型为完成本次委托临时编写的底层文档转换代码。逐行统计实际执行、去重后实现或直接调用 DOCX/ZIP/XML 改写的非空代码。现成文档命令的参数、读取任务、遍历文件和只读成品检查不计入。全部配对须交付有效 DOCX，质量中位数须持平或更高，新写行数中位数须下降至少 50%，且至少两对分别减少 8 行。行数受编码风格影响。它度量为这次任务新承担的实现量，无法推算开发插件本身的工作量、维护成本或用户满意度。

另行判定成品质量、效率和模型输出量。质量优势要求至少两对各多通过两项。时间或工具调用优势要求相应中位数下降 30%；输出量优势要求模型报告的输出 token 中位数下降 25%。这些效率指标都以质量中位数不降低为前提。一个 token 是模型服务计量输出的一段文字。输出 token 并非费用，也不包括输入。

## 每一对的真实交付

表中箭头从裸跑指向启用技能。每次质量满分为 12，覆盖三份底稿。

| 配对 | 成品得分 | 新写转换代码（行） | 用时（秒） | 工具调用 | 输出 token |
|---|---|---|---|---|---|
| 1，裸跑先 | 12 → 12 | 37 → 0 | 62.742 → 66.584 | 4 → 5 | 1587 → 1524 |
| 2，技能先 | 12 → 12 | ≥33 → 0 | 77.564 → 75.469 | 7 → 7 | 1795 → 1658 |
| 3，裸跑先 | 12 → 12 | 36 → 0 | 65.318 → 65.634 | 5 → 5 | 1622 → 1533 |

三份裸跑实现读取 Word 压缩包，修改内部 XML 并写回成品。第二对的首次校验失败，模型修正后交付；执行记录保存了最终脚本和修改事件，首版的逐行差异无法完整还原，因此采用可复核的下界。独立的自动评审代理只收到随机编号的命令和代码，先固定行号与计数，再解开分组；从实际命令本身仍可能推断条件。六次会话都将同一套实现复用于三份模板。[逐行归因](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/source-review.json) · [解盲后计数](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/implementation-outcomes.json) · [去标识的代码记录](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/tree/v0.3.0/benchmarks/template-intake/review-sources)

**预设的实现工作量指标达标。** 新写转换行数的中位数从至少 36 降至 0，三对平均质量和每份成品质量均持平。时间中位数从 65.318 秒变为 66.584 秒，工具调用中位数均为 5。输出 token 中位数从 1622 降至 1533，降幅 5.5%。准确率、耗时、调用数和输出 token 均未达到各自的优势门槛。这三份从共同基础包派生的合成模板只检验此次任务，无法推算真实学校的送审通过率或其他模型的表现。

下载证据包可逐份打开实际成品，也可在本地重新评分。压缩包收录两项实验的全部有效配对及合成输入，含文件级 SHA-256 清单。[证据 ZIP](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/releases/download/v0.3.0/thesis-benchmark-evidence-v1.zip) · [复现步骤](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/v0.3.0/benchmarks/template-intake/README.md)
