# 论文格式一把过：Word 模板起稿、无模板按要求微调 & 65 项盲审自查

[English](README.md) · [无模板配对实测与 Word 实拍](docs/no-template-comparison.md) · [三校模板起稿对照](docs/intake.md) · [同稿审计与开源工具实测](docs/comparison.md) · [Word 样例](docs/evaluation.md)

**一句话看懂它干嘛**：写毕业论文/开题/中期/轮转报告时，**不管你手头有没有学校现成的 Word 模板**，这个 DeepSeek Harness (DSH) 插件都能帮你搞定两件最头疼的事：
1. **起草不返工**：有学院 `.docx/.dotx` 模板就原样保真起稿；**没有现成模板、只有学院发的文字版格式要求**，就直接调用内置**脱敏通用母版 + JSON 参数微调**，一条命令出出版级底稿（三节页码、页眉横线、F9 可加粗目录、规范三线表与盲审隐去全配齐）；
2. **交稿不踩雷**：对照模板或文字规范自动抓出隐藏错处——某一节页边距改偏了、`core.xml` 文档属性里漏了作者姓名、物种/基因没斜体、`20 uL` 没写 `µL`、表格多了竖线或多了未引用文献。

---

## 先看 Input -> Output 真实 Word 实拍（无模板场景）

只给 AI 一份文字版学院格式要求（A4、自定义页边距 `2.8/2.4/2.8/2.5 cm`、盲审隐去、三节页码、Nature 引文与三线表），同一模型 `gpt-6.1-sol` 在 Microsoft Word 16 实测渲染如下：

![无学校模板场景下，无插件裸跑与内置脱敏母版微调的 Microsoft Word 渲染对照](docs/images/no-template-showcase.png)

| 对比维度 | 无插件裸跑（直接让 Agent 写代码） | 开源轮子（Pandoc / format-thesis-docx 等） | 使用 `dsh-plugin-bio-thesis`（v0.3.0） |
|---|---|---|---|
| **无模板按文字要求起稿** | 需现写 **183–208 行** `python-docx`；XML 虽过但 Word 渲染呈“散字封面 + 默认蓝字标题 + 无缩进 + 一级目录未加粗 + 空白页” | Pandoc `--reference-doc` 不保留模板封面和声明页（三校测试仅过 `5/12`） | **0 行手写代码**，1 条命令微调内置脱敏母版：Word 实开 **15/15 出版级版式**（下划线封面、F9 真加粗目录、纯黑标题、页眉线、2 字符缩进） |
| **有学校 `.dotx`/`.docx` 模板起稿** | 三校模板每次现场手写 **33–37 行**底层 XML 转换代码 | `docxtpl` 可填表但无交稿前结构审计 | **0 行转换代码**，18 份成品全过并可直接在 Word 打开 |
| **交稿前缺陷全面审计** | 需现场写脚本逐项核查 | `format-thesis-docx` 擅长自动重排现稿，但同套 7 条校规测试仅自动覆盖 **4/14** 缺陷（`DocFormatTool` **1/14**） | 原生 CLI 自动检出 **14/14** 缺陷（含局部字号覆盖）及无模板 8 类盲审/生科规范缺陷（**8/8**） |
| **盲审防泄露与人工待办清单** | 容易漏查 `docProps/core.xml` 与分段拆分姓名 | 多数工具无跨 XML 身份词与媒体字节防篡改比对 | 同时扫正文、页眉页脚、关系与属性，单独列出签名/图片/跨页人工复核待办及 65 项自查表 |

> **诚实边界说明**：在无模板自主会话配对测试中，若让 Agent 在已生成的内置母版上再行手写底层 XML 补丁，出现过一次手改 XML 导致 Word 报错“文件可能已经损坏”；**推荐始终直接通过 JSON `--profile-config` 与 CLI 参数微调母版**，不要画蛇添足手改底层 XML。[查看完整无模板实测与教训](docs/no-template-comparison.md)

---

## 1 分钟安装与直接套用话术

运行环境：Python 3.10+（**全部处理脚本仅用 Python 标准库，零 `pip install` 第三方依赖**）与 DSH `0.2.2-rc.5+`。

```powershell
dsh plugin --profile web add "github:Spark-xY6/dsh-plugin-bio-thesis#v0.3.0"
```

安装重启该 profile 后，按你手头情况直接复制下面对应的一句话发给 DSH：

- **情况 A（没有模板，只有一段学院文字格式要求）：**
  > 请用 bio-thesis-formatter，以内置脱敏母版按我的文字要求（把页边距、引文样式写进 JSON profile 配置，不要手改生成后的 XML）生成博士/硕士/开题报告底稿「论文工作稿.docx」，并把「待检稿.docx」按同一套要求查出所有格式和盲审泄露问题，输出「格式检查.md」。

- **情况 B（手里已经有学校发的 `.dotx` / `.docx` 模板）：**
  > 使用 bio-thesis-formatter，以「学院模板.dotx」生成「论文工作稿.docx」，题目填为「示例研究」，盲审模式隐去作者和导师。另把「修订稿.docx」与这份模板比较，写出「格式检查.md」，列出我还需要在 Word 里人工核对的地方。

### 独立命令行常用三句（脱离 DSH 也可离线运行）

```sh
# 1) 无模板：用自定义 JSON 配置微调内置脱敏母版起稿（或加 --template school.dotx 用学校模板）
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py init --profile-config presets/custom-university.example.json --mode phd-thesis --review-stage blind --output draft.docx --title-zh "示例研究" --author "示例作者"
# 2) 交稿前全面体检（无模板用 --profile-config；有模板加 --template school.dotx；可加 --fail-on-warn）
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py audit --docx revised.docx --profile-config presets/custom-university.example.json --review-stage blind --sensitive-names "示例作者,示例导师" --md-report review.md
# 3) 导出可编辑 65 项自查表
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py export-checklist --output checklist.docx
```

---

## 进阶实测、边界与开源许可

- 📊 **[无模板文字要求配对实测（Word 渲染对比）](docs/no-template-comparison.md)** · **[三校模板起稿对照](docs/intake.md)** · **[14 缺陷与开源工具横向实测](docs/comparison.md)**
- 🔍 **[已有同类工具调研与借鉴来源](docs/alternatives.md)** · **[真实 DSH 宿主安装与 Word 样例验收](docs/evaluation.md)** · **[持续集成](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/actions)**
- **注意**：生成工作稿中的表格/引文均为排版占位符，不含真实实验数据；交付前仍需在 Microsoft Word 中选中目录按 `F9` 更新页码，并目视核对跨页表格与图片中的签名/身份信息。源代码基于 [MIT 许可证](LICENSE)。
