# 用学校的 Word 模板写论文、查格式

[English](README.md) · [已有工具调研](docs/alternatives.md) · [实测效果](docs/evaluation.md) · [版本下载](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/releases)

学院发来一份 Word 模板，封面、声明、页眉和页码已经排好。这个 DeepSeek Harness（DSH）插件能从模板生成可编辑底稿。经过几轮修改，还能将现稿与原模板对照，指出哪些设置变了。硕士、博士及过程报告均可使用。生命科学术语和引文筛查按需启用。

## 为什么做

模板解决了“从哪里开始写”，交稿前还需要回答“哪一处改错了”。一篇外观看起来正常的论文，可能有一节的页边距被改过，或者文件属性里仍留着作者姓名。我们调研了学校官方模板、Pandoc、DocFormatTool、网页排版工具和开源论文检查技能，将“警告也可阻止放行”和“单列人工复核任务”落实到本版插件。[已有方案、适用场景与来源](docs/alternatives.md)

## 在 DSH 中安装和调用

使用备有 pnpm 和 Python 3.10+ 的 DSH 环境。宿主实测使用 DSH 0.2.2-rc.5 和 Node 24。安装到你正在使用的 profile，然后重启该 profile。

```sh
dsh plugin --profile web add "github:Spark-xY6/dsh-plugin-bio-thesis#v0.2.0"
```

也可以从 [GitHub Release](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/releases/tag/v0.2.0) 下载预构建包，在下载目录运行 `dsh plugin --profile web add ./dsh-plugin-bio-thesis-0.2.0.tgz`。在 DSH 工作目录中放入文件后，可以这样说：

> 使用 bio-thesis-formatter，以“学院模板.dotx”生成“论文工作稿.docx”，题目填为“示例研究”。保留学校封面和页眉。另把“修订稿.docx”与这份模板比较，写出“格式检查.md”，列出我还需要在 Word 里看的地方。

## 直接运行

从仓库根目录执行。独立命令行工具只用 Python 标准库，可在本机离线处理文档。使用 DSH 对话时，文件访问和模型服务取决于你的 DSH 配置。

```sh
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py init --template school.dotx --output draft.docx --title-zh "示例研究"
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py audit --docx revised.docx --template school.dotx --md-report review.md --json-report review.json
python -B skills/bio-thesis-formatter/scripts/bio_thesis.py export-checklist --output checklist.docx
```

- **从学校模板起稿。** 使用无宏、无数字签名的 DOCX 或 DOTX。无占位符也可生成保留原内容的工作稿。如需自动填题目、作者和学院，在普通 Word 文字里放入 `{{title_zh}}`、`{{author}}` 或 `{{school}}`。还支持 `title_en`、`university`、`student_id`、`advisor`、`major`、`date_zh`。校徽直接放入模板。域和嵌套内容控件需另行处理。
- **检查已有稿件。** `audit` 只读对照各节纸张、边距、页码配置、目录域和样式定义。需要检查插图时，加 `--base-docx 上一版.docx`，可检出内嵌媒体的增删和替换。整篇旧稿迁移到另一份模板仍需单独排版。
- **准备盲审。** 生成时传 `--review-stage blind` 遮盖身份占位符。审查现稿时，同时传 `--review-stage blind --sensitive-names "示例作者,示例导师"`，查找 XML、关系和文档属性中的残留。签名图片和嵌入附件需要目视检查。完整报告可能含本机路径、匹配词和原文片段，请保存在自己的工作目录。
- **按学院调整。** 模板决定版式。可用 `--citation-style gb7714` 或 `nature` 启用相应数字引文筛查。需检查三线表时，在 JSON 配置写入 `{"audit":{"template_checks":["Tbl-3Line"]}}` 后传 `--profile-config rules.json`。没有学校模板时，省略 `--template` 可生成内置示例。[配置示例](presets/custom-university.example.json)
- **判读报告。** `FAIL` 表示已启用的规则检出错误，`WARN` 表示需判断，`INFO` 表示提示或未启用。默认有 FAIL 返回 1，只有警告、提示或通过项返回 0，输入或运行错误返回 2。加 `--fail-on-warn` 可让警告也阻止放行。报告单列待人工复核任务。65 项自查表同时包含自动和人工项目，最终目录、分页、跨页表格仍要在 Word 或最终 PDF 中确认。

## 实测效果

一份横向 Letter 学校模板生成盲审稿后，自定义页眉和标题样式保留。将上边距由 1800 改为 720 twip（1.25 英寸改为 0.5 英寸），再把示例作者姓名写回，审计准确指出了两处变化。另有真实 DSH 安装测试和 Microsoft Word 打开后的页面检查。[评测记录、复现步骤与限制](docs/evaluation.md) · [持续集成](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/actions)

[MIT 许可证](LICENSE)
