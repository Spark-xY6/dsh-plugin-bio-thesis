# 论文格式一把过：Word 模板起稿、无模板按校规微调 & 65 项盲审自查

[English](README.md) · [无模板配对实测与 Word 实拍](docs/no-template-comparison.md) · [三校模板起稿对照](docs/intake.md) · [同稿审计与开源工具实测](docs/comparison.md) · [Word 样例](docs/evaluation.md)

写毕业论文、开题、中期或轮转报告时，**不管手头有没有学校现成的 Word 模板**，这个零第三方依赖的 DeepSeek Harness (DSH) 插件都能帮你解决两件最容易返工的事：

1. **排版规矩直接锁进 Word 顶部「样式」栏，拿到文件可放心续写**：有学院 `.docx/.dotx` 模板就原样保真起稿；**只有学院文字要求、没有现成模板时**，一条命令按 JSON 参数自适应重算内置脱敏母版，把标题、正文、图注、表题、三线表与参考文献全部固化进 **Word 上方 19 套专属 `qFormat` 样式**，并在正文中带好可填改的**三线表与图注样板**——人手接过去继续写字、添表、加图也不会越改越乱。
2. **交稿前自动揪出隐蔽硬伤**：逐节核对页边距与奇偶页页眉、扫描 `core.xml` 和拆分 XML 里的盲审身份残留、检查物种/基因斜体、`µL` 单位写法、三线表竖线与未引用文献。

---

## Word「样式」自适应与成稿实拍

只提供一段纯文字格式要求（不含任何 `.docx` 模板），微软 Word 16 实拍如下：插件不靠临时写千行脚本拼表面外观，而是直接重算 `word/styles.xml` 中的 19 个顶部样式并附带现成规范模块；同一条命令切换学校 JSON 参数即可从“某高校（宋体 20 pt / 黑体 / Nature）”自适应切换为“乙高校（仿宋 22 pt / 微软雅黑 / GB/T 7714）”，均一次通过 `FAIL: 0`：

![Word 顶部 19 个自适应样式与某高校至乙高校样式切换实拍](docs/images/no-template-showcase.png)

双盲封面脱敏、三节隔离可刷新目录、以及 7 校 14 项典型暗坑自动体检对比开源同类工具实测如下：

![双盲封面、目录页实拍与 7 校 14 项暗坑自动体检对比](docs/images/result-and-audit-showcase.png)

| 对比维度 | 无插件裸跑（让 Agent 现写代码） | 开源工具对照（Pandoc / format-thesis-docx 等） | 使用 `dsh-plugin-bio-thesis`（v0.3.0） |
|---|---|---|---|
| **Word 顶部样式与可继续编辑性** | 需现写 **183–1,212 行**脚本；Word 样式栏仍堆 29 个默认杂项、缺图注/文献样式、人手改稿极易串版 | Pandoc `--reference-doc` 会丢失原模板封面与声明页（三校仅 `5/12`） | **0 行定制代码**，`<2s` 直出 15 页工作稿并将规则全锁入 **19 个顶部样式 + 现成三线表/图注模块** |
| **多校格式要求自适应切换** | 换一所学校就得重写一遍底层样式代码 | `docxtpl` 偏重填表，缺少随文字校规重算样式的审计闭环 | 同一命令换 `--profile-config` JSON 即重算全部样式（某高校 `宋体 20pt` ⇄ 乙高校 `仿宋 22pt` 均 `FAIL: 0`） |
| **有学校 `.dotx`/`.docx` 模板起稿** | 三校模板每轮需现写 **33–37 行**底层 XML 转换 | 多数模板工具缺少改后结构体检 | **0 行转换代码**，三校 `18/18` 成品在 Word 16 正常打开 |
| **交稿前暗坑体检与防漏** | 易遗漏 `core.xml` 与跨 Run 拆分姓名 | 同套 7 校 14 项规则测试中，`format-thesis-docx` 仅检出 **4/14**，`DocFormatTool` **1/14** | 原生审计自动检出 **14/14 (100%)** 缺陷与 8 类无模板盲审/生科规范错处，并附 65 项自查清单 |

> **诚实工程边界提示**：在无模板自主会话配对测试中，曾发现若让 Agent 在已生成的内置母版上再手工编写底层 XML 补丁，出现过一次手改 XML 导致 Word 报错“文件可能已经损坏”；**推荐始终直接通过 `--profile-config` JSON 与 CLI 参数驱动 `bio_thesis.py init`**，不要事后手改生成好的底层 XML。[查看完整无模板实测与教训](docs/no-template-comparison.md)

---

## 1 分钟安装与直接套用话术

运行环境：Python 3.10+（**全部脚本仅依赖 Python 标准库，无需 `pip install` 第三方包**）与 DSH `0.2.2-rc.5+`。

```powershell
dsh plugin --profile web add "github:Spark-xY6/dsh-plugin-bio-thesis#v0.3.0"
```

安装重启该 profile 后，按你手头情况直接复制下面对应的一句话发给 DSH：

- **情况 A（没有模板，只有一段学院文字格式要求）：**
  > 请用 bio-thesis-formatter，以内置脱敏母版按我的文字要求（把页边距、字号和引文样式写进 JSON profile 配置，不要手改生成后的 XML）生成博士/硕士/开题报告底稿「论文工作稿.docx」，并把「待检稿.docx」按同一套要求查出所有格式和盲审泄露问题，输出「格式检查.md」。

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
- **注意**：生成工作稿中的表格与引文均为排版占位符，不含真实实验数据；交付前仍需在 Microsoft Word 中选中目录按 `F9` 更新页码，并目视核对跨页表格与图片中的签名/身份信息。源代码基于 [MIT 许可证](LICENSE)。
