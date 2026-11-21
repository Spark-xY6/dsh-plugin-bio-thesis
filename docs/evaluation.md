# 从 DSH 安装到 Word 成稿

**0.2.0 版在真实 DeepSeek Harness 中完成了安装、技能加载、按模板生成盲审稿和交稿前检查。** 三个刻意破坏的样例分别检出了页边距偏差、姓名泄露和标题样式改变。全部材料均为合成样例。以下结果回答插件能否在这些条件下工作；学校的最终验收仍取决于本校模板、书面要求和成稿版面。

## 真实宿主如何验证

在 Windows、Node.js 24.13.0、pnpm 11.25.0、Python 3.13.9 和官方 DSH 0.2.2-rc.5 下，测试把发行包安装进一个全新的 profile（独立插件配置）。启动宿主后，直接从 DSH 的技能目录查询并加载 `bio-thesis-formatter`。测试使用宿主返回的资源路径执行生成和审计，全程无需在线模型密钥。自然语言对话中模型能否选对工具、遵守指令，仍待单独评测。

合成院系模板含一个拆在两个 Word 文本片段中的标题标记，同一标题在模板中出现两次。生成后两处标题均已填入，作者被匿名标志替代，输入模板的文件哈希保持一致。逐个对比 DOCX 内部文件，只有存放正文的 `word/document.xml` 发生变化。

| 测试输入 | 期望 | 观察结果 |
| --- | --- | --- |
| 按该模板生成的盲审稿 | 页边距匹配，已知姓名检查通过 | 11 项 PASS、20 项 INFO，0 项 WARN、0 项 FAIL，退出码 0 |
| 只把这份稿件的上边距由 2281 改成 900 个 Word 单位 | 报出页面设置偏差 | 10 项 PASS、20 项 INFO，仅 `Layout-1` 报 FAIL，退出码 1 |

Word 边距的一个单位等于 1/20 磅。第一份模板规定的上边距约为 4.02 厘米，错误稿缩至约 1.59 厘米。INFO 表示该项提供信息、条件缺失或未启用；它不会计入自动通过数。

## 从发行包验证错误能否被拦住

另一组测试从空目录离线安装本地生成的 npm 压缩包。它使用最小模拟对象检查注册入口，然后执行安装包内的命令行。Node 测试报告共 6 项通过，其中包括 1 个总测试和 5 个子测试。

| 场景 | 对结果的独立核对 |
| --- | --- |
| 通用起稿、预设和检查清单 | 安装后能读取随包资源，初始化 DOCX 并导出检查清单。 |
| 从学校 DOTX 起稿 | 验证分段标题替换、作者匿名和模板格式；盲审报告另列 4 项人工复核。 |
| 单改标题样式定义 | 将字号从 23 磅改为 15 磅，仅 `Layout-TemplateStyles` 报 WARN。默认退出 0，加 `--fail-on-warn` 退出 1。两次报告的自动检查和汇总相同。 |
| 同时改坏两处 | 将上边距从 1800 改为 720 个 Word 单位，并把已知姓名重新写入正文。退出 1，恰好检出 `Layout-1` 和 `Acad-10-Blind` 两项失败。 |
| 输出路径与输入重名 | 尝试把新稿或报告写回受保护的输入文件时，命令拒绝写入；比对文件内容确认原件完好。 |

独立的 Python 发布检查还运行了博士、硕士、开题、中期和轮转五种模式，并核对 Word 清单恰有 65 项。另有 4 组回归测试覆盖已知漏报、DOCX/DOTX 模板保留、输入输出保护等行为。这 65 项供作者对照自查；每次自动报告列出的条数随文档、预设和选项变化。当前终稿报告列出 3 项人工待办，盲审报告另加图片、签名和嵌入附件的身份检查。

## Word 打开后看到什么

这张图取自 Microsoft Word 16.0（build 16.0.17932）实际打开插件生成的 DOCX 并导出的单页 PDF。独立构造的院系 DOTX 使用 Letter 纸，Word 读回上边距 102 磅、左边距 78 磅、20 磅粗体标题、1 个页码域和 1 张嵌入图片。该样稿按模板审计为 12 项 PASS、20 项 INFO，0 项 WARN、0 项 FAIL。

![Word 导出样稿，全页显示原页眉、替换后的中文标题、匿名作者、橙色示例图片和页脚页码 1](word-sample.png)

页顶的英文文字来自原模板页眉。中文题目下面的作者栏已替换为“【匿名评审隐去】”。橙色方块是模板中的测试图片，底部的 1 来自 Word 页码域。这里核验的是一页受控样稿；长论文的分页、复杂跨页表格、更新后的目录页码和图片中的身份信息仍需打开 Word 或查看最终 PDF 核对。现有论文可对照模板审计，全文迁移到另一学校模板需另行排版。

## 如何复现

仓库自带 GitHub Actions 工作流。Windows 和 Linux 作业覆盖注册契约、Word 资源、回归和离线压缩包安装；独立作业会安装固定版本的官方 DSH 宿主并运行真实 profile。可在 Actions 中选择 **Release checks → Run workflow** 检查当前提交，或从仓库根目录运行以下 PowerShell 命令。[工作流源码](../.github/workflows/release-checks.yml)

```powershell
# 准备 Node.js 24.9+、pnpm 11.25.0、Python 3.10+，在独立临时目录安装 DSH。
$scratch = New-Item -ItemType Directory -Path (Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString()))
$env:DSH_RUNTIME = Join-Path $scratch.FullName 'host'
$env:PYTHON = (Get-Command python -CommandType Application | Select-Object -First 1).Source
npm install --prefix $env:DSH_RUNTIME --no-audit --no-fund @deepseek-ai/dsh@0.2.2-rc.5
$packed = npm pack --silent --ignore-scripts --pack-destination $scratch.FullName
$env:DSH_PLUGIN_SPEC = Join-Path $scratch.FullName $packed
node tests/dsh-host.mjs
```

成功时最后一行以 `DSH_E2E=PASS` 开头。如需测试公开仓库的特定版本，将 `DSH_PLUGIN_SPEC` 设为 `github:Spark-xY6/dsh-plugin-bio-thesis#<完整的40位提交SHA>`，再运行同一命令。以下命令分别用于复现模拟注册测试、Python 发布和回归检查、离线发行包测试。

```powershell
node --test tests/dsh_plugin.test.mjs
python -B -X utf8 tests/verify_release.py
python -B -X utf8 tests/test_regressions.py -v
node --test tests/packed-e2e.mjs
```
