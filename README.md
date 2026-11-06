# DocxMathTypeBatch · MathType 批量转换器

[![中文图形界面](https://img.shields.io/badge/GUI-中文界面-2563eb)](#三步开始)
[![正式版本 v0.1.5](https://img.shields.io/badge/release-v0.1.5-2563eb)](https://github.com/yuxmic/DocxMathTypeBatch/releases/latest)
[![Windows x64](https://img.shields.io/badge/Windows-10%20%2F%2011%20x64-0078D4)](#运行环境与支持范围)
[![MIT License](https://img.shields.io/badge/license-MIT-16a34a)](LICENSE)

**中文图形界面（GUI），解压即用。将 Word 文档中的 LaTeX 和word原生公式批量转成可双击编辑的 MathType。**

整理论文、毕业设计或技术报告时，拖入 DOCX，点击“开始转换”即可。默认设置轻松上手，也能按需调整字号、行距和公式编号，或交给 AI Agent Skill 自动处理。

**[立即下载 Windows 中文界面版（ZIP）](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1.5/DocxMathTypeBatch-v0.1.5-win-x64.zip)** · [下载 AI Agent Skill](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1.5/docx-latex-to-mathtype-skill-v0.2.2.zip) · [使用说明](docs/使用说明.md)

运行前需安装 **64 位桌面 Word 和完整 MathType**，使用 Windows 10 / 11 x64。本工具免费开源，提供现成 EXE，无需编程或配置开发环境。

![DocxMathTypeBatch 中文主界面](docs/images/ui-main.png)

## 为什么用它

| 你关心的            | 程序提供的                                                     |
| --------------- | --------------------------------------------------------- |
| **小白也能上手**      | 中文 UI，拖入文档即可按默认设置转换，日志提供中文原因和处理建议                         |
| **开箱即用**        | windows系统双击本 EXE，运行只需word、mathtype环境                      |
| **批量省时**        | 多文档队列、添加文件夹，一次处理多份 DOCX。本机合成样本实测，100 个 LaTeX 公式常规转换约 45 秒 |
| **选项丰富**        | 转换范围、Word 原生公式开关、行内/行外字号、行距、章/节/起始序号和右编号格式均可设置            |
| **支持 AI Skill** | 提供内含转换器的完整 Skill，用自然语言要求本机 Agent 完成批量转换                   |
| **结果方便继续用**     | 生成可编辑 MathType 对象，另存新文档保留原件，支持基本的 MathType → LaTeX 反向导出   |

## 常见写法与排版，一次整理

兼容常见 LaTeX 写法、段落缩进和新旧编号，减少转换前后的手工整理。

| 常见情况 | 程序如何处理 |
| --- | --- |
| **新旧公式混排** | 正文中的 LaTeX 和 Word 原生公式一起转换，已有 MathType 公式保留内容，符合条件的行外公式按设置统一排版、补号或更新编号 |
| **行外公式带缩进或 Tab** | 处理首行缩进、字符缩进、悬挂缩进及前导 Tab，保留段落左右缩进，在可用宽度内居中 |
| **公式后有手打编号** | 清理独立 LaTeX 或 Word 原生公式后，通过空格或 Tab 手打的 `(27)`、`（27）`、`(3-2)`、`(3.2)` 等编号，开启右编号后生成 MathType 原生编号 |
| **已有 MathType 原生编号** | 开启右编号后，新公式补号，已有编号更新序号，相关编号引用同步更新。之后仍可用 MathType 格式化编号、全文更新和追加编号 |
| **先转换，以后再补编号** | 正文行外公式转成 MathType 后，仍可单独补号或调整起始序号。重复处理时复用已有原生编号，避免叠加编号 |
| **常见 LaTeX 命令写法不同** | 自动适配 `\bm`、`\mathup`、`\lvert`、`\mleft`、`\coloneqq` 等常见命令，减少逐式修改源码的操作 |
| **多行公式被 Word 换行拆开** | 支持 `aligned`、`align*`、`cases`、`array` 和矩阵等常见结构。定界符完整的结构化公式跨 Word 段落或软换行时，仍可整体识别，保留公式内部的分行、分列 |
| **复制来的字母带帽号、横线** | 自动适配 `Ŷ`、`Ȳ`、`Ŷ` 等已支持的字母顶饰，LaTeX 和 Word 原生公式都有对应处理 |
| **少量公式转换失败** | 常规模式下，局部公式失败可保留原式并继续处理其他公式。中文摘要提供原文页码、公式片段、原因和建议，方便定位处理 |
| **文档还含图片、表格和附件** | 常规模式检查表格结构、已有 MathType 数据、图片及嵌入附件，例如嵌入的 Excel，核对这些内容在转换和保存后得到保留 |

完整兼容写法见[公式书写与兼容性](docs/MathType兼容LaTeX书写规范.md)，编号及排版设置见[使用说明](docs/使用说明.md)。

## 三步开始

1. **下载并打开**。解压上方 ZIP，双击 `DocxMathTypeBatch.exe`，进入中文界面。也可[单独下载 EXE](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1.5/DocxMathTypeBatch.exe)。
2. **拖入 Word 文档**。先保存并关闭文档，再拖入一个或多个 `.docx`，也可点击“添加文件夹”。
3. **点击“开始转换”**。默认设置即可使用，完成后双击列表中的文档打开结果，新公式可在 Word 中双击编辑。

默认使用标准模式，转换所有正文，字号和行距跟随原文，不加右编号。三级滑块可选标准／快速／极速；极速专注正文行内、行外的 LaTeX 和 OMML，只保留指定字号及简单右编号，输出未经核验，支持范围见[使用说明](docs/使用说明.md)。输出重名时自动添加序号。

首次运行可能出现 Windows 未签名程序提示。转换期间请避免同时操作 MathType，以免中断转换服务。详细步骤见[使用说明](docs/使用说明.md)。

## 转换效果

LaTeX 文本和 Word 原生公式转换后，可继续在 MathType 中编辑。需要公式编号时，可生成 MathType 原生右编号。

| 转换前                                                             | 转换后                                                             |
| --------------------------------------------------------------- | --------------------------------------------------------------- |
| ![转换前，Word 中的 LaTeX 文本和原生公式](docs/images/conversion-before.png) | ![转换后，可编辑的 MathType 公式和行外右编号](docs/images/conversion-after.png) |

上图来自同一份合成文档的实际 Word 渲染，包含行内式、分式、求和、积分、分段函数、矩阵和多行公式。

## 实测速度

**极速模式：含图的真实复杂 1015 式文档，12 pt、开启进度和右编号，本机三轮实测 49.325～56.147 秒。** 全部 1015 式成功，生成 649 个连续编号；独立公式后的手动旧号和多余空格、Tab 也可在转换成功后清理。

以下为标准／快速模式的早期合成样本成绩：

| 合成样本                     | 常规模式   | 快速模式   |
| ------------------------ | ------ | ------ |
| 20 个 LaTeX 分式 / 80 段     | 约 12 秒 | —      |
| 20 个 Word 原生分式 / 80 段    | 约 13 秒 | —      |
| 100 个简单 LaTeX 公式 / 400 段 | 约 45 秒 | 约 22 秒 |

包含 Word 启动、转换、保存和退出清理。常规模式保留逐式及保存后核验，快速模式省略这些核验。实际耗时随公式复杂度、文档和机器环境变化。完整环境、精确计时和核验记录见[实测报告](docs/实测报告.md)。

极速直接将支持的 OMML 结构编译成 LaTeX 后导入 MathType；粗斜体等不能可靠表达的公式整式沿用 MathML 路径。真实样本中分别导入 838 次和 142 次，另有 35 次资产缓存命中。上述三轮使用版本升号前的同一转换代码；简单 1001 式合成样本此前曾用时 31.708 秒，实际速度随公式和机器环境变化。只读输入可正常转换，原件保持只读。详细阶段计时及检查范围见实测报告。

## 按需调整公式格式

可选择全部正文、仅行内或仅行外公式，分别设置行内/行外字号和段落行距。行外右编号支持章号、节号、起始序号和格式预览，也可补充或更新已有 MathType 原生编号。章、节值由用户指定。

高级设置提供 Word 原生公式（OMML）转换开关，以及表格、文本框、页眉页脚、脚注尾注的实验性转换。额外位置只替换公式，设置方法与适用范围见[使用说明](docs/使用说明.md)。

<details open>
<summary>查看高级设置界面</summary>

![DocxMathTypeBatch 高级设置界面](docs/images/ui-advanced.png)

设置在当前程序会话中保留，重启后恢复默认。

</details>

## AI Agent Skill

安装完整 Skill 后，可直接向本机 Agent 提出要求。

> 将这份 DOCX 的公式转成 MathType，正文行外公式添加右编号，结果另存为新文件。

1. [下载完整 Skill v0.2.2](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1.5/docx-latex-to-mathtype-skill-v0.2.2.zip)，将其中的 `docx-latex-to-mathtype` 文件夹放入所用 Agent 的 Skill 目录并重新加载。
2. 向 Agent 提供文档和转换要求。本机需安装 Word、MathType，并支持运行 Windows PowerShell 5.1 或 PowerShell 7 脚本。

完整包内含 v0.1.5 转换器 EXE，可随整个文件夹移动。日常图形界面使用直接下载 Windows 程序包。安装与调用详情见[Agent 使用说明](docs/使用说明.md#使用-agent-自动调用)和[Skill 调用参数](skills/docx-latex-to-mathtype/references/cli.md)。

## 运行环境与支持范围

| 项目              | 要求或说明                                                          |
| --------------- | -------------------------------------------------------------- |
| Windows         | Windows 10（1703 及以上）或 Windows 11，x64                           |
| Word / MathType | 已正确安装并注册的 64 位桌面 Microsoft Word、完整 MathType，包含适配 64 位 Word 的组件 |
| LaTeX 输入        | 正文中使用 `$...$`、`$$...$$`、`\(...\)` 或 `\[...\]` 的公式              |
| Word 原生公式       | 默认与 LaTeX 一起转换，可在高级设置关闭                                        |
| 反向导出            | 支持 MathType → LaTeX，可选择行外定界符，保留 Word 原生公式                      |

LaTeX 以 ASCII 公式为基础，支持常见命令与多行结构的兼容适配。公式独占段落时按行外处理，夹在正文中时按行内处理。中文正文保留，适配后仍含不支持字符的公式保留原式。完整写法见[公式书写与兼容性](docs/MathType兼容LaTeX书写规范.md)。

遇到“部分完成”时，可在“日志”中查看问题公式的原文页码、片段、原因及建议。高公式显示不全时可调整行距，过长公式需自行调整。修订、限制编辑及实验性位置的处理说明见[使用说明](docs/使用说明.md)。

## 文档与反馈

[使用说明](docs/使用说明.md) · [公式兼容性](docs/MathType兼容LaTeX书写规范.md) · [实测报告](docs/实测报告.md) · [更新记录](CHANGELOG.md) · [开发说明](docs/轻量化转换大纲.md) · [参与贡献](CONTRIBUTING.md)

欢迎通过 [Issues](https://github.com/yuxmic/DocxMathTypeBatch/issues) 反馈问题。请提供程序、Windows、Word、MathType 版本和复现步骤，示例、截图及日志提交前请去除私人信息。

本工具采用 [MIT 许可证](LICENSE)，允许使用、修改、分发和商用，须保留版权及许可声明。Microsoft Word 和 MathType 的使用遵循各自许可条款。
