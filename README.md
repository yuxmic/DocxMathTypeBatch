# DocxMathTypeBatch

将 Word 文档中的 LaTeX 和 Word 原生公式批量转换为可编辑的 MathType 对象。

DocxMathTypeBatch 是一款 Windows 桌面工具，适合在 Word 中使用 LaTeX 编写公式、需要统一转换或整理公式格式的场景。程序调用本机的 Word 和 MathType 完成转换，将结果保存为新的 `.docx` 文件，保留原文档。

[使用说明](docs/使用说明.md) · [公式书写与兼容性](docs/MathType兼容LaTeX书写规范.md)

## 下载 v0.1.1

- **[下载 Windows x64 ZIP](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1.1/DocxMathTypeBatch-v0.1.1-win-x64.zip)**：推荐，包含程序、简要使用说明和 MIT 许可证，解压后运行。
- [下载单文件 EXE](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1.1/DocxMathTypeBatch.exe)
- [下载配套 Skill v0.1](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1.1/docx-latex-to-mathtype-skill-v0.1.zip)：供 Agent 自动调用，可选安装。
- [版本说明与校验文件](https://github.com/yuxmic/DocxMathTypeBatch/releases/tag/v0.1.1)

运行前请安装下方列出的 Word 和 MathType。程序尚未进行代码签名，首次运行可能出现 Windows 安全提示。

## 主要功能

- **批量转换**：拖入多个文档，或添加文件夹中的 DOCX，按队列依次处理。
- **正文原生公式转换**：默认将 Word 原生公式（OMML）与 LaTeX 一起转换，共用格式和编号设置。
- **选择正文范围**：所有正文、仅行内或仅行外。
- **调整正文公式格式**：字号和段落行距默认跟随原文，也可分别指定。
- **添加行外编号**：支持章号、节号、起始序号和编号格式，可补充或更新已有 MathType 原生编号。
- **基本反向导出**：将正文中的 MathType 公式导出为 LaTeX 文本。
- **查看处理结果**：逐文档查看转换详情；无法转换的公式保留原式，并记录原因。支持安全停止当前批次。

## 运行环境

| 项目 | 要求 |
| --- | --- |
| 操作系统 | Windows x64 |
| Word | 已安装并正确注册的 64 位桌面版 Microsoft Word |
| MathType | 完整桌面版，包含适配 64 位 Word 的组件 |

普通界面操作无需安装 Python、PowerShell 或 .NET；可选的 Skill 脚本使用 Windows PowerShell 5.1 或 PowerShell 7。

## 使用方法

1. 下载上方 ZIP，解压后运行 `DocxMathTypeBatch.exe`。
2. 保存并关闭待转换文档，将一个或多个 `.docx` 文件拖入程序。
3. 选择转换方向和正文范围，按需调整字号、行距与右编号。默认选择“所有正文”，字号和行距跟随原文，不添加编号。
4. 选择输出位置，点击“开始转换”。结果另存为新文件，重名时自动添加序号。
5. 在列表中查看完成状态，通过“日志”查看具体原因。打开结果文档后，可双击新生成的 MathType 对象继续编辑。

反向转换时，切换到 **MathType → LaTeX**，可选择行外公式的输出定界符。详细设置见[使用说明](docs/使用说明.md)。

## Agent 自动调用

EXE 与独立的 [docx-latex-to-mathtype Skill](skills/docx-latex-to-mathtype/SKILL.md) 分别安装：

1. 下载上方 Skill ZIP，将其中的 `docx-latex-to-mathtype` 文件夹安装到所用 Agent 的 Skill 目录，并重新加载 Skill。
2. 正常打开 v0.1.1 EXE 一次，程序会在本机登记自身位置；也可向 Agent 提供一次 EXE 位置进行连接。
3. 向 Agent 提出要求，例如：“将这份 DOCX 的公式转成 MathType，正文行外公式添加右编号，结果另存为新文件。”

移动 EXE 后重新打开或重新连接即可，Skill 目录可以与 EXE 分开存放。

普通界面操作无需安装 Skill。自动调用需要 Agent 能在装有 Word 和 MathType 的 Windows 本机执行脚本；原 v0.1 不支持连接协议。安装与使用见[使用说明](docs/使用说明.md#使用-agent-自动调用)。

## 额外位置的公式转换（实验性）

正文中的 LaTeX 和 Word 原生公式（OMML）默认共同转换，参与字号、行距、行外布局和右编号；不需要额外开启 OMML。

“高级设置”右侧提供默认关闭的实验性功能：两个基本方向均可按需开启表格、文本框、页眉页脚及脚注尾注中的公式转换。

实验项只替换公式，不应用字号、段落行距或右编号设置，仅在当前程序会话中保留勾选状态。表格、文本框选项也作用于已启用的其他文档区域，例如页眉内的文本框需要同时开启“页眉页脚”和“文本框”。

## 公式格式与支持范围

支持 `$...$`、`$$...$$`、`\(...\)` 和 `\[...\]`。例如，Word 正文可以写为：

```latex
正文中的 $a+b=c$ 作为行内公式处理。

\[\frac{x+1}{y+2}\]
```

公式独占段落时按行外处理，夹在普通文字中时按行内处理。支持 `cases`、`array`、矩阵等完整多行结构；`aligned` 会在导入时转换为 MathType 兼容的形式。

| 文档内容 | 当前处理方式 |
| --- | --- |
| 正文中带定界符的 LaTeX | 按所选范围转换 |
| 正向转换时已有的 MathType 对象 | 保留公式内容，可统一独立行外对象的布局并补充或更新编号 |
| 正文中的 Word 自带公式（OMML） | 默认正向转换，并参与正文格式和编号；反向保持 |
| 表格、文本框、页眉页脚、脚注尾注 | 默认保持；可单独开启实验性转换 |

公式代码目前以 **ASCII LaTeX** 为支持范围，包含中文等非 ASCII 字符的公式会保留原式；中文正文可正常保留。程序不自动拆分或缩放过长公式，也不将缺少定界符的裸 LaTeX 作为转换目标。兼容写法与示例见[公式书写与兼容性](docs/MathType兼容LaTeX书写规范.md)。

## 问题反馈

个别公式或实验区域处理失败时，程序会保留原内容并输出已经核验的转换结果，文件状态显示“部分完成”。可打开结果，并在“转换详情”查看跳过的位置及原因。

反馈时请提供 Windows、Word 和 MathType 版本、复现步骤，以及脱敏后的最小公式或文档示例。遇到转换失败，可附上“转换详情”中的相关日志，并先移除其中的文档内容和个人路径。

## 许可

本项目采用 [MIT 许可证](LICENSE)，允许使用、修改、分发和商用，须保留版权及许可声明。Microsoft Word 和 MathType 的使用遵循各自的许可条款。

源码构建和开发约定见[开发说明](docs/轻量化转换大纲.md)。
