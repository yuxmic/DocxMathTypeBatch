# DocxMathTypeBatch

将 Word 文档中的 LaTeX 公式批量转换为可编辑的 MathType 对象。

DocxMathTypeBatch 是一款 Windows 桌面工具，适合在 Word 中使用 LaTeX 编写公式、需要统一转换或整理公式格式的场景。程序调用本机的 Word 和 MathType 完成转换，将结果保存为新的 `.docx` 文件，保留原文档。

[使用说明](docs/使用说明.md) · [公式书写与兼容性](docs/MathType兼容LaTeX书写规范.md)

## 下载 v0.1

- **[下载 Windows x64 ZIP](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1/DocxMathTypeBatch-v0.1-win-x64.zip)**：推荐，包含程序、简要使用说明和 MIT 许可证，解压后运行。
- [下载单文件 EXE](https://github.com/yuxmic/DocxMathTypeBatch/releases/download/v0.1/DocxMathTypeBatch.exe)
- [版本说明与校验文件](https://github.com/yuxmic/DocxMathTypeBatch/releases/tag/v0.1)

运行前请安装下方列出的 Word 和 MathType。程序尚未进行代码签名，首次运行可能出现 Windows 安全提示。

## 主要功能

- **批量转换**：拖入多个文档，或添加文件夹中的 DOCX，按队列依次处理。
- **选择公式范围**：转换全部公式、仅行内公式或仅行外公式。
- **调整公式格式**：字号和段落行距默认跟随原文，也可分别指定。
- **添加行外编号**：支持章号、节号、起始序号和编号格式，可补充或更新已有 MathType 原生编号。
- **基本反向导出**：将正文中的 MathType 公式导出为 LaTeX 文本。
- **查看处理结果**：逐文档查看转换详情；无法转换的公式保留原式，并记录原因。支持安全停止当前批次。

## 运行环境

| 项目 | 要求 |
| --- | --- |
| 操作系统 | Windows x64 |
| Word | 已安装并正确注册的 64 位桌面版 Microsoft Word |
| MathType | 完整桌面版，包含适配 64 位 Word 的组件 |

运行编译后的程序无需安装 Python、PowerShell 或 .NET。

## 使用方法

1. 下载上方 ZIP，解压后运行 `DocxMathTypeBatch.exe`。
2. 保存并关闭待转换文档，将一个或多个 `.docx` 文件拖入程序。
3. 选择转换方向和公式范围，按需调整字号、行距与右编号。默认转换全部公式，字号和行距跟随原文，不添加编号。
4. 选择输出位置，点击“开始转换”。结果另存为新文件，重名时自动添加序号。
5. 在列表中查看完成状态，通过“日志”查看具体原因。打开结果文档后，可双击新生成的 MathType 对象继续编辑。

反向转换时，切换到 **MathType → LaTeX**，可选择行外公式的输出定界符。详细设置见[使用说明](docs/使用说明.md)。

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
| Word 自带公式（OMML） | 保持原样 |
| 表格、文本框、页眉页脚、脚注尾注 | 不转换 |

公式代码目前以 **ASCII LaTeX** 为支持范围，包含中文等非 ASCII 字符的公式会保留原式；中文正文可正常保留。程序不自动拆分或缩放过长公式，也不将缺少定界符的裸 LaTeX 作为转换目标。兼容写法与示例见[公式书写与兼容性](docs/MathType兼容LaTeX书写规范.md)。

## 问题反馈

反馈时请提供 Windows、Word 和 MathType 版本、复现步骤，以及脱敏后的最小公式或文档示例。遇到转换失败，可附上“转换详情”中的相关日志，并先移除其中的文档内容和个人路径。

## 许可

本项目采用 [MIT 许可证](LICENSE)，允许使用、修改、分发和商用，须保留版权及许可声明。Microsoft Word 和 MathType 的使用遵循各自的许可条款。

源码构建和开发约定见[开发说明](docs/轻量化转换大纲.md)。
