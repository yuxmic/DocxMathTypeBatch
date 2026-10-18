# MathType 兼容的 LaTeX 公式书写规范

供共同撰写 Word 文档时使用。以下写法已在当前 MathType 环境中用 **Toggle TeX** 实测；其他 LaTeX 命令不能仅凭语法正确就认为 MathType 一定支持。

项目的 `build/bin/docx-math-convert.exe` 单向原生批处理与 Toggle TeX 是不同入口。原生 TeX Input 导入 `\text{汉字}` 可能产生错误符号，当前批处理保留任何非 ASCII 公式内容并报告 `UNSUPPORTED_NON_ASCII_TEX`；中文正文和位置扫描仍正常。优先用 ASCII LaTeX 命令表达符号；可靠 Unicode 公式导入须另行验收，不能用 API 返回成功代替页面检查。

## 基本写法

- 行内公式：`$a+b=c$`
- 独立成行的公式：`\[\frac{a}{b}=c\]`
- 在 Word 中选中完整公式（包括定界符），再使用 MathType 标签卡的 **Toggle TeX**。一次只转换选中的内容；重复点击会反向转换。

## 多行、多列公式

优先使用 `array`。`{rl}` 表示两列分别右对齐、左对齐；`&` 分列，`\\` 换行。每行的列数应与列格式一致。

```latex
\[\begin{array}{rl}a&=b+c\\d&=e+f\end{array}\]
```

等号单独占一列时用 `{rcl}`：

```latex
\[\begin{array}{rcl}a&=&b+c\\d&=&e+f\end{array}\]
```

本机还验证了 `cases`（分段函数）、`pmatrix`（带圆括号矩阵）、空单元格和嵌套 `array` 可以转换。公式内部的实际分行使用 `\\`。

C++ 批处理支持完整结构化公式跨 Word 段落或软换行；`\\` 后紧跟 Word 换行时，仍整体识别到右定界符，仅将 Word 换行合并为源码空白，保留 TeX 分行和所有分支。跨段识别不越过表格、域和下一公式。手动使用 Toggle TeX 时，仍建议选中一段连续的完整公式。

## 已知限制与检查

- 手工 Toggle TeX 仍不建议直接使用 `\begin{aligned}...\end{aligned}`，本机测试会标红。2026-10-17 的 C++ 批处理已在导入副本中自动改写为 `array`，按顶层列数保留同一行的多组等式、嵌套环境、`&` 和 `\\`，无需修改 Word 原文。
- 原生 TeX Input 实测包含 `\text{A\&B}` 的公式仍可能回显命令，且 API 返回成功；复杂宏需要核对实际页面和原生导出。开发时不要仅按 API 返回码判断公式正确。
- `array` 的列格式竖线 `|` 和 `\hline` 在本机测试中没有显示；需要分隔线时应在 MathType 中手动设置并核对。
- 转换后逐项检查：无红色残留命令、行列内容与对齐正确、公式可双击编辑。手动 Toggle TeX 不会自动继承 Word 正文的字号；C++ 批处理默认读取源字号并应用到新公式，字体采用默认配置，段落格式沿用原文。
