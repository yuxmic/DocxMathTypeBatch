# MathType 兼容的 LaTeX 公式书写规范

供共同撰写 Word 文档时使用。基本与 `array` 写法已在当前 MathType 环境中用 **Toggle TeX** 实测；下方自动兼容规则在本程序的原生 TeX Input 路径启用。其他 LaTeX 命令不能仅凭语法正确就认为 MathType 一定支持。

项目的 `build/bin/docx-math-convert.exe` 单向原生批处理与 Toggle TeX 是不同入口。原生 TeX Input 导入 `\text{汉字}` 可能产生错误符号，当前批处理先适配已支持的 Unicode 顶饰，适配后仍含非 ASCII 字符的公式保留原式并报告 `UNSUPPORTED_NON_ASCII_TEX`；中文正文和位置扫描仍正常。优先用 ASCII LaTeX 命令表达符号；其他 Unicode 公式导入须另行验收，不能用 API 返回成功代替页面检查。

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

- 批处理遇到未转义的 `%` 注释时保留原式并报告 `UNSUPPORTED_TEX_COMMENT`，避免 MathType 把注释排成公式内容。需要显示百分号时使用 `\%`。
- 手工 Toggle TeX 仍不建议直接使用 `\begin{aligned}...\end{aligned}`，本机测试会标红。2026-10-17 的 C++ 批处理已在导入副本中自动改写为 `array`，按顶层列数保留同一行的多组等式、嵌套环境、`&` 和 `\\`，无需修改 Word 原文。
- 原生 TeX Input 实测包含 `\text{A\&B}` 的公式仍可能回显命令，且 API 返回成功；复杂宏需要核对实际页面和原生导出。开发时不要仅按 API 返回码判断公式正确。
- `array` 的列格式竖线 `|` 和 `\hline` 在本机测试中没有显示；需要分隔线时应在 MathType 中手动设置并核对。
- 转换后逐项检查：无红色残留命令、行列内容与对齐正确、公式可双击编辑。手动 Toggle TeX 不会自动继承 Word 正文的字号；C++ 批处理默认读取源字号并应用到新公式，字体采用默认配置，段落格式沿用原文。

## 已启用的自动兼容改写

2026-10-22 使用原生 TeX Input 与 Word 导出的 PDF 对照了 40 组常见写法，并补查了定义等号和转义文字。以下规则在 `compatibleTex` 中启用，正文和实验区域共用；只修改传给 MathType 的文本副本，报告中的源公式保持用户原写法。

| 原写法 | 导入 MathType 的写法 |
| --- | --- |
| `\bm{…}` | `\boldsymbol{…}` |
| `\textnormal{…}`、`\textup{…}`、`\textmd{…}` | `\text{…}` |
| `\mathup{…}` | `\mathrm{…}` |
| `\lvert`、`\rvert` | `\vert` |
| `\mleft`、`\mright` | `\left`、`\right` |
| `\dotsc`、`\dotso` | `\ldots` |
| `\dotsm` | `\cdots` |
| `\coloneqq`、`\eqqcolon` | 组合关系符 `:=`、`=:`，字距由 MathType 排版 |
| `\widetilde{…}` | `\tilde{…}`，由 MathType 根据参数生成单字顶饰或宽模板 |
| `\mathring{…}`、`\ring{…}` | `\overset{\circ}{…}`，使用 MathType 的上方圆圈排版 |
| `\widebar{…}`、`\overbar{…}` | `\overline{…}`，保留整组覆盖 |
| `\widecheck{…}`、`\widebreve{…}` | 同形顶饰居中，宽度由 MathType 的可用字形决定 |
| `\overarc{…}`、`\overparen{…}`、`\wideparen{…}` | MathType 原生上弧模板 |
| `\enspace`、`\enskip`、`\hspace{…}`、`\kern`、`\mkern` | 将明确写出的长度换成 MathType 可用的数学间距 |
| `aligned`、`align*`、`split` 环境 | `array`，按实际顶层列数使用 `rl` 交替列格式 |
| `alignedat`、`alignat*` 环境 | 消费字面的列组数参数，再按实际顶层列数使用 `array` |
| `gathered`、`gather*` 环境 | 居中的 `array` |
| `subarray` 环境 | 保留原有列格式参数，改用 `array` |

`\mathup` 与 `\mathrm` 的别名关系见 [unicode-math 文档](https://mirrors.ctan.org/macros/unicodetex/latex/unicode-math/unicode-math.pdf)。多行结构和点号的用途依据 [amsmath 用户指南](https://texdoc.org/serve/amsldoc/0)，定义等号依据 [mathtools 文档](https://tug.ctan.org/macros/latex/contrib/mathtools/mathtools.pdf)。不同环境统一使用 MathType 的数组排版，保留数学内容、分行和分列，不复制 LaTeX 宏包的精细字距。

规则按完整命令名匹配，`\bmfoo` 和转义的 `\\bm` 保持原样。文字参数内仅适配已识别的双锐音字母，如 `\text{\H{o}}`。需要 MathML 的兼容公式先在独立的临时 MathType 对象中解析，再把最终结构一次写进成品。普通公式直接导入。自动编号继续通过主界面的设置完成。

`\boxed`、`\fbox` 和含 `\&` 的复杂文字在本机仍有未解析红字；没有验证出可靠的等效写法，保持 MathType 的部分解析结果。任意自定义命令仍需具体定义。本程序不增加未解析命令的逐式导出检查，也不因此回滚或丢弃正常转换结果；这是 MathType 原生的[部分解析行为](https://docs.wiris.com/mathtype-7-interface/mathtype-7-ui-basic-usage)。

开发回归覆盖 42 条正文、行内、跨段及表格公式，核对 PDF 无红字、反向导出的内容、嵌套行列及原件未变。`native-api-probe --compatible-tex-self-test` 不启动 Office，单独测试命令边界并计量正式改写函数；本机一次测试对约 965 字符的样本改写 1 万次耗时约 67 毫秒，此项只统计内存文本处理。

## Unicode 与 Word 原生顶饰适配

2026-10-28 的本地修复在正文和实验区域共用入口启用以下处理，仅修改导入副本：

- `Ŷ`、`Ȳ`、`Ȭ` 等预组合字母按 Unicode 规范分解为基础字符和已支持顶饰；`Ŷ` 等组合写法也构造成明确的顶饰结构。分解使用 Windows 的 [FoldStringW / MAP_COMPOSITE](https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-foldstringw)，不做兼容字体折叠。
- LaTeX 中可转成 ASCII 的字母顶饰改写为嵌套命令后再检查字符范围，例如 `Ŷ` → `\hat{Y}`、`Ȭ` → `\bar{\tilde{O}}`。数学和文字中的 `Ő/ő/Ű/ű` 使用 MathML 保留字形与文字样式。
- Word 原生公式在 Office 生成 MathML 后，将顶饰位置的组合符转换为 MathType 能识别的符号码；预组合字母和跨相邻符号节点的组合写法构造成 `mover`。基础字符的字体属性保留，`XŶ` 只装饰最后的 `Y`。
- 当前顶饰表包含抑扬符、波浪线、横线、短音符、单点至四点、圆圈、钩号、尖音符、重音符及左右/双向箭头。没有改写规则的命令继续交给 MathType 原有解析器。

反向导出会把厂商给短弧、箭头和下横线生成的嵌套定界符改为普通命令，也会修正上方圆圈出现的双上标。厂商无法导出的公式会保留 MathType 对象。

### Word 原生宽空格和上顶饰

保留普通文字词间空格、数学变量之间有宽度的普通空格和不换行空格。逗号后由 Word 自行处理的普通数学空格交给 MathType 自动排版。U+2001 宽空格使用等宽的 U+2003，U+2009 等已支持的窄空格沿用原生写法，其他明确宽度按 MathType 可用的档位取近似。公式内 Tab 依赖 Word 制表位，欧甘空格 U+1680 有可见字形，U+2061–U+2064 有运算含义。本机 MathType 无法保留这些情况时保留源公式。

标准、快速和极速共用这些规则。Word 的上横线、双横线、上弧、长箭头和居中点饰使用可用的原生结构。覆盖整组的左鱼叉和方括号在本机无法保真时保留原式。可用原生花括号模板表示的整组花括号会继续转换。

直接写入的 Unicode 双横线、倒短弧和单字鱼叉，以及叠在横线或尖帽上的点饰，都按顶饰结构导入。普通文字中的组合上横线保留文字样式。LaTeX 的双锐音字母支持数学变量和文字写法。单字符上方方括号使用 MathType 的固定宽度字形。带宽度参数的 `\overarc` 会保留原式。

普通文字和数学变量中的不换行空格使用显式间距，避免预览出现问号。下标后的空格可以保留。逗号和二元加号旁的显式空格会补回 MathType 默认排版的部分自动间距，实际宽度仍随字号和 MathType 可用档位取近似。Office 为 `sin(x)` 等原生函数生成的应用标记允许导入，原式中单独写入的不可见运算符仍按保真保护处理。

标准和快速模式在 Word 保存工作副本后，会把保留公式内的 Tab 写回 Word 重开时能识别的文本形式。该处理只涉及公式内部，段首 Tab 和旧编号前空格仍按原规则清理。直接在 Word 中再次保存这些原生公式时，Word 仍可能自行丢弃其中的 Tab。

本机用 188 式真实样本加 8 式合成回归运行三档。三档各成功 196 式、生成 38 个连续右编号，原生公式数据均无 U+FFFD 替代字符。核对了宽空格、窄空格、上横线、波浪线、尖帽、倒帽、普通运算符及原件的 45 条下横线。倒帽正向能显示，本机 MathType 的 LaTeX 回导仍可能报错。
