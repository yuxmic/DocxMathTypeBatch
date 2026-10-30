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
| `\mathring{…}` | `\overset{\circ}{…}`，使用 MathType 的上方圆圈排版 |
| `aligned`、`align*`、`split` 环境 | `array`，按实际顶层列数使用 `rl` 交替列格式 |
| `alignedat`、`alignat*` 环境 | 消费字面的列组数参数，再按实际顶层列数使用 `array` |
| `gathered`、`gather*` 环境 | 居中的 `array` |
| `subarray` 环境 | 保留原有列格式参数，改用 `array` |

`\mathup` 与 `\mathrm` 的别名关系见 [unicode-math 文档](https://mirrors.ctan.org/macros/unicodetex/latex/unicode-math/unicode-math.pdf)。多行结构和点号的用途依据 [amsmath 用户指南](https://texdoc.org/serve/amsldoc/0)，定义等号依据 [mathtools 文档](https://tug.ctan.org/macros/latex/contrib/mathtools/mathtools.pdf)。不同环境统一使用 MathType 的数组排版，保留数学内容、分行和分列，不复制 LaTeX 宏包的精细字距。

规则按完整命令名匹配，`\bmfoo` 和转义的 `\\bm` 保持原样；`\text`、`\mbox`、`\operatorname` 等文字参数内部保持原样。原本已支持的 `\operatorname`、`\dfrac`、`\substack`、`\xrightarrow`、`\mathscr` 等不额外替换。本轮不移除 `\tag`、`\label`，也不将带自动编号的 `align` 当作无编号的 `align*`；自动编号继续通过主界面的原有设置完成。

`\boxed`、`\fbox` 和含 `\&` 的复杂文字在本机仍有未解析红字；没有验证出可靠的等效写法，保持 MathType 的部分解析结果。任意自定义命令仍需具体定义。本程序不增加未解析命令的逐式导出检查，也不因此回滚或丢弃正常转换结果；这是 MathType 原生的[部分解析行为](https://docs.wiris.com/mathtype-7-interface/mathtype-7-ui-basic-usage)。

开发回归覆盖 42 条正文、行内、跨段及表格公式，核对 PDF 无红字、反向导出的内容、嵌套行列及原件未变。`native-api-probe --compatible-tex-self-test` 不启动 Office，单独测试命令边界并计量正式改写函数；本机一次测试对约 965 字符的样本改写 1 万次耗时约 67 毫秒，此项只统计内存文本处理。

## Unicode 与 Word 原生顶饰适配

2026-10-28 的本地修复在正文和实验区域共用入口启用以下处理，仅修改导入副本：

- `Ŷ`、`Ȳ`、`Ȭ` 等预组合字母按 Unicode 规范分解为基础字符和已支持顶饰；`Ŷ` 等组合写法也构造成明确的顶饰结构。分解使用 Windows 的 [FoldStringW / MAP_COMPOSITE](https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-foldstringw)，不做兼容字体折叠。
- LaTeX 中可转成 ASCII 的字母顶饰改写为嵌套命令后再检查字符范围，例如 `Ŷ` → `\hat{Y}`、`Ȭ` → `\bar{\tilde{O}}`。文字参数内部不改写，未适配字符继续保留原式。
- Word 原生公式在 Office 生成 MathML 后，将顶饰位置的组合符转换为 MathType 能识别的符号码；预组合字母和跨相邻符号节点的组合写法构造成 `mover`。基础字符的字体属性保留，`XŶ` 只装饰最后的 `Y`。
- 当前顶饰表包含抑扬符、波浪线、横线、短音符、单点至四点、圆圈、钩号、尖音符、重音符及左右/双向箭头。没有改写规则的命令继续交给 MathType 原有解析器。

集中回归验证了 19 个正文/表格公式的页面、实际原生结构及粗体保留，另有 2 个未适配字符样本保留原式。检查使用 `test_tex_compatibility.py accent-fixture/accent-check` 和开发专用的 `native-api-probe --accent-export`；正式正向转换的缓存回显和成功判定保持原有规则。MathType 的短音符原生导出可能包含嵌套 `$`，反向转换会保留该对象并报告 `UNSUPPORTED_OR_INVALID_NATIVE_EXPORT`，其他公式继续导出；集中回归另含 `accent-reverse-check`。宽顶饰模板在本机 PDF 中可能缺失的渲染问题也仍需另行处理。
