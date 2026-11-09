# Skill v0.2.3 调用约定

要求包内 EXE 支持 `CliContractVersion: 1`、报告 `SchemaVersion: 1`。Skill 包内包含已匹配的 v0.1.6 转换器 `bin/DocxMathTypeBatch.exe`，更新时替换完整包。文档样本、本机配置和日志由用户在本机管理。

## 程序路径与查询

转换脚本使用相对于 Skill 脚本目录的 `../bin/DocxMathTypeBatch.exe`，不读取注册表，也不搜索安装位置。整个 Skill 目录移动后无需重新连接，普通 GUI 启动也不写入程序路径登记。

`connect.ps1` 仅无参数诊断包内程序身份及接口，不接受外部 EXE 路径。正常转换直接调用 `convert.ps1`，无需额外检测程序位置或先连接。包内 EXE 缺失时返回 `EXECUTABLE_NOT_FOUND`，重新解压完整 Skill 包即可。

连接前先读取 EXE 文件资源的产品名称；旧版或其他程序不会被拿来探测不认识的 GUI 参数。`--describe --json` 在 5 秒内返回静态身份与能力，不启动 Word。查询超时只结束该查询进程。

## convert.ps1 参数

| 参数 | 作用 |
|---|---|
| `-InputPath`、`-OutputPath` | 必填；输入已存在，输出全新，父目录已创建 |
| `-Direction ToMathType/ToLatex` | 默认正向，反向使用 ToLatex |
| `-Scope all/inline/display` | 默认所有基本正文；同样筛选已开启的额外区域 |
| `-FontSize 12` | 正向指定 1–256 pt。标准与快速省略时跟随原文，极速必须指定 |
| `-RightNumbers` | 标准与快速补号或更新编号，极速只新增默认右编号。默认关闭 |
| `-NumberFormat '(#C1-#E1)'` | 新增编号格式；默认 `(#E1)` |
| `-NumberStart`、`-NumberChapter`、`-NumberSection` | 1–999999 的整数，默认 1 |
| `-CustomNumber` | 使用自定义编号格式，配合 RightNumbers |
| `-InlineSpacing`、`-DisplaySpacing` | 默认 follow；single、onehalf、double、exact:20、atleast:20、multiple:1.25 |
| `-DisplayDelimiter bracket/dollar` | 反向行外使用 `\[...\]` 或 `$$...$$`；行内使用 `$...$` |
| `-Tables`、`-Textboxes`、`-Headers`、`-Notes` | 显式启用四类额外位置，默认关闭 |
| `-SkipOmml` | 标准或快速正向转换时保留 Word 原生公式，只转换 LaTeX。默认不传 |
| `-ScanReportPath`、`-TargetId @(2,5)` | 配对传入，只转换本次扫描清单中的指定公式 |
| `-FastConversion` | 快速模式，通过 Word 转换，跳过转换前内容及语法核验、逐式复查、保存重开及资产核验。默认关闭 |
| `-UltraConversion` | 极速模式，直接转换正文 LaTeX 和 OMML，不启动 Word，须指定 `-FontSize`。默认关闭 |
| `-TaskDirectory` | 可选全新任务目录；默认创建本机用户数据下的独立目录 |
| `-TimeoutSeconds` | 无进展超时，5–3600 秒，默认 120；不是整份任务总时限 |

编号格式支持 `#C1`、`#S1`、`#E1`，末尾可为 `a/A/i/I`，并且仅含一个 `#E`。章、节、起始值作用于目标序列；已有编号保留其原生显示格式。反向不支持字号、右编号和段落行距设置。

额外位置仅转换公式。OMML 默认参与正向转换。用户要求保留 Word 原生公式时，在标准或快速模式传入 `-SkipOmml`，同时作用于正文及启用的额外位置，不计为拒绝或未完成。反向始终保留 OMML。该选项要求 EXE 能力含 `OptionalOmmlConversion: true`，底层参数为 `--skip-omml`，报告 `ConvertOmml` 表示本次有效的正向 OMML 设置。表格或文本框开关需要其父区域也启用。保护、修订和普通域等仍构成边界。LaTeX 文本以 ASCII 为基础，支持顶饰适配的程序可先将 `Ŷ`、`Ŷ` 等写法转换为命令，其余未适配字符保留原文。标准和快速的 OMML 导入使用 Office 原生转换资源。是否成功以本次报告为准。

## 只转换指定公式

`scan.ps1` 支持 `-InputPath`、`-TaskDirectory` 和 `-TimeoutSeconds`。只读扫描成功时，`FINISHED.Status` 为 `SCANNED`，`ReportPath` 指向本机 JSON 清单。若文档受保护，扫描返回 `SKIPPED_PROTECTED_DOCUMENT`。根据 `Items` 的 `Source`、`Kind`、`Paragraph` 和同段的短上下文 `Before/After` 确定目标。ID 对应公式在当前输入中的一次出现，源文档的 `SourceSHA256` 及公式的 Word 范围和源内容会在转换前复核。

```powershell
# 读取清单，先用源码和上下文确认 ID，再传入所选 ID。
$scan = & (Join-Path $skillDir 'scripts/scan.ps1') -InputPath $inputDocx |
    ForEach-Object { $_ | ConvertFrom-Json } | Where-Object Event -eq FINISHED
if ($scan.Status -ne 'SCANNED') { throw "扫描未完成，状态为 $($scan.Status)" }
(Get-Content -LiteralPath $scan.ReportPath -Raw -Encoding utf8 | ConvertFrom-Json).Items |
    Where-Object Status -eq READY | Select-Object Id, Kind, Paragraph, Source, Before, After
```

选定后再调用转换。下面的 `2,5` 是清单里的 ID，文档中可见的公式编号可用来核对上下文。

```powershell
& (Join-Path $skillDir 'scripts/convert.ps1') -InputPath $inputDocx -OutputPath $outputDocx -ScanReportPath $scan.ReportPath -TargetId @(2,5)
```

此入口要求标准正向正文和默认范围 `all`，目标可混选行内、行外、LaTeX 和 OMML。可传固定字号，默认跟随原文。只替换公式本体，保留公式外的文字、空白、手打编号及原段落设置。使用默认行距，关闭额外位置、跳过 OMML 和编号选项。跨段的 LaTeX 不可选。结果返回 `SelectedTargetIds` 与 `UnselectedFormulas`，未选项不计为失败。

`TARGET_SOURCE_CHANGED` 表示文档在扫描后发生变化，`TARGET_SELECTION_MISMATCH` 表示所选位置或源码未通过复核。遇到这两种结果时重新扫描并核对目标。扫描或选择失败时结束本次调用，避免对全篇执行转换。

## 极速模式

只在用户明确要求极速时传入 `-UltraConversion`。极速适合批量转换正文普通段落中的公式。它直接读写 DOCX 并调用 MathType 生成可编辑对象，运行时不启动 Word。仍需按本 Skill 的安装要求准备 Word 和 MathType，OMML 转换会用到 Office 的转换资源。

- **转换范围**固定为 `ToMathType` 和 `all`，同时转换行内、行外的 LaTeX 和 OMML。识别 `$...$`、`$$...$$`、`\(...\)`、`\[...\]`，支持同一普通段落内多式混排及跨连续文本片段的 LaTeX。域、内容控件和段落边界会分隔 LaTeX 扫描。表格、文本框、页眉页脚、脚注尾注及已有 MathType 保留在文档中。
- **字号和行距**须显式传入 `-FontSize`，同一字号作用于所有新公式。行距使用默认 `follow`。
- **右编号**使用 `-RightNumbers`，为独立行外式从（1）开始顺序编号。启用编号时，包含待编号行外式的文档须单节、单栏、有页宽和页边距，且没有既有 MathType 原生编号。普通段落只有一式和形如 `(27)`、`（27）`、`(3-2)` 的手打尾号时，转换成功后清理尾号及多余空白，即使未启用右编号。单式生成失败时保留原式和尾号。
- **参数组合**使用默认编号格式及起始值。反向、仅行内或仅行外、`-FastConversion`、`-SkipOmml`、额外位置、调整行距和自定义编号均需另选模式。文档的编号条件不符合时，可在用户接受的情况下省略 `-RightNumbers`。需要自定义编号或更新引用时使用标准或快速模式。

```powershell
# 极速转换全部正文公式，字号 12 pt，行外按默认格式补号。
& (Join-Path $skillDir 'scripts/convert.ps1') -InputPath $inputDocx -OutputPath $outputDocx -UltraConversion -FontSize 12 -RightNumbers
```

极速能力由 EXE 的 `ConversionModes` 包含 `Ultra` 声明，底层使用 `--worker-convert`、`size=字号` 和 `--ultra`。极速没有逐式复查、版式提醒、保存重开、恢复重试或自动切换模式。任务结果交付后需查看输出。

## 任务结果与诊断

脚本只向标准输出发送 JSON Lines，不混入 worker 原始日志，也不启动 GUI。`STARTED` 给出 `Output`、`TaskDirectory`、`StopFile`、`LogPath`；`FINISHED` 给出状态、退出码、结果路径、计数、区域问题和保留项摘要。报告/CLI 版本仍为 1；Skill 版本与包内程序版本分别维护。

`FINISHED` 的稳定字段：

| 字段 | 含义 |
|---|---|
| `SkillVersion`、`ApplicationVersion` | 分别为 Skill 0.2.3 与实际包内程序版本 |
| `Status`、`ExitCode` | COMPLETE/PARTIAL 配合 0 表示已交付；不能只看退出码 |
| `ConversionMode` | 完成转换时的 Standard / Fast / Ultra |
| `VerificationMode` | STANDARD_VERIFIED / FAST_UNVERIFIED / ULTRA_UNVERIFIED。任务未完成为 NOT_COMPLETED |
| `WordStarted` | 极速为 false。标准或快速的原始报告未提供此字段时为 null |
| `Targets`、`Succeeded` | 待转换总数与成功数；不逐项返回成功公式 |
| `FailedPreserved`、`FailedUnverified`、`RejectedPreserved`、`NotAttempted` | 失败已保留、快速或极速失败未经核验、识别拒绝、未完成计数 |
| `NumbersAdded`、`NumbersUpdated`、`NumberReferencesUpdated`、`ExistingDisplaysFormatted` | 编号、引用与已有行外式排版计数 |
| `PreservedItems` | 只含 FAILED_PRESERVED、FAILED_UNVERIFIED、REJECTED_PRESERVED、NOT_ATTEMPTED 项 |
| `Issues` | 区域或阶段问题，带原报告 `Stage/Reason`；不会混入成功公式 |
| `Warnings` | 独立版面提醒；不增加失败数，不改变 Status |
| `ReportPath`、`TaskDirectory`、`LogPath` | 完整结果和技术日志的位置，按需读取 |

问题投影统一保留 `Id/Region/Status/Reason/Stage/SourceLocation/Page/SourceExcerpt`。`SourceExcerpt` 为单行短片段，连同省略号最多 160 个 UTF-16 字符，不拆代理对；完整 `Source` 只在本机报告中。逐式原报告不提供 `Stage` 时为 null，不从错误文字猜测阶段；`Issues` 另保留 `Key/OriginalPreserved`。`Page` 只采用转换前原文从首页连续计数的页序（含封面），未知为 null；页眉页脚的共用内容由节和类型定位，无唯一页。任务整体失败无已交付报告时返回顶层 `ErrorCode/Reason/Stage`，不虚构逐式结果或计数。

默认模式中，退出码 0 配合 `COMPLETE` / `PARTIAL` 表示已交付核验结果；部分完成仍保留输出。退出码 124、125、126、127 分别对应超时、服务失联、恢复问题和取消。其他失败由 `ErrorCode` 区分，如 EXECUTABLE_NOT_FOUND、INCOMPATIBLE_CONTRACT、BUSY、TRACK_CHANGES_ENABLED、DOCUMENT_PROTECTED、ENGINE_ERROR。修订及限制编辑错误同时提供中文说明。PowerShell 参数绑定错误由宿主返回，尚未进入转换脚本。

只有用户明确接受跳过核验时才使用 `-FastConversion`。它要求 EXE 的能力描述含 `ExperimentalFastConversion: true`，底层参数为 `--experimental-fast`。此时 `COMPLETE/PARTIAL` 只表示已保存输出，不表示内容已核验；`FastConversion` 为 true、`Validation` 为 `SKIPPED_EXPERIMENTAL_FAST`、`WordReopenVerified` 为 false、`PostConversionScans` 为 0，原件字节及资产核验字段为 null。`FailedUnverified` 与逐式 `FAILED_UNVERIFIED` 单列未经核验的失败，不计入 `FailedPreserved`；快速模式的位置坐标也未经保存后核验。默认模式沿用原有核验契约。

极速交付时，`ConversionMode: Ultra`、`VerificationMode: ULTRA_UNVERIFIED`、`Validation: NONE`、`WordStarted: false`、`WordReopenVerified: false`。`FastConversion: false` 表示未选择快速档。用 `ConversionMode` 和 `VerificationMode` 区分三档，不能仅凭 `FastConversion` 为 false 判断已核验。极速的失败式通过 `Issues` 返回短片段及 `BodyBlock`，`PreservedItems` 为空。`BodyBlock` 是正文顶层块从 1 起的编号，计入表格等块，页码保持 null。

完整 `result.json`、`worker.log`、`worker.errors` 保留在本机任务目录，可能含私人路径与公式源码，不应默认回显或随 Skill 上传。`FinalStart/FinalEnd` 是最终 DOCX 的 Word 主故事范围坐标，不能作为普通字符串下标；文档再次编辑后需重扫。实验项当前报告所在区域及结果，不承诺区域内精确坐标。

常规正向报告的 `Warnings` 单列版式提醒，不计入 `Issues` 或 `PARTIAL`。完整报告每条包含 `Id/Region/Code/Source/SourceLocation/Page/Actual/Limit`，FINISHED 将 `Source` 投影为 `SourceExcerpt`，保留其余值，并以 `Status: WARNING`、`Stage: LAYOUT_CHECK`、`Reason: Code` 供 AI 分类。尺寸单位为磅；`FORMULA_TOO_WIDE` 表示正文独占行外式超出单栏/等宽栏段落可用宽度，`EXACT_LINE_SPACING` 表示公式高于固定行距，均提示疑似显示问题；`LAYOUT_CHECK_UNAVAILABLE` 表示读取版式失败（尺寸为 0，不用于判断）。`LayoutWarningsChecked` 表示是否执行这一有限范围的检查，`LayoutWarningScope` 说明检查范围。快速、极速或反向时为 false，`Warnings` 为空。不检查非等宽栏、表格/文本框容器宽高及编号碰撞。

需要读取进度时查看本任务日志中的 `UI_PHASE`、`UI_EXPERIMENTAL_PHASE`、`UI_STAGE_PROGRESS`。重试可能撤回中间计数，以最终报告为准。创建本任务停止标记后保留监管进程完成清理。

## 开发入口边界

EXE 公共后台入口为 `--worker-convert`、`--worker-reverse`。`--worker-scan` 是只读基本正文扫描，不能据此宣称全部实验区域已预检。`--help` 与 `--describe --json` 都不启动 Word。

直接 worker 的位置参数依次是模式、输入、新输出、新报告、字号/反向定界符、停止标记，然后才是选项。使用 Skill 时优先调用包装脚本，它负责 Windows 参数引用、独立路径、取消及报告读取。
