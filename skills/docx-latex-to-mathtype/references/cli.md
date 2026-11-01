# Skill v0.2.1 调用约定

要求包内 EXE 支持 `CliContractVersion: 1`、报告 `SchemaVersion: 1`。Skill 包内包含已匹配的 `bin/DocxMathTypeBatch.exe`，更新时替换完整包；不包含文档样本、本机配置或日志。

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
| `-FontSize 12` | 正向指定 1–256 pt；省略时跟随原文 |
| `-RightNumbers` | 正文行外公式补号/更新编号，默认关闭 |
| `-NumberFormat '(#C1-#E1)'` | 新增编号格式；默认 `(#E1)` |
| `-NumberStart`、`-NumberChapter`、`-NumberSection` | 1–999999 的整数，默认 1 |
| `-CustomNumber` | 使用自定义编号格式，配合 RightNumbers |
| `-InlineSpacing`、`-DisplaySpacing` | 默认 follow；single、onehalf、double、exact:20、atleast:20、multiple:1.25 |
| `-DisplayDelimiter bracket/dollar` | 反向行外使用 `\[...\]` 或 `$$...$$`；行内使用 `$...$` |
| `-Tables`、`-Textboxes`、`-Headers`、`-Notes` | 显式启用四类额外位置，默认关闭 |
| `-SkipOmml` | 正向保留正文及已启用额外位置中的 Word 原生公式，只转换 LaTeX；默认不传 |
| `-FastConversion` | 默认关闭；实验性直接保存模式，跳过转换前内容/语法核验、逐式复查及保存后重开重扫/资产核验，可能漏转或错位 |
| `-TaskDirectory` | 可选全新任务目录；默认创建本机用户数据下的独立目录 |
| `-TimeoutSeconds` | 无进展超时，5–3600 秒，默认 120；不是整份任务总时限 |

编号格式支持 `#C1`、`#S1`、`#E1`，末尾可为 `a/A/i/I`，并且仅含一个 `#E`。章、节、起始值作用于目标序列；已有编号保留其原生显示格式。反向不支持字号、右编号和段落行距设置。

额外位置仅转换公式；OMML 默认参与正向转换。用户要求保留 Word 原生公式时传入 `-SkipOmml`，同时作用于正文及启用的额外位置，不计为拒绝或未完成；反向始终保留 OMML。该选项要求 EXE 能力含 `OptionalOmmlConversion: true`，底层参数为 `--skip-omml`，报告 `ConvertOmml` 表示本次有效的正向 OMML 设置。表格/文本框开关需要其父区域也启用。保护、修订和普通域等仍构成边界。LaTeX 文本以 ASCII 为基础，支持顶饰适配的程序可先将 `Ŷ`、`Ŷ` 等写法转换为命令，其余未适配字符保留原文；OMML 使用 Office 原生转换资源。是否成功以连接程序的实际报告为准。

## 任务结果与诊断

脚本只向标准输出发送 JSON Lines，不混入 worker 原始日志，也不启动 GUI。`STARTED` 给出 `Output`、`TaskDirectory`、`StopFile`、`LogPath`；`FINISHED` 给出状态、退出码、结果路径、计数、区域问题和保留项摘要。报告/CLI 版本仍为 1；Skill 版本与包内程序版本分别维护。

`FINISHED` 的稳定字段：

| 字段 | 含义 |
|---|---|
| `SkillVersion`、`ApplicationVersion` | 分别为 Skill 0.2.1 与实际包内程序版本，不能混为同一版本 |
| `Status`、`ExitCode` | COMPLETE/PARTIAL 配合 0 表示已交付；不能只看退出码 |
| `VerificationMode` | STANDARD_VERIFIED / FAST_UNVERIFIED；任务未完成为 NOT_COMPLETED |
| `Targets`、`Succeeded` | 待转换总数与成功数；不逐项返回成功公式 |
| `FailedPreserved`、`FailedUnverified`、`RejectedPreserved`、`NotAttempted` | 失败已保留、快速失败未经核验、识别拒绝、未完成计数 |
| `NumbersAdded`、`NumbersUpdated`、`NumberReferencesUpdated`、`ExistingDisplaysFormatted` | 编号、引用与已有行外式排版计数 |
| `PreservedItems` | 只含 FAILED_PRESERVED、FAILED_UNVERIFIED、REJECTED_PRESERVED、NOT_ATTEMPTED 项 |
| `Issues` | 区域或阶段问题，带原报告 `Stage/Reason`；不会混入成功公式 |
| `Warnings` | 独立版面提醒；不增加失败数，不改变 Status |
| `ReportPath`、`TaskDirectory`、`LogPath` | 完整结果和技术日志的位置，按需读取 |

问题投影统一保留 `Id/Region/Status/Reason/Stage/SourceLocation/Page/SourceExcerpt`。`SourceExcerpt` 为单行短片段，连同省略号最多 160 个 UTF-16 字符，不拆代理对；完整 `Source` 只在本机报告中。逐式原报告不提供 `Stage` 时为 null，不从错误文字猜测阶段；`Issues` 另保留 `Key/OriginalPreserved`。`Page` 只采用转换前原文从首页连续计数的页序（含封面），未知为 null；页眉页脚的共用内容由节和类型定位，无唯一页。任务整体失败无已交付报告时返回顶层 `ErrorCode/Reason/Stage`，不虚构逐式结果或计数。

默认模式中，退出码 0 配合 `COMPLETE` / `PARTIAL` 表示已交付核验结果；部分完成仍保留输出。退出码 124、125、126、127 分别对应超时、服务失联、恢复问题和取消。其他失败由 `ErrorCode` 区分，如 EXECUTABLE_NOT_FOUND、INCOMPATIBLE_CONTRACT、BUSY、TRACK_CHANGES_ENABLED、DOCUMENT_PROTECTED、ENGINE_ERROR。修订及限制编辑错误同时提供中文说明。PowerShell 参数绑定错误由宿主返回，尚未进入转换脚本。

只有用户明确接受跳过核验时才使用 `-FastConversion`。它要求 EXE 的能力描述含 `ExperimentalFastConversion: true`，底层参数为 `--experimental-fast`。此时 `COMPLETE/PARTIAL` 只表示已保存输出，不表示内容已核验；`FastConversion` 为 true、`Validation` 为 `SKIPPED_EXPERIMENTAL_FAST`、`WordReopenVerified` 为 false、`PostConversionScans` 为 0，原件字节及资产核验字段为 null。`FailedUnverified` 与逐式 `FAILED_UNVERIFIED` 单列未经核验的失败，不计入 `FailedPreserved`；快速模式的位置坐标也未经保存后核验。默认模式沿用原有核验契约。

完整 `result.json`、`worker.log`、`worker.errors` 保留在本机任务目录，可能含私人路径与公式源码，不应默认回显或随 Skill 上传。`FinalStart/FinalEnd` 是最终 DOCX 的 Word 主故事范围坐标，不能作为普通字符串下标；文档再次编辑后需重扫。实验项当前报告所在区域及结果，不承诺区域内精确坐标。

常规正向报告的 `Warnings` 单列版式提醒，不计入 `Issues` 或 `PARTIAL`。完整报告每条包含 `Id/Region/Code/Source/SourceLocation/Page/Actual/Limit`，FINISHED 将 `Source` 投影为 `SourceExcerpt`，保留其余值，并以 `Status: WARNING`、`Stage: LAYOUT_CHECK`、`Reason: Code` 供 AI 分类。尺寸单位为磅；`FORMULA_TOO_WIDE` 表示正文独占行外式超出单栏/等宽栏段落可用宽度，`EXACT_LINE_SPACING` 表示公式高于固定行距，均提示疑似显示问题；`LAYOUT_CHECK_UNAVAILABLE` 表示读取版式失败（尺寸为 0，不用于判断）。`LayoutWarningsChecked` 表示是否执行这一有限范围的检查，`LayoutWarningScope` 说明检查范围；快速或反向时为 false，`Warnings` 为空。不检查非等宽栏、表格/文本框容器宽高及编号碰撞。

需要读取进度时查看本任务日志中的 `UI_PHASE`、`UI_EXPERIMENTAL_PHASE`、`UI_STAGE_PROGRESS`。重试可能撤回中间计数，以最终报告为准。创建本任务停止标记后保留监管进程完成清理。

## 开发入口边界

EXE 公共后台入口为 `--worker-convert`、`--worker-reverse`。`--worker-scan` 是只读基本正文扫描，不能据此宣称全部实验区域已预检。`--help` 与 `--describe --json` 都不启动 Word。

直接 worker 的位置参数依次是模式、输入、新输出、新报告、字号/反向定界符、停止标记，然后才是选项。使用 Skill 时优先调用包装脚本，它负责 Windows 参数引用、独立路径、取消及报告读取。
