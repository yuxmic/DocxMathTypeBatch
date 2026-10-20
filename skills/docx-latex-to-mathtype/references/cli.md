# Skill v0.1 调用约定

要求 EXE 支持 `CliContractVersion: 1`、报告 `SchemaVersion: 1`。EXE 与 Skill 可以分别升级；连接脚本按接口版本检查兼容性。此包不包含 EXE、文档样本、本机配置或日志。

## 首次连接

正常 GUI 启动后，EXE 把自身完整路径登记到当前用户设置 `Software\DocxMathTypeBatch` 的 `AgentExecutable`。最近一次正常 GUI 启动或显式连接的版本成为已绑定版本；worker、查询与 UI 测试不更新登记。

`connect.ps1` 不带参数时读取并检查绑定；传入 `-ExePath` 时先验证产品身份、查询接口，再保存新绑定。失败不改写旧绑定。路径失效时正常打开新位置的 EXE 一次，或显式重新连接。

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
| `-TaskDirectory` | 可选全新任务目录；默认创建本机用户数据下的独立目录 |
| `-TimeoutSeconds` | 无进展超时，5–3600 秒，默认 120；不是整份任务总时限 |

编号格式支持 `#C1`、`#S1`、`#E1`，末尾可为 `a/A/i/I`，并且仅含一个 `#E`。章、节、起始值作用于目标序列；已有编号保留其原生显示格式。反向不支持字号、右编号和段落行距设置。

额外位置仅转换公式；正文 OMML 默认参与主线且无需单独开关。表格/文本框开关需要其父区域也启用。保护、修订和普通域等仍构成边界。LaTeX 文本以 ASCII 为支持范围，非 ASCII 导入保留原文；OMML 使用 Office 原生转换资源。

## 任务结果与诊断

脚本只向标准输出发送 JSON Lines，不混入 worker 原始日志。`STARTED` 给出 `Output`、`TaskDirectory`、`StopFile`、`LogPath`；`FINISHED` 给出状态、退出码、结果路径、计数、区域问题和保留项摘要。

退出码 0 配合 `COMPLETE` / `PARTIAL` 表示已交付核验结果；部分完成仍保留输出。退出码 124、125、126、127 分别对应超时、服务失联、恢复问题和取消。其他失败由 `ErrorCode` 区分，如 NEEDS_CONNECTION、EXECUTABLE_NOT_FOUND、INCOMPATIBLE_CONTRACT、BUSY、ENGINE_ERROR。PowerShell 参数绑定错误由宿主返回，尚未进入转换脚本。

完整 `result.json`、`worker.log`、`worker.errors` 保留在本机任务目录，可能含私人路径与公式源码，不应默认回显或随 Skill 上传。`FinalStart/FinalEnd` 是最终 DOCX 的 Word 主故事范围坐标，不能作为普通字符串下标；文档再次编辑后需重扫。实验项当前报告所在区域及结果，不承诺区域内精确坐标。

需要读取进度时查看本任务日志中的 `UI_PHASE`、`UI_EXPERIMENTAL_PHASE`、`UI_STAGE_PROGRESS`。重试可能撤回中间计数，以最终报告为准。创建本任务停止标记后保留监管进程完成清理。

## 开发入口边界

EXE 公共后台入口为 `--worker-convert`、`--worker-reverse`。`--worker-scan` 是只读基本正文扫描，不能据此宣称全部实验区域已预检。`--help` 与 `--describe --json` 都不启动 Word。

直接 worker 的位置参数依次是模式、输入、新输出、新报告、字号/反向定界符、停止标记，然后才是选项。使用 Skill 时优先调用包装脚本，它负责 Windows 参数引用、独立路径、取消及报告读取。
