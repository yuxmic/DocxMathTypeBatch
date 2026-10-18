# DocxMathTypeBatch

原生 C++ / Win32 的 Word 公式转换工具。将正文中的 LaTeX 转换为可编辑 MathType，并生成独立 DOCX；支持多文档队列、公式范围、字号、行距及可选行外右编号，保留基本反向导出。

## 运行和开发

运行成品需要 Windows x64、64 位桌面 Word 和完整 MathType。开发构建需要 PowerShell 7、MSVC C++ 工具和 Windows SDK。Python 只用于测试，成品不依赖 Python、PowerShell 或 .NET。

在项目根目录运行：

```powershell
python -m pip install -r requirements-dev.txt
.\build.ps1 -Target All
.\test.ps1 -Suite Static
.\test.ps1 -Suite Core -SkipBuild
.\package.ps1
```

程序生成到 `build/bin/DocxMathTypeBatch.exe`，发布包生成到 `dist/`。打包只包含程序和简要使用说明，不自动运行转换测试，也不把历史测试结果视为当前版本验收。

## 目录约定

| 目录 | 内容 | 提交到 Git |
| --- | --- | --- |
| `src/` | 程序源码和共用头文件 | 是 |
| `assets/` | 图标、资源和生成脚本 | 是 |
| `docs/` | 使用说明、当前开发约定和说明原稿 | 是 |
| `tests/` | 测试脚本、合成样本生成器和原生测试工具 | 是，排除输出 |
| `tests/output/` | 本轮合成 DOCX、日志、截图、PDF 和报告 | 否 |
| `build/` | 编译中间文件、生成头文件、EXE 和构建记录 | 否 |
| `dist/` | 发布包和经过检查的源码导出包 | 否 |
| `.local/` | 私人输入、私人结果及本机维护记录 | 否 |
| `临时文件夹/` | 本轮临时材料及待手动清理的旧资料 | 否 |

整个源码目录可移动到其他位置；构建和公开测试不读取旁边的项目、旧版本目录或私人样本。Word、MathType 和编译工具通过系统安装信息定位。

## 测试

`Static` 检查公开文件、脚本语法和合成样本；`Core` 检查正反向、扫描范围、布局、编号和高级设置；`Safety` 检查取消、故障、超时、安装路径和外链；`UI` 检查界面、队列和诊断；`All` 顺序运行后三组。

```powershell
.\test.ps1 -Suite Safety -SkipBuild
.\test.ps1 -Suite UI -SkipBuild
```

测试输入来自代码生成的通用公式，不包含个人论文。每个专项自行准备所需输入，不要求预先保存的 JSON 报告。性能比较需单独指定项目内的本地基线成品，日常功能测试跳过性能比较。安全测试会操作它核实身份的任务进程，运行期间请勿手动操作这些测试窗口。

## GitHub 上传范围

`upload-files.txt` 是唯一公开文件清单，`.gitignore` 按该清单逐个放行。新增公开文件时同时更新这两个文件；本地目录和文档格式始终排除。

```powershell
.\check-upload.ps1
.\check-upload.ps1 -Export
```

检查会拒绝本机身份、绝对文件路径、私人样本名称、疑似密钥和清单外的已跟踪文件。`-Export` 只导出清单中的源码与资源，并回读 ZIP 校验文件名和哈希。不要把完整工作目录压缩上传；本地日志和截图可能包含文档内容、窗口标题或路径。

使用方法见 [使用说明](docs/使用说明.md)，实现约定见 [轻量化转换大纲](docs/轻量化转换大纲.md)，公式兼容约定见 [MathType兼容LaTeX书写规范](docs/MathType兼容LaTeX书写规范.md)。项目暂不添加开源许可证，授权方式待定。
