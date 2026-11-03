# 参与贡献

欢迎通过 [Issues](https://github.com/yuxmic/DocxMathTypeBatch/issues) 反馈问题或提出建议，也可以提交 Pull Request。请先搜索已有问题，说明实际使用场景。

## 反馈转换问题

请提供程序、Windows、Word 和 MathType 版本，以及转换方向、常规或快速模式、正文范围和相关设置。写出操作步骤、预期结果与实际结果；如果只有某个公式出错，优先提供能够复现的最小 LaTeX 或合成样本。

“日志 → 原始日志”可提供定位线索。提交前请去掉文档内容中的私人信息、个人路径和用户名；截图也只展示脱敏内容。不要上传真实论文、未公开材料或整份原始日志来代替最小复现。

## 本机构建

在 Windows x64 上安装 PowerShell 7、MSVC x64 C++ 工具链和 Windows SDK。工具链可通过 Visual Studio 或 Visual Studio Build Tools 的 C++ 桌面开发组件安装。构建和启动：

```powershell
git clone https://github.com/yuxmic/DocxMathTypeBatch.git
cd DocxMathTypeBatch
.\build.ps1 -Target Gui
.\build\bin\DocxMathTypeBatch.exe
```

构建使用原生 C++17 和静态运行库。实际转换另需已安装并注册的 64 位桌面 Word 和完整 MathType；运行程序无需 Python 或 .NET。

## 选择相关检查

Python 仅用于开发测试。先安装测试依赖，再按改动范围选择检查：

```powershell
python -m pip install -r requirements-dev.txt

# 文档、上传清单和脚本：检查公开文件、脚本语法及合成样本，不启动 Office。
.\test.ps1 -Suite Static

# 纯规则或工作进程协议：构建并检查原生纯逻辑，不启动 Office。
.\test.ps1 -Suite Pure
```

涉及转换、编号或版式时，先执行 `.\build.ps1 -Target All`，再选用 `tests/` 下对应的专项，例如 `.\tests\test-native-numbering.ps1 -SkipBuild`。确需检查全部核心转换流程时，可运行 `.\test.ps1 -Suite Core -SkipBuild`。Office 专项会启动本轮任务窗口，请勿在运行中手动操作它们。

测试通过代码生成通用 DOCX，结果保存在被 Git 忽略的 `tests/output/`。不要将本机报告、临时文件或个人样本加入提交。构建、打包和模块职责详见[开发说明](docs/轻量化转换大纲.md)。

## 提交改动

- 一次解决一个明确问题，沿用现有模块和接口；脚本中的必要注释使用中文。
- 新增公开文件时，同时维护 `upload-files.txt` 和 `.gitignore` 的逐文件规则；新增文件类型还需调整 `check-upload.ps1` 的明确放行范围。公开截图只使用合成样本。
- 提交前运行 `.\check-upload.ps1`，确认待提交文件只有本次改动。程序、Skill 内置 EXE 和 ZIP 属于构建或发布产物，不随源码提交。
- Pull Request 说明问题、最终行为、实际运行的检查及尚未验证的范围。涉及界面或版式时附合成样本截图；无需为了文档改动运行整套 Office 回归。

贡献代码采用本项目的 [MIT 许可证](LICENSE)。
