// B 双栏原生界面：同一 EXE 的独立 STA 工作进程执行扫描与转换，UI 不持有 Word 对象。
#include "docx_scan.h"
#include "mathtype_api.h"
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <cmath>
#include <set>
#include <sstream>
#include <thread>
#include <share.h>
#include "layout_options.h"
#include <gdiplus.h>
#include <d2d1.h>
#include <dwrite.h>
// 启用系统新版控件外观；绘制仍使用原生 Windows API。
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

int convertMain(int argc, wchar_t** argv);
int recoverPreferencesMain(const std::wstring& originalName);
ScanResult scanReverseDocument(IDispatch* document);
constexpr UINT WorkerLine = WM_APP + 1, WorkerDone = WM_APP + 2, BeginInput = WM_APP + 3, NextDocument = WM_APP + 4, WorkerLog = WM_APP + 5;
// 辅助文字和禁用文字分别使用统一色值，避免系统灰字与正文混淆。
constexpr COLORREF MutedTextColor=RGB(145,158,169),DisabledTextColor=RGB(178,188,197);
enum Control { Title=100, Forward, Reverse, FileHeading, FilePath, Choose, Existing,
    StageHeading, Detail, StyleHeading, ModeLabel, ModeBox, SizeLabel, SizeBox,
    SizeSpin, StyleNote, OutputHeading, DirectoryLabel, DirectoryBox, ChooseDirectory, NameLabel, NameBox,
    OutputNote, Status, Progress, Start, Stop, OpenOutput, OpenFolder, OpenReport, Environment,
    InlineDelimiterLabel, InlineDelimiterNote, DisplayDelimiterLabel, Bracket, Dollar, FollowSize, AddNumbers,
    QueueList, ClearQueue, AddFolder, OriginalDirectory, CustomDirectory, Advanced, Elapsed, NumberLabel };

enum class DocumentState { Pending, Running, Complete, Partial, Failed, Cancelled };
constexpr const wchar_t* DocumentStates[]={L"待处理",L"转换中",L"完成",L"部分完成",L"失败",L"已取消"};
struct QueuedDocument {
    std::wstring path, output, error;
    ULONGLONG bytes{};
    DocumentState state{DocumentState::Pending};
    std::wstring context, stage;
    std::string trace;
    DWORD exitCode{STILL_ACTIVE};
};

// 保留诊断末尾，防止长批次日志无限占用内存；失败原因通常位于末尾。
constexpr size_t MaximumLogBytes=256*1024;
static void appendWorkerLog(std::string& log,const std::string& line) {
    log+=line;log+="\r\n";
    if(log.size()>MaximumLogBytes){
        const std::string notice="[Earlier log truncated; keeping the last 256 KiB]\r\n";
        size_t start=log.size()-(MaximumLogBytes-notice.size());const auto newline=log.find('\n',start);
        if(newline!=std::string::npos&&newline+1<log.size())start=newline+1;
        log.erase(0,start);log.insert(0,notice);
    }
}

struct FileStamp {
    BY_HANDLE_FILE_INFORMATION value{};
    bool valid{};
};
// 文件选择只读元数据；文档内容在点击开始后由转换进程锁定并识别。
static FileStamp fileStamp(const std::wstring& path) {
    FileStamp result;
    HANDLE file = CreateFileW(path.c_str(), FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                              nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file != INVALID_HANDLE_VALUE) { result.valid = GetFileInformationByHandle(file, &result.value) != 0; CloseHandle(file); }
    return result;
}
static std::wstring windowText(HWND window) {
    const int length = GetWindowTextLengthW(window);
    std::wstring result(static_cast<size_t>(length) + 1, L'\0');
    GetWindowTextW(window, result.data(), length + 1); result.resize(static_cast<size_t>(length)); return result;
}
static std::wstring trim(std::wstring value) {
    const auto first = value.find_first_not_of(L" \t\r\n");
    return first == std::wstring::npos ? L"" : value.substr(first, value.find_last_not_of(L" \t\r\n") - first + 1);
}
static std::wstring parentPath(const std::wstring& path) { return path.substr(0, path.find_last_of(L"\\/")); }
static std::wstring fileName(const std::wstring& path) { return path.substr(path.find_last_of(L"\\/") + 1); }
static std::wstring fromUtf8(const std::string& bytes) {
    if (bytes.empty()) return {};
    require(bytes.size() <= INT_MAX, "UTF-8 worker record bounds");
    const int length = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0);
    require(length > 0, "Valid UTF-8 worker record");
    std::wstring value(static_cast<size_t>(length), L'\0');
    require(MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), static_cast<int>(bytes.size()), value.data(), length) == length,
            "Decode worker record"); return value;
}
static std::wstring unhex(const std::string& encoded) {
    if (encoded == "-") return {};
    require(encoded.size() % 2 == 0 && encoded.size() <= 4 * 1024 * 1024, "Worker string bounds");
    auto digit = [](char c) { require((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F'), "Worker hexadecimal digit"); return c <= '9' ? c - '0' : c - 'A' + 10; };
    std::string bytes;
    for (size_t index = 0; index < encoded.size(); index += 2) bytes += static_cast<char>((digit(encoded[index]) << 4) | digit(encoded[index + 1]));
    return fromUtf8(bytes);
}
static std::wstring reasonText(const std::wstring& reason) {
    if (reason == L"EMPTY") return L"公式为空";
    if (reason == L"UNBALANCED_BRACES") return L"花括号未配对";
    if (reason == L"UNBALANCED_ENVIRONMENTS") return L"LaTeX 环境未配对，请检查 begin/end 的名称和顺序";
    if (reason == L"UNMATCHED_OPEN") return L"缺少结束定界符或遇到结构边界";
    if (reason == L"UNMATCHED_CLOSE") return L"缺少起始定界符";
    if (reason == L"MIXED_DELIMITERS") return L"定界符混用";
    if (reason == L"AMBIGUOUS_DOLLARS") return L"连续美元符含义不明确";
    if (reason == L"UNSAFE_MISSING_DELIMITER") return L"公式缺少起始反斜杠，且不满足安全补全条件，请补齐定界符";
    if (reason == L"UNSUPPORTED_LINE_BREAK") return L"公式内含换行，原文保留";
    if (reason == L"UNSUPPORTED_NON_ASCII_TEX") return L"公式内容含非 ASCII 字符，原文保留";
    if (reason == L"STOP_REQUESTED") return L"已安全停止，尚未转换";
    if (reason == L"UNSUPPORTED_NON_ASCII_EXPORT") return L"导出内容含当前不支持的非 ASCII 字符，原对象保留";
    if (reason == L"UNSUPPORTED_OR_INVALID_NATIVE_EXPORT") return L"导出内容为空、含换行或无法校验，原对象保留";
    if (reason == L"UNREADABLE_MATHTYPE_NATIVE") return L"未能读取原生公式，原对象保留";
    if (reason == L"MATH_TYPE_IN_EXCLUDED_BOUNDARY") return L"公式位于修订、内容控件或普通域内，原对象保留";
    if (reason == L"LINKED_MATHTYPE_PRESERVED") return L"链接型 MathType 对象暂不导出";
    if (reason == L"UNSUPPORTED_MATHTYPE_CLASS") return L"当前不支持此 MathType 对象类型，原对象保留";
    if (reason == L"NO_TARGET_FORMULAS") return L"未找到需要转换的公式，请检查定界符及公式模式";
    if (reason == L"TASK_TIMEOUT") return L"Word 或 MathType 接口长时间未返回，任务已自动终止";
    if (reason == L"GUI_COMMIT_TIMEOUT") return L"主界面未及时确认输出，任务未完成交付";
    if (reason == L"MATHTYPE_SERVICE_DISCONNECTED") return L"MathType 服务已失联，转换任务已终止";
    if (reason == L"PREFERENCE_RECOVERY_REQUIRED") return L"MathType 公式样式尚未确认恢复，请恢复 MathType 服务后重试";
    if (reason == L"MATHTYPE_TASK_BUSY") return L"另一个公式转换任务正在运行，请等待该任务结束";
    if (reason == L"TASK_COPY_CLEANUP_FAILED") return L"任务工作副本未能清理，请查看下方日志中的文件路径和系统错误码";
    if (reason == L"Word registration" || reason == L"Registered Word executable path") return L"未检测到桌面版 Word，或 Word 自动化组件未正确注册，请安装或修复 Word";
    if (reason == L"Start Word" || reason == L"Create Word atomically inside task job" || reason == L"Task Word native object starts within deadline") return L"无法启动任务所需的 Word 实例，请确认 Word 可以正常启动且已完成激活";
    if (reason == L"Load x64 MathPage.wll" || reason == L"Missing MathType export") return L"无法加载 MathType 64 位接口，安装组件可能缺失或版本不兼容，请修复完整 MathType 安装";
    if (reason == L"MathType initialization") return L"MathType 接口初始化失败，请确认 MathType 能正常启动且已完成授权";
    if (reason == L"Unencrypted DOCX ZIP required") return L"文档不是有效的未加密 DOCX 文件，请确认文件格式并移除打开密码";
    if (reason.find(L"(Win32=32)") != std::wstring::npos || reason.find(L"(Win32=33)") != std::wstring::npos) return L"文档被其他程序占用，请保存并关闭文档后重试";
    if (reason.find(L"(Win32=2)") != std::wstring::npos || reason.find(L"(Win32=3)") != std::wstring::npos) return L"找不到文档或所在文件夹，请确认文件未被移动或删除";
    if (reason.find(L"(Win32=5)") != std::wstring::npos) return L"无法访问文档，当前用户没有所需的文件权限";
    if (reason == L"Task copy must be writable") return L"任务文档为只读状态，无法写入转换结果";
    if (reason == L"Disable revision tracking in source before batch conversion") return L"文档开启了修订模式，请先关闭修订并保存文档后重试";
    if (reason == L"Protected document cannot be converted") return L"文档启用了编辑保护，请先解除保护后重试";
    if (reason == L"Existing output directory required") return L"输出文件夹不存在或无法访问，请重新选择保存位置";
    if (reason == L"Create independent task copy" || reason == L"Reserve independent task copy" || reason == L"Fill reserved task copy") return L"无法在输出文件夹创建文档工作副本，请检查写入权限和磁盘剩余空间";
    if (reason == L"Publish verified file without overwrite") return L"无法交付输出文件，请检查目标文件是否被占用、文件夹写入权限及磁盘剩余空间";
    if (reason == L"Output already exists" || reason == L"No free output file name") return L"输出文件名已被占用，请修改文件名后缀或选择其他文件夹";
    if (reason == L"Write TeX into new object") return L"MathType 未能导入该公式的 LaTeX，请检查命令和环境是否受 MathType 支持";
    if (reason == L"New equation contains native content") return L"MathType 未生成有效的公式内容，该公式转换未成功";
    if (reason == L"New equation native style matches task EQP") return L"生成公式的字号、样式或行内/行外类型与设置不一致";
    if (reason == L"Apply isolated equation preferences") return L"无法向新公式应用指定的字号和样式";
    if (reason == L"New equation dimensions" || reason == L"Saved OLE dimensions") return L"生成的公式宽度或高度无效";
    if (reason == L"Query native bytes" || reason == L"Read native bytes" || reason == L"Query TeX length" || reason == L"Read TeX length") return L"无法读取 MathType 公式数据，公式对象可能损坏或接口未正常返回";
    if (reason == L"Native export delimiter matches true equation type") return L"MathType 导出公式的定界符与实际行内/行外类型不一致";
    if (reason == L"Close converted equation" || reason == L"Close failed task equation") return L"MathType 公式对象未能正常关闭，当前处理未完成";
    if (reason == L"Recovery must restore exact source text, character/paragraph formatting and object count" || reason == L"Reverse recovery restores source, formatting and OLE" || reason == L"Undo failed conversion" || reason == L"Undo failed reverse replacement") return L"转换失败后无法确认原式已完整恢复，已终止整份文档的转换";
    if (reason == L"Reopened target native content and style identical at its original position" || reason == L"No unexpected residual items") return L"保存后核验发现公式内容、位置或残留原文不符合预期，未交付输出文档";
    if (reason == L"IDispatch::Invoke" || reason == L"GetIDsOfNames") return L"Word 或 MathType 自动化接口调用失败，请查看下方后台日志中的接口名称及详细错误";
    // 已有中文安装错误直接保留；未知内部错误也必须有中文说明，原始信息由调用者附上。
    if (std::any_of(reason.begin(),reason.end(),[](wchar_t c){return c>=0x4e00&&c<=0x9fff;})) return reason;
    return L"程序在当前阶段处理失败，请查看下方原始错误和后台日志";
}
static std::wstring failureText(const std::wstring& reason) {
    const auto chinese=reasonText(reason);
    return chinese==reason?chinese:chinese+L"\r\n原始错误："+reason;
}
static void saveNew(const std::wstring& path, const std::string& bytes) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(file != INVALID_HANDLE_VALUE, "Create new GUI task file"); DWORD written{};
    const bool saved = bytes.size() <= MAXDWORD && WriteFile(file, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) && written == bytes.size();
    const bool flushed = FlushFileBuffers(file) != 0, closed = CloseHandle(file) != 0;
    if (!saved || !flushed || !closed) { DeleteFileW(path.c_str()); throw std::runtime_error("Save complete GUI task file"); }
}
static bool legalName(const std::wstring& name) {
    if (name.size() <= 5 || name.size() > 180 || name.find_first_of(L"<>:\"/\\|?*\r\n\t") != std::wstring::npos ||
        name.back() == L'.' || name.back() == L' ' || _wcsicmp(name.c_str() + name.size() - 5, L".docx") != 0) return false;
    auto stem = name.substr(0, name.find('.')); std::transform(stem.begin(), stem.end(), stem.begin(), towupper);
    if (stem == L"CON" || stem == L"PRN" || stem == L"AUX" || stem == L"NUL") return false;
    if (stem.size() == 4 && (stem.substr(0, 3) == L"COM" || stem.substr(0, 3) == L"LPT") && stem[3] >= L'1' && stem[3] <= L'9') return false;
    return std::none_of(name.begin(), name.end(), [](wchar_t c) { return c < 32; });
}
static std::wstring freshOutput(const std::wstring& directory, const std::wstring& name, const std::wstring& input) {
    require(legalName(name), "Valid new DOCX file name required");
    const DWORD attributes = GetFileAttributesW(directory.c_str());
    require(attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY), "Existing output directory required");
    const auto stem = name.substr(0, name.size() - 5);
    for (unsigned int index = 1; index < 10000; ++index) {
        const auto path = directory + L"\\" + stem + (index == 1 ? L"" : L" (" + std::to_wstring(index) + L")") + L".docx";
        if (_wcsicmp(path.c_str(), input.c_str()) == 0) continue;
        if (GetFileAttributesW(path.c_str()) == INVALID_FILE_ATTRIBUTES) return path;
    }
    throw std::runtime_error("No free output file name");
}
// CreateProcess 的参数按 Windows 反斜杠/引号规则编码；不经过任何命令 shell。
static std::wstring processQuote(const std::wstring& argument) {
    std::wstring result = L"\""; size_t slashes{};
    for (const wchar_t character : argument) {
        if (character == L'\\') { ++slashes; continue; }
        result.append(slashes * (character == L'"' ? 2 : 1), L'\\'); slashes = 0;
        if (character == L'"') result += L'\\'; result += character;
    }
    result.append(slashes * 2, L'\\'); return result + L'"';
}
// 环境检测只读注册表，避免 CLSIDFromProgID 对未知类产生注册信息。
static bool registeredClass(const wchar_t* progId) {
    wchar_t value[64]{};DWORD bytes=sizeof(value),type{};CLSID cls{};
    const auto key=std::wstring(progId)+L"\\CLSID";
    return RegGetValueW(HKEY_CLASSES_ROOT,key.c_str(),nullptr,RRF_RT_REG_SZ,&type,value,&bytes)==ERROR_SUCCESS&&SUCCEEDED(CLSIDFromString(value,&cls));
}
struct App {
    ComPtr<ID2D1Factory> drawingFactory;ComPtr<ID2D1DCRenderTarget> textTarget;ComPtr<IDWriteFactory> textFactory;
    ComPtr<IDWriteTextFormat> bodyText,headingText,titleText;
    HWND window{}; HINSTANCE instance{}; UINT dpi{96}; HFONT body{}, heading{}, title{}; HBRUSH white{}, background{};HICON documentIcon{};
    std::wstring executable, input, task, eqp, stopFile, output, resultReport, testDirectory, testMode;
    std::wstring install, mathInstallError;
    std::vector<std::wstring> failures;
    bool reverseMode{},progressMarquee{},testScanProgress{},testSavedProgress{},testReopenMarquee{},cancelling{},acceptingResult{};
    FileStamp stamp; std::thread worker; bool busy{}, converting{}, scanned{}, closePending{}, workerFailed{}, summary{},
        wordDetected{}, mathDetected{}, wordAvailable{}, mathAvailable{}, testControls{}, testLayout{}, testParameterPaint{};
    long existing{}; size_t targets{}, inlines{}, completed{}, succeeded{}, failed{}, unattempted{}, rejected{};
    int testStep{}, exitCode{}; std::wstring error, outcome, lastStage, currentFormula;
    size_t testBackgroundPaints{},testRefreshes{};
    std::vector<QueuedDocument> documents;
    std::vector<size_t> batchOrder;
    std::vector<std::wstring> batchTestInputs;
    size_t batchPosition{};
    bool batchRunning{}, batchStopping{}, batchTest{};
    ULONGLONG batchStarted{}, batchElapsed{};
    EquationSpacing inlineSpacing,displaySpacing;
    EquationNumbering numbering;
    bool chapterNumber{},sectionNumber{};
    int chapterStyle{},sectionStyle{},equationStyle{},numberEnclosure{};
    std::wstring numberSeparator{L"."};
    HIMAGELIST queueImages{};
    HWND hotControl{},pressedControl{};int hotHalf{};
    std::wstring interfaceFont{L"Microsoft YaHei UI"};
    ~App() {
        if (worker.joinable()) worker.join(); cleanupTask();
        for (auto font : {body, heading, title}) if (font) DeleteObject(font);
        if (white) DeleteObject(white); if (background) DeleteObject(background);
        if(documentIcon)DestroyIcon(documentIcon);
        if(queueImages)ImageList_Destroy(queueImages);
    }
    HWND control(int id) const { return GetDlgItem(window, id); }
    void set(int id, const std::wstring& value) const { if(windowText(control(id))!=value)SetWindowTextW(control(id), value.c_str()); }
    void enable(int id,bool enabled) const {if((IsWindowEnabled(control(id))!=FALSE)!=enabled)EnableWindow(control(id),enabled);}
    void show(int id,bool visible) const {if(((GetWindowLongPtrW(control(id),GWL_STYLE)&WS_VISIBLE)!=0)!=visible)ShowWindow(control(id),visible?SW_SHOW:SW_HIDE);}
    void animateProgress(bool active) {
        if(progressMarquee==active)return;
        auto bar=control(Progress);if(!active)SendMessageW(bar,PBM_SETMARQUEE,FALSE,0);
        const auto style=GetWindowLongPtrW(bar,GWL_STYLE);SetWindowLongPtrW(bar,GWL_STYLE,active?style|PBS_MARQUEE:style&~PBS_MARQUEE);
        if(active)SendMessageW(bar,PBM_SETMARQUEE,TRUE,40);progressMarquee=active;
    }
    void countProgress(size_t current,size_t total) {
        animateProgress(false);SendMessageW(control(Progress),PBM_SETRANGE32,0,static_cast<LPARAM>(total));SendMessageW(control(Progress),PBM_SETPOS,current,0);
        if(!testDirectory.empty())require((GetWindowLongPtrW(control(Progress),GWL_STYLE)&PBS_MARQUEE)==0&&
            SendMessageW(control(Progress),PBM_GETPOS,0,0)==static_cast<LRESULT>(current),"Native progress bar shows reported count");
    }
    int px(int value) const { return MulDiv(value, static_cast<int>(dpi), 96); }
    bool cleanupTask() {
        // 临时目录由本实例唯一创建，仅清理登记的任务文件，不递归删除或枚举用户目录。
        if (task.empty()) return true;
        bool clean=true;
        for (const auto& path : {eqp, stopFile, resultReport,stopFile+L".commit"}) if (!path.empty() && !DeleteFileW(path.c_str()) && GetLastError()!=ERROR_FILE_NOT_FOUND) clean=false;
        if(!RemoveDirectoryW(task.c_str())&&GetLastError()!=ERROR_PATH_NOT_FOUND)clean=false;
        if(clean){task.clear(); eqp.clear(); stopFile.clear(); resultReport.clear();}return clean;
    }
    void newTask() {
        require(cleanupTask(),"Previous GUI task files are removed"); wchar_t temporary[32768]{}; require(GetTempPathW(32768, temporary) > 0, "Windows task temporary directory");
        GUID id{}; check(CoCreateGuid(&id), "Create GUI task identity"); wchar_t guid[40]{}; StringFromGUID2(id, guid, 40);
        const auto path = std::wstring(temporary) + L"DocxMathTypeBatch-" + guid;
        require(CreateDirectoryW(path.c_str(), nullptr) != 0, "Create unique GUI task directory"); task = path;
        resultReport = task + L"\\result.json"; eqp = task + L"\\style.eqp"; stopFile = task + L"\\stop.flag";
    }
    HWND add(int id, const wchar_t* cls, const wchar_t* label, DWORD style=0, DWORD extended=0) {
        if(!wcscmp(cls,L"COMBOBOX"))style|=CBS_OWNERDRAWFIXED|CBS_HASSTRINGS;
        HWND child = CreateWindowExW(extended, cls, label, WS_CHILD | WS_VISIBLE | style, 0, 0, 1, 1, window,
                                     reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), instance, nullptr);
        require(child != nullptr, "Create native GUI control"); SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(body), FALSE); return child;
    }
    void createFonts() {
        // 明确使用界面字体，清除系统消息字体的字族约束，避免映射到宋体。
        LOGFONTW font{}; font.lfCharSet=DEFAULT_CHARSET; font.lfQuality=CLEARTYPE_QUALITY; font.lfWeight=FW_NORMAL;
        wcscpy_s(font.lfFaceName,interfaceFont.c_str()); font.lfHeight=-px(12);
        HFONT nextBody=CreateFontIndirectW(&font); font.lfHeight=-px(14); font.lfWeight=FW_SEMIBOLD;
        HFONT nextHeading=CreateFontIndirectW(&font); font.lfHeight=-px(20); HFONT nextTitle=CreateFontIndirectW(&font);
        require(nextBody&&nextHeading&&nextTitle,"Create DPI interface fonts");
        for(HWND child=GetWindow(window,GW_CHILD);child;child=GetWindow(child,GW_HWNDNEXT))SendMessageW(child,WM_SETFONT,reinterpret_cast<WPARAM>(nextBody),TRUE);
        for(int id:{FileHeading,StyleHeading,OutputHeading})SendMessageW(control(id),WM_SETFONT,reinterpret_cast<WPARAM>(nextHeading),TRUE);
        SendMessageW(control(Title),WM_SETFONT,reinterpret_cast<WPARAM>(nextTitle),TRUE);
        for(auto old:{body,heading,title})if(old)DeleteObject(old);body=nextBody;heading=nextHeading;title=nextTitle;
    }
    void create() {
        // DirectWrite 避免本机 GDI 将中文界面字形回退为宋体；仍仅使用系统绘图库。
        check(D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED,drawingFactory.GetAddressOf()),"Create native text drawing factory");
        check(DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED,__uuidof(IDWriteFactory),reinterpret_cast<IUnknown**>(textFactory.GetAddressOf())),"Create DirectWrite font factory");
        auto properties=D2D1::RenderTargetProperties(D2D1_RENDER_TARGET_TYPE_SOFTWARE,D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM,D2D1_ALPHA_MODE_IGNORE));
        check(drawingFactory->CreateDCRenderTarget(&properties,&textTarget),"Create native text target");
        ComPtr<IDWriteFontCollection> interfaceFonts;check(textFactory->GetSystemFontCollection(&interfaceFonts),"Read interface font collection");
        UINT32 familyIndex{};BOOL notoAvailable{};check(interfaceFonts->FindFamilyName(L"Noto Sans SC",&familyIndex,&notoAvailable),"Find Chinese sans-serif interface family");
        const wchar_t* interfaceFamily=notoAvailable?L"Noto Sans SC":L"Microsoft YaHei UI";
        // 原生 GDI 控件使用稳定可识别的字体，避免 Noto 可变字体被回退为宋体。
        interfaceFont=L"Microsoft YaHei UI";
        auto format=[&](float size,DWRITE_FONT_WEIGHT weight,ComPtr<IDWriteTextFormat>& result){check(textFactory->CreateTextFormat(interfaceFamily,nullptr,weight,DWRITE_FONT_STYLE_NORMAL,DWRITE_FONT_STRETCH_NORMAL,size,L"zh-CN",&result),"Create native text format");};
        format(12,DWRITE_FONT_WEIGHT_NORMAL,bodyText);format(14,DWRITE_FONT_WEIGHT_SEMI_BOLD,headingText);format(20,DWRITE_FONT_WEIGHT_SEMI_BOLD,titleText);
        for(auto text:{bodyText.Get(),headingText.Get(),titleText.Get()}){ComPtr<IDWriteInlineObject> ellipsis;
            check(textFactory->CreateEllipsisTrimmingSign(text,&ellipsis),"Create interface text ellipsis");
            DWRITE_TRIMMING trimming{DWRITE_TRIMMING_GRANULARITY_CHARACTER,0,0};check(text->SetTrimming(&trimming,ellipsis.Get()),"Trim long interface labels");}
        dpi=GetDpiForWindow(window);white=CreateSolidBrush(RGB(255,255,255));background=CreateSolidBrush(RGB(246,248,250));createFonts();
        add(Title,L"STATIC",L"文档公式转换");
        add(Forward,L"BUTTON",L"LaTeX → MathType",BS_OWNERDRAW|WS_TABSTOP|WS_GROUP);
        add(Reverse,L"BUTTON",L"MathType → LaTeX",BS_OWNERDRAW|WS_TABSTOP);
        add(FileHeading,L"STATIC",L"待处理文档");
        add(FilePath,L"STATIC",L"拖入 .docx 文件，可多选",SS_LEFT|SS_ENDELLIPSIS);
        add(Choose,L"BUTTON",L"添加文件",BS_OWNERDRAW|WS_TABSTOP);
        add(AddFolder,L"BUTTON",L"添加文件夹",BS_OWNERDRAW|WS_TABSTOP);
        add(ClearQueue,L"BUTTON",L"清空列表",BS_OWNERDRAW|WS_TABSTOP);
        add(Existing,L"STATIC",L"0 个文件",SS_ENDELLIPSIS);
        add(QueueList,WC_LISTVIEWW,L"",LVS_REPORT|LVS_SHOWSELALWAYS|LVS_SHAREIMAGELISTS|WS_TABSTOP);
        ListView_SetExtendedListViewStyle(control(QueueList),LVS_EX_FULLROWSELECT|LVS_EX_DOUBLEBUFFER|LVS_EX_LABELTIP|LVS_EX_INFOTIP);
        ListView_SetBkColor(control(QueueList),RGB(255,255,255));ListView_SetTextBkColor(control(QueueList),RGB(255,255,255));ListView_SetTextColor(control(QueueList),RGB(48,65,79));
        const wchar_t* columns[]={L"文档名称",L"大小",L"状态",L"移除",L"日志"};
        for(int column=0;column<5;++column){LVCOLUMNW item{};item.mask=LVCF_TEXT|LVCF_WIDTH|LVCF_FMT;item.pszText=const_cast<wchar_t*>(columns[column]);item.cx=px(80);item.fmt=column>=3?LVCFMT_CENTER:LVCFMT_LEFT;ListView_InsertColumn(control(QueueList),column,&item);}
        add(StageHeading,L"STATIC",L"待转换 0   完成 0   失败 0");
        add(StyleHeading,L"STATIC",L"转换设置");add(ModeLabel,L"STATIC",L"公式模式");
        add(ModeBox,L"COMBOBOX",L"",CBS_DROPDOWNLIST|WS_VSCROLL|WS_TABSTOP);
        for(const auto* mode:{L"全部公式",L"仅行内公式",L"仅行外公式"})SendMessageW(control(ModeBox),CB_ADDSTRING,0,reinterpret_cast<LPARAM>(mode));
        SendMessageW(control(ModeBox),CB_SETCURSEL,0,0);add(SizeLabel,L"STATIC",L"公式字号");
        add(SizeBox,L"EDIT",L"12",ES_AUTOHSCROLL|WS_TABSTOP);SendMessageW(control(SizeBox),EM_SETLIMITTEXT,16,0);
        add(SizeSpin,UPDOWN_CLASSW,L"",UDS_ARROWKEYS|UDS_NOTHOUSANDS);
        SendMessageW(control(SizeSpin),UDM_SETBUDDY,reinterpret_cast<WPARAM>(control(SizeBox)),0);SendMessageW(control(SizeSpin),UDM_SETRANGE32,1,256);
        add(FollowSize,L"COMBOBOX",L"",CBS_DROPDOWNLIST|WS_TABSTOP);
        for(const auto* label:{L"跟随原文",L"指定字号"})SendMessageW(control(FollowSize),CB_ADDSTRING,0,reinterpret_cast<LPARAM>(label));
        SendMessageW(control(FollowSize),CB_SETCURSEL,0,0);
        add(NumberLabel,L"STATIC",L"行外右编号");add(AddNumbers,L"COMBOBOX",L"",CBS_DROPDOWNLIST|WS_TABSTOP);
        for(const auto* label:{L"不添加",L"添加右编号"})SendMessageW(control(AddNumbers),CB_ADDSTRING,0,reinterpret_cast<LPARAM>(label));
        SendMessageW(control(AddNumbers),CB_SETCURSEL,0,0);
        add(Advanced,L"BUTTON",L"高级设置",BS_OWNERDRAW|WS_TABSTOP);
        add(InlineDelimiterLabel,L"STATIC",L"行内定界符");add(InlineDelimiterNote,L"STATIC",L"$...$（固定）");add(DisplayDelimiterLabel,L"STATIC",L"行外定界符");
        add(Bracket,L"BUTTON",L"\\[...\\]",BS_AUTORADIOBUTTON|WS_TABSTOP|WS_GROUP);add(Dollar,L"BUTTON",L"$$...$$",BS_AUTORADIOBUTTON);
        SendMessageW(control(Bracket),BM_SETCHECK,BST_CHECKED,0);
        add(OutputHeading,L"STATIC",L"输出设置");add(DirectoryLabel,L"STATIC",L"保存位置");
        add(OriginalDirectory,L"BUTTON",L"原文件夹",BS_AUTORADIOBUTTON|WS_GROUP|WS_TABSTOP);
        add(CustomDirectory,L"BUTTON",L"指定目录",BS_AUTORADIOBUTTON);SendMessageW(control(OriginalDirectory),BM_SETCHECK,BST_CHECKED,0);
        add(DirectoryBox,L"EDIT",L"原文件所在文件夹",ES_READONLY|ES_AUTOHSCROLL);
        add(ChooseDirectory,L"BUTTON",L"浏览",BS_OWNERDRAW|WS_TABSTOP);add(NameLabel,L"STATIC",L"文件名后缀");
        add(NameBox,L"EDIT",L"_MathType",ES_AUTOHSCROLL|WS_TABSTOP);SendMessageW(control(NameBox),EM_SETLIMITTEXT,64,0);
        add(OutputNote,L"STATIC",L"保留原件，重名时自动添加序号");
        add(Status,L"STATIC",L"就绪");add(Elapsed,L"STATIC",L"已用时 00:00");add(Progress,PROGRESS_CLASSW,L"",PBS_SMOOTH);
        add(Start,L"BUTTON",L"开始转换",BS_OWNERDRAW|WS_TABSTOP);add(Stop,L"BUTTON",L"安全停止",BS_OWNERDRAW|WS_TABSTOP);
        createFonts();
        for(HWND child=GetWindow(window,GW_CHILD);child;child=GetWindow(child,GW_HWNDNEXT)){
            wchar_t kind[32]{};GetClassNameW(child,kind,32);if(!_wcsicmp(kind,L"Static"))SetWindowSubclass(child,paintLabel,1,reinterpret_cast<DWORD_PTR>(this));
        }
        for(int id:{SizeBox,DirectoryBox,NameBox}){SetWindowSubclass(control(id),paintInput,1,reinterpret_cast<DWORD_PTR>(this));SetWindowPos(control(id),nullptr,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_FRAMECHANGED);}
        SendMessageW(control(ModeBox),CB_SETITEMHEIGHT,static_cast<WPARAM>(-1),px(30));
        for(int id:{Forward,Reverse,Choose,AddFolder,ClearQueue,Advanced,ChooseDirectory,Start,Stop,ModeBox,FollowSize,AddNumbers,SizeSpin,Progress,Bracket,Dollar,QueueList})
            SetWindowSubclass(control(id),paintControl,1,reinterpret_cast<DWORD_PTR>(this));
        SHFILEINFOW fileInfo{};if(SHGetFileInfoW(L".docx",FILE_ATTRIBUTE_NORMAL,&fileInfo,sizeof(fileInfo),SHGFI_ICON|SHGFI_LARGEICON|SHGFI_USEFILEATTRIBUTES))documentIcon=fileInfo.hIcon;
        queueImages=ImageList_Create(px(24),px(32),ILC_COLOR32|ILC_MASK,1,1);require(queueImages!=nullptr,"Create document queue icons");
        if(documentIcon)ImageList_AddIcon(queueImages,documentIcon);ListView_SetImageList(control(QueueList),queueImages,LVSIL_SMALL);
        SetWindowTheme(control(QueueList),L"Explorer",nullptr);
        SetWindowSubclass(ListView_GetHeader(control(QueueList)),paintQueueHeader,1,reinterpret_cast<DWORD_PTR>(this));
        for(int id:{OriginalDirectory,CustomDirectory})SetWindowSubclass(control(id),paintControl,1,reinterpret_cast<DWORD_PTR>(this));
        wordDetected=registeredClass(L"Word.Application");
        install=findMathTypeInstall(mathInstallError);mathDetected=!install.empty();
        if(testDirectory.empty())loadSettings();if(testMode==L"reverse-dollar")CheckRadioButton(window,Bracket,Dollar,Dollar);
        updateDirectionControls();DragAcceptFiles(window,TRUE);DragAcceptFiles(control(QueueList),TRUE);layout();refresh();
    }
    bool validDelimiter() const {
        return (SendMessageW(control(Bracket),BM_GETCHECK,0,0)==BST_CHECKED) != (SendMessageW(control(Dollar),BM_GETCHECK,0,0)==BST_CHECKED);
    }
    void updateDirectionControls() {
        for (int id : {ModeLabel,ModeBox,SizeLabel,SizeBox,SizeSpin,FollowSize,AddNumbers,NumberLabel}) show(id,!reverseMode);
        for (int id : {InlineDelimiterLabel,InlineDelimiterNote,DisplayDelimiterLabel,Bracket,Dollar}) show(id,reverseMode);
        set(StyleHeading,reverseMode?L"导出格式":L"转换设置");
        show(Advanced,!reverseMode);
    }
    void setDirection(bool reverse) {
        if(batchRunning||reverseMode==reverse)return;reverseMode=reverse;updateDirectionControls();
        set(NameBox,reverse?L"_LaTeX":L"_MathType");refresh();
        if(!testDirectory.empty()&&!input.empty()&&testStep==0&&testMode!=L"preview"&&!batchTest)testAdvance();
        InvalidateRect(control(Forward),nullptr,TRUE);InvalidateRect(control(Reverse),nullptr,TRUE);
    }
    void loadSettings() {
        HKEY key{}; if (RegOpenKeyExW(HKEY_CURRENT_USER,L"Software\\DocxMathTypeBatch",0,KEY_QUERY_VALUE,&key) != ERROR_SUCCESS) return;
        for (const auto& entry : {std::make_pair(L"PointSize",SizeBox)}) {
            wchar_t value[128]{}; DWORD type{}, bytes=sizeof(value);
            if (RegQueryValueExW(key,entry.first,nullptr,&type,reinterpret_cast<BYTE*>(value),&bytes)==ERROR_SUCCESS && type==REG_SZ && bytes<sizeof(value)) set(entry.second,value);
        } RegCloseKey(key);
    }
    void saveSettings() {
        if (!testDirectory.empty()) return;
        std::wstring why; double points{}; if (!style(points,why)) return;
        HKEY key{}; if (RegCreateKeyExW(HKEY_CURRENT_USER,L"Software\\DocxMathTypeBatch",0,nullptr,0,KEY_SET_VALUE,nullptr,&key,nullptr)!=ERROR_SUCCESS) return;
        for (const auto& entry : {std::make_pair(L"PointSize",SizeBox)}) {
            const auto value=trim(windowText(control(entry.second)));
            RegSetValueExW(key,entry.first,0,REG_SZ,reinterpret_cast<const BYTE*>(value.c_str()),static_cast<DWORD>((value.size()+1)*sizeof(wchar_t)));
        } RegCloseKey(key);
    }
    void layout() {
        RECT client{};GetClientRect(window,&client);const int width=MulDiv(client.right,96,static_cast<int>(dpi)),height=MulDiv(client.bottom,96,static_cast<int>(dpi));
        const int middle=width*56/100,left=32,right=middle+24,lw=middle-56,rw=width-right-32;
        auto move=[&](int id,int x,int y,int w,int h){MoveWindow(control(id),px(x),px(y),px(std::max(w,1)),px(std::max(h,1)),TRUE);};
        move(Title,18,16,300,32);move(Forward,width-314,17,148,30);move(Reverse,width-166,17,148,30);
        move(FileHeading,left,88,144,24);move(Existing,left+146,88,120,24);move(ClearQueue,middle-96,86,72,28);
        move(FilePath,left+14,136,lw-204,30);move(Choose,middle-194,134,76,32);move(AddFolder,middle-112,134,88,32);
        move(QueueList,left,196,lw,height-324);move(StageHeading,left,height-116,lw,24);
        // 总宽度不变，预留纵向滚动条，其余空间交给文档名称。
        ListView_SetColumnWidth(control(QueueList),0,px(lw)-px(52)-px(76)-px(36)-px(44)-GetSystemMetricsForDpi(SM_CXVSCROLL,dpi));
        ListView_SetColumnWidth(control(QueueList),1,px(52));ListView_SetColumnWidth(control(QueueList),2,px(76));
        ListView_SetColumnWidth(control(QueueList),3,px(36));ListView_SetColumnWidth(control(QueueList),4,px(44));
        move(StyleHeading,right,88,rw-90,24);move(Advanced,width-116,86,84,28);
        move(ModeLabel,right,130,74,34);move(ModeBox,right+82,130,rw-82,240);
        const bool fixed=SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==1;
        move(SizeLabel,right,178,74,34);move(FollowSize,right+82,178,rw-82-(fixed?100:0),200);move(SizeBox,width-120,178,64,34);move(SizeSpin,width-56,178,24,34);
        move(NumberLabel,right,226,74,34);move(AddNumbers,right+82,226,rw-82,200);
        move(InlineDelimiterLabel,right,132,88,30);move(InlineDelimiterNote,right+90,132,rw-90,30);
        move(DisplayDelimiterLabel,right,180,88,30);move(Bracket,right+90,178,88,32);move(Dollar,right+182,178,80,32);
        move(OutputHeading,right,286,rw,24);move(DirectoryLabel,right,316,74,28);
        move(OriginalDirectory,right+82,312,90,30);move(CustomDirectory,right+180,312,90,30);
        move(DirectoryBox,right+82,350,rw-142,34);move(ChooseDirectory,width-84,350,52,34);
        move(NameLabel,right,398,78,32);move(NameBox,right+82,396,rw-82,34);move(OutputNote,right,444,rw,22);
        move(Status,18,height-60,width-410,24);move(Elapsed,width-394,height-60,138,24);
        move(Stop,width-250,height-62,90,34);move(Start,width-154,height-62,136,34);move(Progress,18,height-20,width-36,5);
        InvalidateRect(window,nullptr,TRUE);for(int id:{ModeBox,FollowSize,AddNumbers}){SendMessageW(control(id),CB_SETITEMHEIGHT,0,px(28));SendMessageW(control(id),CB_SETITEMHEIGHT,static_cast<WPARAM>(-1),px(28));}
    }
    bool style(double& points,std::wstring& why) const {
        if(SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==0){points=0;return true;}
        const auto size=trim(windowText(control(SizeBox))); wchar_t* end{};
        points=wcstod(size.c_str(),&end);
        if (size.empty() || !end || *end || !std::isfinite(points) || points<1 || points>256 || std::abs(points*100-std::round(points*100))>0.000001) {why=L"主字号应为 1–256 pt，最多两位小数。";return false;}
        return true;
    }
    std::wstring environmentProblem() const {
        std::wstring problem=mathDetected?L"":mathInstallError;
        if(!wordDetected){if(!problem.empty())problem+=L"\r\n";problem+=L"未检测到桌面版 Word：未安装或自动化组件未注册，请安装或修复 Word。";}
        return problem;
    }
    void refresh() {
        if(!control(Start))return;
        if(!testDirectory.empty())++testRefreshes;
        std::wstring why;double points{};const bool validStyle=reverseMode?validDelimiter():style(points,why);
        const bool original=SendMessageW(control(OriginalDirectory),BM_GETCHECK,0,0)==BST_CHECKED;
        const auto attributes=GetFileAttributesW(windowText(control(DirectoryBox)).c_str());
        const auto suffix=windowText(control(NameBox));
        const bool validOutput=legalName(L"document"+suffix+L".docx")&&(original||(attributes!=INVALID_FILE_ATTRIBUTES&&(attributes&FILE_ATTRIBUTE_DIRECTORY)));
        enable(Start,!batchRunning&&!documents.empty()&&wordDetected&&mathDetected&&validStyle&&validOutput);
        for(int id:{Choose,AddFolder,ClearQueue,ModeBox,FollowSize,NameBox,Forward,Reverse,Bracket,Dollar,OriginalDirectory,CustomDirectory,Advanced})enable(id,!batchRunning);
        enable(ChooseDirectory,!batchRunning&&!original);enable(DirectoryBox,!batchRunning&&!original);
        const bool fixed=SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==1;
        show(SizeBox,!reverseMode&&fixed);show(SizeSpin,!reverseMode&&fixed);enable(SizeBox,!batchRunning&&fixed);enable(SizeSpin,!batchRunning&&fixed);
        enable(AddNumbers,!batchRunning&&SendMessageW(control(ModeBox),CB_GETCURSEL,0,0)!=1);
        show(Stop,batchRunning);enable(Stop,batchRunning&&!closePending&&!batchStopping);
        show(Progress,batchRunning);
        size_t pending{};for(const auto& item:documents)pending+=item.state==DocumentState::Pending;
        set(Start,(pending||!batchStarted?L"开始转换":L"重新转换")+(documents.empty()?std::wstring():L"（"+std::to_wstring(pending?pending:documents.size())+L"）"));
        if(!batchRunning&&!summary){
            if(!mathDetected||!wordDetected)set(Status,environmentProblem());else if(documents.empty())set(Status,L"就绪 · 请添加文档");
            else if(!validStyle)set(Status,reverseMode?L"请选择行外定界符":why);else if(!validOutput)set(Status,L"请设置有效的输出位置和名称");else set(Status,L"准备就绪");
        }
    }
    void redrawCounts() {
        size_t pending{},done{},errors{};for(const auto& item:documents){pending+=item.state==DocumentState::Pending;done+=item.state==DocumentState::Complete||item.state==DocumentState::Partial;errors+=item.state==DocumentState::Failed;}
        set(Existing,std::to_wstring(documents.size())+L" 个文件");set(StageHeading,L"待转换 "+std::to_wstring(pending)+L"    完成 "+std::to_wstring(done)+L"    失败 "+std::to_wstring(errors));
    }
    void launch(const std::vector<std::wstring>& arguments) {
        const bool hadCounts=scanned;
        if (worker.joinable()) worker.join(); busy=true; workerFailed=false; summary=false; scanned=false;cancelling=acceptingResult=false; error.clear(); outcome.clear();lastStage=L"启动后台任务";currentFormula.clear();failures.clear(); targets=inlines=completed=succeeded=failed=unattempted=rejected=0;
        require(SetEnvironmentVariableW(L"DOCX_MATH_GUI_COMMIT",(stopFile+L".commit").c_str())!=0,"Set current GUI publication acknowledgement");
        if(hadCounts)redrawCounts();
        animateProgress(true);refresh();
        const auto target=window; const auto program=executable;const bool trace=!testDirectory.empty();
        worker=std::thread([target,program,arguments,trace]() {
            HANDLE read=nullptr,write=nullptr,nul=INVALID_HANDLE_VALUE; PROCESS_INFORMATION process{}; DWORD code=1;
            std::string log;
            auto post=[&](const std::string& line){appendWorkerLog(log,line);if(!trace&&line.rfind("UI_",0)!=0&&line.rfind("CONVERT_PROGRESS",0)!=0&&line.rfind("OWNED_WORD_PID=",0)!=0)return;auto value=std::make_unique<std::string>(line); if(PostMessageW(target,WorkerLine,0,reinterpret_cast<LPARAM>(value.get()))) value.release();};
            try {
                SECURITY_ATTRIBUTES security{sizeof(security),nullptr,TRUE};
                require(CreatePipe(&read,&write,&security,0)!=0&&SetHandleInformation(read,HANDLE_FLAG_INHERIT,0)!=0,"Create worker output pipe");
                nul=CreateFileW(L"NUL",GENERIC_READ,FILE_SHARE_READ|FILE_SHARE_WRITE,&security,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,nullptr);
                require(nul!=INVALID_HANDLE_VALUE,"Close worker standard input");
                STARTUPINFOW start{sizeof(start)}; start.dwFlags=STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW; start.wShowWindow=SW_HIDE;
                start.hStdInput=nul; start.hStdOutput=write; start.hStdError=write;
                std::wstring command=processQuote(program); for(const auto& argument:arguments) command+=L" "+processQuote(argument);
                // STARTUPINFOEX 将可继承句柄限制为管道与 NUL，避免并发实例泄露其他句柄。
                SIZE_T bytes{}; InitializeProcThreadAttributeList(nullptr,1,0,&bytes);
                std::vector<unsigned char> storage(bytes); auto attributes=reinterpret_cast<LPPROC_THREAD_ATTRIBUTE_LIST>(storage.data());
                require(InitializeProcThreadAttributeList(attributes,1,0,&bytes)!=0,"Worker handle inheritance list");
                HANDLE handles[]={write,nul};
                const bool updated=UpdateProcThreadAttribute(attributes,0,PROC_THREAD_ATTRIBUTE_HANDLE_LIST,handles,sizeof(handles),nullptr,nullptr)!=0;
                STARTUPINFOEXW extended{}; extended.StartupInfo=start; extended.StartupInfo.cb=sizeof(extended); extended.lpAttributeList=attributes;
                const bool started=updated&&CreateProcessW(program.c_str(),command.data(),nullptr,nullptr,TRUE,CREATE_NO_WINDOW | EXTENDED_STARTUPINFO_PRESENT,
                                                           nullptr,nullptr,&extended.StartupInfo,&process)!=0;
                DeleteProcThreadAttributeList(attributes); require(started,"Start isolated native worker");
                CloseHandle(process.hThread); process.hThread=nullptr; CloseHandle(write); write=nullptr; CloseHandle(nul); nul=INVALID_HANDLE_VALUE;
                std::string pending; char buffer[4096]; DWORD count{}; bool oversized=false;
                while(ReadFile(read,buffer,sizeof(buffer),&count,nullptr)&&count) {
                    pending.append(buffer,count); size_t newline{};
                    while((newline=pending.find('\n'))!=std::string::npos) {auto line=pending.substr(0,newline);while(!line.empty()&&line.back()=='\r')line.pop_back();post(line);pending.erase(0,newline+1);}
                    if(pending.size()>4*1024*1024){pending.clear();oversized=true;}
                }
                if(!pending.empty())post(pending); WaitForSingleObject(process.hProcess,INFINITE);
                require(GetExitCodeProcess(process.hProcess,&code)!=0,"Read worker exit code");
                if(oversized){post("UI_ERROR "+wireHex("Worker output exceeded record bounds"));code=1;}
            } catch(const std::exception& exception) {post("UI_ERROR "+wireHex(exception.what()));}
            if(read)CloseHandle(read);if(write)CloseHandle(write);if(nul!=INVALID_HANDLE_VALUE)CloseHandle(nul);
            if(process.hThread)CloseHandle(process.hThread);if(process.hProcess)CloseHandle(process.hProcess);
            auto savedLog=std::make_unique<std::string>(std::move(log));if(PostMessageW(target,WorkerLog,0,reinterpret_cast<LPARAM>(savedLog.get())))savedLog.release();
            PostMessageW(target,WorkerDone,code,0);
        });
    }
    void updateRow(size_t index) {
        const auto& document=documents.at(index);
        auto value=fileName(document.path);ListView_SetItemText(control(QueueList),static_cast<int>(index),0,value.data());
        wchar_t bytes[48]{};swprintf_s(bytes,document.bytes>=1024*1024?L"%.1f MB":L"%.0f KB",document.bytes>=1024*1024?document.bytes/1048576.0:std::max(1.0,document.bytes/1024.0));
        ListView_SetItemText(control(QueueList),static_cast<int>(index),1,bytes);
        ListView_SetItemText(control(QueueList),static_cast<int>(index),2,const_cast<wchar_t*>(DocumentStates[static_cast<int>(document.state)]));
        ListView_SetItemText(control(QueueList),static_cast<int>(index),3,const_cast<wchar_t*>(L"×"));
        ListView_SetItemText(control(QueueList),static_cast<int>(index),4,const_cast<wchar_t*>(L"查看"));
    }
    void addInputs(const std::vector<std::wstring>& paths) {
        require(!batchRunning,"Wait for current batch before changing inputs");size_t skipped{};
        for(const auto& path:paths){
            const auto resolved=absolutePath(path.c_str());const auto name=fileName(resolved);
            if(name.size()<=5||name.rfind(L"~$",0)==0||_wcsicmp(name.c_str()+name.size()-5,L".docx")!=0||
                std::any_of(documents.begin(),documents.end(),[&](const QueuedDocument& item){return _wcsicmp(item.path.c_str(),resolved.c_str())==0;})){++skipped;continue;}
            const auto identity=fileStamp(resolved);if(!identity.valid||(identity.value.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY)){++skipped;continue;}
            ULARGE_INTEGER size{};size.HighPart=identity.value.nFileSizeHigh;size.LowPart=identity.value.nFileSizeLow;
            documents.push_back({resolved,L"",L"",size.QuadPart,DocumentState::Pending});
            auto label=fileName(resolved);LVITEMW row{};row.mask=LVIF_TEXT|LVIF_IMAGE;row.iItem=static_cast<int>(documents.size()-1);row.pszText=label.data();row.iImage=0;
            ListView_InsertItem(control(QueueList),&row);updateRow(documents.size()-1);
        }
        if(!documents.empty()){input=documents.front().path;if(SendMessageW(control(OriginalDirectory),BM_GETCHECK,0,0)==BST_CHECKED)set(DirectoryBox,parentPath(input));}
        summary=false;redrawCounts();refresh();if(skipped)set(Status,L"已跳过 "+std::to_wstring(skipped)+L" 个重复或非 DOCX 文件");
    }
    void chooseInput(const std::wstring& path) {
        addInputs({path});
        if(!testDirectory.empty()&&testMode!=L"preview"&&!batchTest)testAdvance();
    }
    void removeDocument(int index) {
        if(batchRunning||index<0||static_cast<size_t>(index)>=documents.size())return;
        documents.erase(documents.begin()+index);ListView_DeleteItem(control(QueueList),index);
        input=documents.empty()?L"":documents.front().path;summary=false;redrawCounts();refresh();
    }
    void updateElapsed() {
        const auto milliseconds=batchRunning?GetTickCount64()-batchStarted:batchElapsed;
        // 用系统单调时钟的差值计时，界面漏掉刷新也不会使计时变慢。
        wchar_t duration[64]{};swprintf_s(duration,L"已用时 %02llu:%02llu",milliseconds/60000,(milliseconds/1000)%60);set(Elapsed,duration);
    }
    void saveNewBatchReady() {
        require(!busy&&task.empty()&&!scanned,"Batch selection does not launch Office");
        saveNew(testDirectory+L"\\gui-ready.json","{\"NoPreScan\":true,\"Documents\":"+std::to_string(documents.size())+"}\n");SetTimer(window,2,1500,nullptr);
    }
    void finishBatch() {
        batchElapsed=GetTickCount64()-batchStarted;batchRunning=false;KillTimer(window,4);updateElapsed();animateProgress(false);
        summary=true;redrawCounts();refresh();
        const bool hasFailures=std::any_of(documents.begin(),documents.end(),[](const QueuedDocument& item){return item.state==DocumentState::Failed;});
        set(Status,batchStopping?L"已停止":hasFailures?L"批次完成 · 双击失败文档查看详情":L"批次完成");
        if(batchTest){
            std::string json="{\"Status\":\"PASSED\",\"ElapsedMilliseconds\":"+std::to_string(batchElapsed)+",\"Stopped\":"+(batchStopping?std::string("true"):std::string("false"))+",\"Items\":[";
            for(size_t i=0;i<documents.size();++i){const auto& item=documents[i];json+=(i?",":"")+std::string("{\"Input\":")+quote(item.path)+",\"State\":"+std::to_string(static_cast<int>(item.state))+",\"Output\":"+quote(item.output)+",\"Error\":"+quote(item.error)+",\"ExitCode\":"+std::to_string(item.exitCode)+",\"Stage\":"+quote(item.stage)+",\"Details\":"+quote(detailsText(static_cast<int>(i)))+"}";}
            saveNew(testDirectory+L"\\batch-check.json",json+"]}\n");if(testMode!=L"batch-diagnostics")SetTimer(window,3,2000,nullptr);
        }
        if(closePending)DestroyWindow(window);
    }
    void nextDocument() {
        if(!batchRunning)return;
        if(batchStopping||batchPosition>=batchOrder.size()){finishBatch();return;}
        auto& document=documents.at(batchOrder.at(batchPosition));input=document.path;
        document.state=DocumentState::Running;document.trace.clear();document.exitCode=STILL_ACTIVE;lastStage=L"准备任务";document.stage.clear();
        SYSTEMTIME started{};GetLocalTime(&started);wchar_t time[40]{};swprintf_s(time,L"%04d-%02d-%02d %02d:%02d:%02d",started.wYear,started.wMonth,started.wDay,started.wHour,started.wMinute,started.wSecond);
        document.context=L"程序："+executable+L"\r\n开始时间："+time+L"\r\n方向："+(reverseMode?L"MathType → LaTeX":L"LaTeX → MathType")+L"\r\n";
        if(!reverseMode)document.context+=L"公式模式："+windowText(control(ModeBox))+L"\r\n字号："+windowText(control(FollowSize))+(SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==1?L" "+windowText(control(SizeBox))+L" pt":L"")+L"\r\n右编号："+windowText(control(AddNumbers))+L"\r\n行内行距："+spacingArgument(inlineSpacing)+L"；行外行距："+spacingArgument(displaySpacing)+L"\r\n";
        updateRow(batchOrder[batchPosition]);ListView_EnsureVisible(control(QueueList),static_cast<int>(batchOrder[batchPosition]),FALSE);
        try{
            newTask();std::wstring styleArgument;
            if(!reverseMode){std::wstring why;double points{};require(style(points,why),"Valid point size required");styleArgument=points==0?L"paragraph":L"size="+std::to_wstring(points);}
            else{require(validDelimiter(),"One reverse display delimiter required");styleArgument=SendMessageW(control(Dollar),BM_GETCHECK,0,0)==BST_CHECKED?L"dollar":L"bracket";}
            const auto directory=SendMessageW(control(OriginalDirectory),BM_GETCHECK,0,0)==BST_CHECKED?parentPath(input):windowText(control(DirectoryBox));
            const auto name=fileName(input);output=freshOutput(directory,name.substr(0,name.size()-5)+windowText(control(NameBox))+L".docx",input);
            document.context+=L"输出候选："+output+L"\r\n";
            if(!reverseMode&&SendMessageW(control(AddNumbers),CB_GETCURSEL,0,0)==1)document.context+=L"编号格式："+numbering.format+L"；起始值："+std::to_wstring(numbering.start)+L"；章："+std::to_wstring(numbering.chapter)+L"；节："+std::to_wstring(numbering.section)+L"\r\n";
            if(reverseMode)document.context+=L"行外定界符："+std::wstring(SendMessageW(control(Dollar),BM_GETCHECK,0,0)==BST_CHECKED?L"$$...$$":L"\\[...\\]")+L"\r\n";
            converting=true;std::vector<std::wstring> arguments{reverseMode?L"--worker-reverse":L"--worker-convert",input,output,resultReport,styleArgument,stopFile,install};
            if(!reverseMode){const auto mode=SendMessageW(control(ModeBox),CB_GETCURSEL,0,0);arguments.push_back(mode==1?L"--scope=inline":mode==2?L"--scope=display":L"--scope=all");if(mode!=1&&SendMessageW(control(AddNumbers),CB_GETCURSEL,0,0)==1){
                    arguments.push_back(L"--numbers");if(numbering.custom)arguments.push_back(L"--number-custom");arguments.push_back(L"--number-format="+numbering.format);
                    arguments.push_back(L"--number-start="+std::to_wstring(numbering.start));arguments.push_back(L"--number-chapter="+std::to_wstring(numbering.chapter));arguments.push_back(L"--number-section="+std::to_wstring(numbering.section));
                }
                if(inlineSpacing.mode)arguments.push_back(L"--inline-spacing="+spacingArgument(inlineSpacing));
                if(displaySpacing.mode)arguments.push_back(L"--display-spacing="+spacingArgument(displaySpacing));}
            if(batchTest){FILE* trace=_wfsopen((testDirectory+L"\\gui-worker.log").c_str(),L"ab",_SH_DENYNO);require(trace!=nullptr,"Open batch worker trace");fprintf(trace,"UI_BATCH_BEGIN %zu\n",batchPosition);fclose(trace);}
            launch(arguments);
        }catch(const std::exception& exception){
            document.state=DocumentState::Failed;document.error=failureText(fromUtf8(exception.what()));document.stage=lastStage;document.exitCode=1;busy=false;updateRow(batchOrder[batchPosition]);
            if(!cleanupTask()){document.error+=L"；临时文件未清理："+task;batchStopping=true;}
            ++batchPosition;PostMessageW(window,NextDocument,0,0);
        }
    }
    void startConversion() {
        if(!IsWindowEnabled(control(Start)))return;
        batchOrder.clear();for(size_t i=0;i<documents.size();++i)if(documents[i].state==DocumentState::Pending)batchOrder.push_back(i);
        if(batchOrder.empty())for(size_t i=0;i<documents.size();++i){documents[i].state=DocumentState::Pending;documents[i].output.clear();documents[i].error.clear();documents[i].context.clear();documents[i].trace.clear();documents[i].stage.clear();documents[i].exitCode=STILL_ACTIVE;updateRow(i);batchOrder.push_back(i);}
        batchPosition=0;batchRunning=true;batchStopping=false;batchStarted=GetTickCount64();batchElapsed=0;SetTimer(window,4,1000,nullptr);updateElapsed();nextDocument();
    }
    void stop() {
        if(!batchRunning||batchStopping)return;batchStopping=true;
        if(!busy||acceptingResult){refresh();return;}cancelling=true;
        if(GetFileAttributesW(stopFile.c_str())==INVALID_FILE_ATTRIBUTES)saveNew(stopFile,"");
        set(Status,L"正在停止并清理当前文档…");EnableWindow(control(Stop),FALSE);
    }
    void line(const std::string& value) {
        // 仅开发自检记录实际工作协议，便于按所属 PID/启动时间注入外部服务故障。
        if(!testDirectory.empty()) {
            // 允许外部验收脚本同时读取，避免安全 fopen 的共享限制制造伪 GUI 故障。
            FILE* trace=_wfsopen((testDirectory+L"\\gui-worker.log").c_str(),L"ab",_SH_DENYNO);require(trace!=nullptr,"Open test worker trace");
            const bool written=fwrite(value.data(),1,value.size(),trace)==value.size()&&fputc('\n',trace)!=EOF;
            const bool closed=fclose(trace)==0;require(written&&closed,"Write complete test worker trace");
        }
        std::istringstream stream(value);std::string kind;stream>>kind;
        if(kind=="UI_CANCELLED"){outcome=L"CANCELLED";summary=false;error.clear();workerFailed=false;return;}
        if(kind=="UI_COMMIT_READY"){
            if(!cancelling){saveNew(stopFile+L".commit","");acceptingResult=true;show(Stop,false);}return;
        }
        if(cancelling&&kind!="UI_ERROR")return;
        if(kind=="UI_ERROR"){std::string encoded;require(static_cast<bool>(stream>>encoded),"Worker error record");
            const auto message=unhex(encoded);if(cancelling&&message!=L"TASK_COPY_CLEANUP_FAILED")return;
            if(!error.empty())error+=L"\r\n";
            if(!currentFormula.empty())error+=currentFormula+L"\r\n";
            error+=failureText(message);
            workerFailed=true;return;}
        if(kind=="UI_ITEM_BEGIN"||kind=="UI_ITEM_REJECTED") {
            size_t index{};long paragraph{},start{},end{};std::string source,reason;
            require(static_cast<bool>(stream>>index>>paragraph>>start>>end>>source)&&index>0,"Formula diagnostic context");
            auto excerpt=unhex(source);for(auto& c:excerpt)if(c==L'\r'||c==L'\n'||c==L'\v'||c==L'\t')c=L' ';
            const auto context=L"公式 "+std::to_wstring(index)+L"（原文第 "+std::to_wstring(paragraph)+L" 段，Word 位置 "+std::to_wstring(start)+L"–"+std::to_wstring(end)+L"）\r\n原式片段："+excerpt;
            if(kind=="UI_ITEM_BEGIN")currentFormula=context;
            else{require(static_cast<bool>(stream>>reason),"Rejected formula reason");failures.push_back(context+L"\r\n"+failureText(unhex(reason))+L"\r\n该项未转换，原文保留。");}
            return;
        }
        if(kind=="UI_ITEM_END"){currentFormula.clear();return;}
        if(value.rfind("OWNED_WORD_PID=",0)==0)wordAvailable=true;
        if(kind=="UI_SCAN_COUNTS") {
            require(static_cast<bool>(stream>>existing>>targets>>inlines>>rejected)&&inlines<=targets,"Scan counters");scanned=true;
            if(!testDirectory.empty())saveNew(testDirectory+(batchTest?L"\\scan-"+std::to_wstring(batchPosition)+L".json":L"\\gui-scan.json"),"{\"Ready\":"+std::to_string(targets)+",\"Rejected\":"+std::to_string(rejected)+",\"ExistingMathType\":"+std::to_string(existing)+"}\n");
            redrawCounts();return;
        }
        if(kind=="UI_PHASE") {
            std::string phase;stream>>phase;
            currentFormula.clear();
            const wchar_t* label=phase=="PREPARING"?L"正在检查输入文档…":phase=="OPENING"?L"正在加载文档…":phase=="SCANNING"?L"正在识别文档公式…":
                phase=="SCANNING_OBJECTS"?L"正在识别已有 MathType…":phase=="INITIALIZING"?L"正在准备公式样式…":phase=="CONVERTING"?L"正在转换公式…":
                phase=="VERIFYING"?L"正在核对公式、残留原文和文档结构…":phase=="SAVING"?L"正在保存副本…":phase=="REOPENING"?L"正在重新打开并核验新公式…":
                phase=="RESCANNING"?L"正在检查保存后的正文…":phase=="RESCANNING_OBJECTS"?L"正在检查保存后的公式对象…":phase=="CLOSING"?L"正在关闭检查文档…":
                phase=="PACKAGE_VERIFY"?L"正在检查输出文件…":phase=="PACKAGE_COMPARE"?L"正在核对原文件及公式数据…":L"正在完成输出…";
            lastStage=fromUtf8(phase)+L" · "+label;set(Status,std::to_wstring(batchPosition+1)+L"/"+std::to_wstring(batchOrder.size())+L" · "+label);
            if(phase=="CONVERTING"){if(targets)countProgress(0,targets);else animateProgress(true);}else animateProgress(true);
            if(phase=="REOPENING")testReopenMarquee=progressMarquee&&(GetWindowLongPtrW(control(Progress),GWL_STYLE)&PBS_MARQUEE)!=0;
            return;
        }
        if(kind=="UI_STAGE_PROGRESS") {
            std::string stage;size_t current{},total{};require(static_cast<bool>(stream>>stage>>current>>total)&&total>0&&current<=total&&total<=INT_MAX,"Stage progress record");
            const wchar_t *label{},*unit=L"项";bool scanning=false,saved=false;
            if(stage=="INPUT_PARTS")label=L"正在检查输入文件";
            else if(stage=="SAVED_PARTS"){label=L"正在检查输出文件";saved=true;}
            else if(stage=="SCAN_TEXT"||stage=="RESCAN_TEXT"){saved=stage=="RESCAN_TEXT";scanning=!saved;label=saved?L"正在检查保存后的正文":L"正在识别公式";unit=L"段";}
            else if(stage=="SCAN_OBJECTS"||stage=="RESCAN_OBJECTS"){saved=stage=="RESCAN_OBJECTS";scanning=!saved;label=saved?L"正在检查保存后的公式对象":L"正在识别已有公式";unit=L"个对象";}
            else if(stage=="SCAN_RANGES"||stage=="RESCAN_RANGES"){saved=stage=="RESCAN_RANGES";scanning=!saved;label=saved?L"正在核对保留原文":L"正在核对识别范围";}
            else if(stage=="BIND_OBJECTS")label=L"正在核对公式位置";
            else if(stage=="VERIFY_POSITIONS"){label=L"正在核对保存后的公式位置";saved=true;}
            else if(stage=="VERIFY_ITEMS"){label=L"正在核验转换结果";saved=true;}
            else if(stage=="NUMBER_EQUATIONS")label=L"正在添加公式编号";
            require(label!=nullptr,"Known progress stage");lastStage=fromUtf8(stage)+L" · "+label+L" "+std::to_wstring(current)+L" / "+std::to_wstring(total);countProgress(current,total);
            set(Status,std::wstring(label)+L" "+std::to_wstring(current)+L" / "+std::to_wstring(total)+L" "+unit);
            if(scanning&&current)testScanProgress=true;if(saved&&current)testSavedProgress=true;return;
        }
        if(kind=="CONVERT_PROGRESS") {
            size_t total{};require(sscanf_s(value.c_str(),"CONVERT_PROGRESS %zu/%zu",&completed,&total)==2&&completed<=total&&total<=INT_MAX,"Conversion progress record");
            countProgress(completed,total);
            set(Status,L"正在转换 "+std::to_wstring(completed)+L" / "+std::to_wstring(total)+L"  ·  成功 "+std::to_wstring(succeeded)+L"，保留 "+std::to_wstring(failed));
            if((testMode==L"stop"||testMode==L"reverse-stop"||(batchTest&&testMode==L"batch-stop"&&batchPosition==1))&&completed==1)stop();return;
        }
        if(kind=="UI_ITEM_RESULT") {
            size_t index{};std::string state,reason;require(static_cast<bool>(stream>>index>>state>>reason)&&index>0,"Conversion item record");
            succeeded+=state=="SUCCEEDED";failed+=state=="FAILED_PRESERVED";
            if(state=="FAILED_PRESERVED")failures.push_back((currentFormula.empty()?L"公式 "+std::to_wstring(index):currentFormula)+L"\r\n"+failureText(unhex(reason))+L"\r\n该式转换失败，原文保留。");
            currentFormula.clear();
            return;
        }
        if(kind=="UI_CONVERT_SUMMARY") {
            std::string state;size_t total{};require(static_cast<bool>(stream>>state>>total>>succeeded>>failed>>unattempted>>rejected)&&total==succeeded+failed+unattempted,"Final conversion counters");
            require(state=="COMPLETE"||state=="PARTIAL","Known conversion outcome");outcome=fromUtf8(state);summary=true;return;
        }
    }
    void done(DWORD code) {
        if(worker.joinable())worker.join();busy=false;animateProgress(false);
        const bool cancelled=code==TaskCancelExit&&outcome==L"CANCELLED";
        workerFailed=!cancelled&&(workerFailed||code!=0||!summary||GetFileAttributesW(output.c_str())==INVALID_FILE_ATTRIBUTES||GetFileAttributesW(resultReport.c_str())==INVALID_FILE_ATTRIBUTES);
        auto& document=documents.at(batchOrder.at(batchPosition));
        document.state=cancelled?DocumentState::Cancelled:workerFailed?DocumentState::Failed:outcome==L"PARTIAL"?DocumentState::Partial:DocumentState::Complete;
        document.error=error;document.stage=lastStage;document.exitCode=code;
        if(!workerFailed&&!cancelled){
            mathAvailable=true;document.output=output;
            if(!testDirectory.empty())require(CopyFileW(resultReport.c_str(),(testDirectory+(batchTest?L"\\report-"+std::to_wstring(batchPosition)+L".json":L"\\report.json")).c_str(),TRUE)!=0,"Keep test report evidence");
        }else if(workerFailed&&document.error.empty())document.error=(currentFormula.empty()?L"":currentFormula+L"\r\n")+(code!=0?L"后台进程未正常完成，原文件保留。":!summary?L"未收到转换结果摘要，原文件保留。":GetFileAttributesW(output.c_str())==INVALID_FILE_ATTRIBUTES?L"未找到已核验的输出文档，原文件保留。":L"未找到转换核验报告，原文件保留。");
        // 整份文档失败时也保留之前逐式发现的问题，后续清理错误不能覆盖原始原因。
        for(const auto& reason:failures){if(!document.error.empty())document.error+=L"\r\n\r\n";document.error+=reason;}
        if(rejected){if(!document.error.empty())document.error+=L"\r\n";document.error+=std::to_wstring(rejected)+L" 项不满足识别条件，原文保留。";}
        if(cancelled&&document.error.empty())document.error=L"用户已停止当前批次，当前文档未交付转换结果。";
        const auto taskDirectory=task;
        const bool clean=cleanupTask();
        if(!clean){document.state=DocumentState::Failed;document.error+=L"\r\n临时文件未清理："+task;workerFailed=true;batchStopping=true;}
        updateRow(batchOrder[batchPosition]);redrawCounts();
        if(batchTest)saveNew(testDirectory+L"\\item-"+std::to_wstring(batchPosition)+L".json","{\"Code\":"+std::to_string(code)+",\"TaskDirectory\":"+quote(taskDirectory)+",\"Cleaned\":"+(clean?std::string("true"):std::string("false"))+"}\n");
        if(cancelled){batchStopping=true;if(!testDirectory.empty()&&!batchTest&&!closePending){
            require(clean&&GetFileAttributesW(output.c_str())==INVALID_FILE_ATTRIBUTES,"Cancelled task files are removed");
            saveNew(testDirectory+L"\\gui-check.json","{\"Status\":\"PASSED\",\"Outcome\":\"CANCELLED\",\"Input\":"+quote(input)+",\"Output\":"+quote(output)+",\"NoOutput\":true,\"NoReport\":true,\"ControlValidation\":true,\"SizeArrows\":true,\"ParameterRepaint\":true,\"KeyboardTabNavigation\":true,\"Layout\":true,\"NoPreScan\":true}\n");testStep=2;SetTimer(window,3,1500,nullptr);
        }}
        if(!testDirectory.empty()&&!batchTest&&workerFailed){exitCode=1;DestroyWindow(window);return;}
        ++batchPosition;
        if(batchStopping||batchPosition>=batchOrder.size())finishBatch();else PostMessageW(window,NextDocument,0,0);
        if(!testDirectory.empty()&&!batchTest&&!cancelled&&!workerFailed&&!closePending&&testMode!=L"preview")testAdvance();
    }
    void selectFile() {
        std::vector<wchar_t> paths(65536);OPENFILENAMEW dialog{sizeof(dialog)};dialog.hwndOwner=window;dialog.lpstrFilter=L"Word 文档 (*.docx)\0*.docx\0\0";
        dialog.lpstrFile=paths.data();dialog.nMaxFile=static_cast<DWORD>(paths.size());dialog.Flags=OFN_FILEMUSTEXIST|OFN_PATHMUSTEXIST|OFN_NOCHANGEDIR|OFN_EXPLORER|OFN_ALLOWMULTISELECT;
        if(GetOpenFileNameW(&dialog)){
            std::vector<std::wstring> selected;const std::wstring first=paths.data();auto next=paths.data()+first.size()+1;
            if(!*next)selected.push_back(first);else for(;*next;next+=wcslen(next)+1)selected.push_back(first+L"\\"+next);
            addInputs(selected);
        }
    }
    void selectDirectory(bool addFolder=false) {
        ComPtr<IFileOpenDialog> dialog;check(CoCreateInstance(CLSID_FileOpenDialog,nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&dialog)),"Open output folder picker");
        DWORD options{};check(dialog->GetOptions(&options),"Folder dialog options");check(dialog->SetOptions(options|FOS_PICKFOLDERS|FOS_FORCEFILESYSTEM),"Folder selection mode");
        dialog->SetTitle(addFolder?L"选择文档所在文件夹":L"选择输出文件夹");const auto shown=dialog->Show(window);if(shown==HRESULT_FROM_WIN32(ERROR_CANCELLED))return;check(shown,"Show folder picker");
        ComPtr<IShellItem> selected;check(dialog->GetResult(&selected),"Selected output folder");PWSTR path{};check(selected->GetDisplayName(SIGDN_FILESYSPATH,&path),"Output folder path");
        const std::wstring directory=path;CoTaskMemFree(path);
        if(addFolder){
            std::vector<std::wstring> paths;WIN32_FIND_DATAW entry{};const auto search=FindFirstFileW((directory+L"\\*.docx").c_str(),&entry);
            if(search!=INVALID_HANDLE_VALUE){do{if(!(entry.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY))paths.push_back(directory+L"\\"+entry.cFileName);}while(FindNextFileW(search,&entry));FindClose(search);}
            std::sort(paths.begin(),paths.end(),[](const auto& a,const auto& b){return _wcsicmp(a.c_str(),b.c_str())<0;});addInputs(paths);
        }else{set(DirectoryBox,directory);CheckRadioButton(window,OriginalDirectory,CustomDirectory,CustomDirectory);refresh();}
    }
    void openDocument(int index) {
        if(index<0||static_cast<size_t>(index)>=documents.size())return;const auto& item=documents[static_cast<size_t>(index)];
        if(item.state==DocumentState::Failed||item.state==DocumentState::Cancelled)showDetails(index);
        else if(!item.output.empty())require(reinterpret_cast<INT_PTR>(ShellExecuteW(window,L"open",item.output.c_str(),nullptr,nullptr,SW_SHOWNORMAL))>32,"Open verified result");
        else if(!item.error.empty()||!item.context.empty())showDetails(index);
    }
    std::wstring detailsText(int index) const {
        const auto& item=documents.at(static_cast<size_t>(index));
        std::wstring text=L"文档："+item.path+L"\r\n结果："+DocumentStates[static_cast<int>(item.state)]+L"\r\n退出码："+(item.exitCode==STILL_ACTIVE?std::wstring(L"尚未结束"):std::to_wstring(item.exitCode))+L"\r\n最后阶段："+item.stage+L"\r\n";
        if(!item.error.empty())text+=L"原因："+item.error+L"\r\n";
        else if(item.state==DocumentState::Pending){const auto problem=environmentProblem();if(!problem.empty())text+=L"未开始原因："+problem+L"\r\n";}
        text+=L"\r\n"+item.context+L"\r\n本次后台日志\r\n";
        // 诊断输出允许替换无效 UTF-8 字节，不能因日志解码影响转换结果。
        const int length=MultiByteToWideChar(CP_UTF8,0,item.trace.data(),static_cast<int>(item.trace.size()),nullptr,0);
        if(length>0){std::wstring decoded(static_cast<size_t>(length),L'\0');MultiByteToWideChar(CP_UTF8,0,item.trace.data(),static_cast<int>(item.trace.size()),decoded.data(),length);text+=decoded;}
        else text+=L"尚无后台日志；任务完成或失败后可在这里查看。";
        return text;
    }
    struct DetailsView {App* app;std::wstring text;};
    static INT_PTR CALLBACK detailsProc(HWND dialog,UINT message,WPARAM wparam,LPARAM lparam) {
        auto data=reinterpret_cast<DetailsView*>(GetWindowLongPtrW(dialog,DWLP_USER));
        if(message==WM_INITDIALOG){
            data=reinterpret_cast<DetailsView*>(lparam);SetWindowLongPtrW(dialog,DWLP_USER,reinterpret_cast<LONG_PTR>(data));auto app=data->app;
            SetWindowTextW(dialog,L"转换详情");
            RECT outer{0,0,app->px(680),app->px(450)},parent{};AdjustWindowRectExForDpi(&outer,WS_POPUP|WS_CAPTION|WS_SYSMENU|DS_MODALFRAME,FALSE,WS_EX_DLGMODALFRAME,app->dpi);GetWindowRect(app->window,&parent);
            SetWindowPos(dialog,nullptr,(parent.left+parent.right-outer.right+outer.left)/2,(parent.top+parent.bottom-outer.bottom+outer.top)/2,outer.right-outer.left,outer.bottom-outer.top,SWP_NOZORDER);
            auto edit=CreateWindowExW(0,L"EDIT",L"",WS_CHILD|WS_VISIBLE|WS_TABSTOP|WS_BORDER|WS_VSCROLL|WS_HSCROLL|ES_MULTILINE|ES_AUTOVSCROLL|ES_AUTOHSCROLL|ES_READONLY,app->px(16),app->px(16),app->px(648),app->px(374),dialog,reinterpret_cast<HMENU>(static_cast<INT_PTR>(301)),app->instance,nullptr);
            require(edit!=nullptr,"Create conversion log view");SendMessageW(edit,WM_SETFONT,reinterpret_cast<WPARAM>(app->body),TRUE);SendMessageW(edit,EM_SETLIMITTEXT,8*1024*1024,0);SetWindowTextW(edit,data->text.c_str());
            for(const auto& button:{std::make_pair(302,L"复制全部"),std::make_pair(IDOK,L"关闭")}){
                auto child=CreateWindowExW(0,L"BUTTON",button.second,WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_OWNERDRAW,app->px(button.first==302?456:564),app->px(404),app->px(100),app->px(32),dialog,reinterpret_cast<HMENU>(static_cast<INT_PTR>(button.first)),app->instance,nullptr);
                require(child!=nullptr,"Create log dialog button");SendMessageW(child,WM_SETFONT,reinterpret_cast<WPARAM>(app->body),TRUE);SetWindowSubclass(child,paintControl,1,reinterpret_cast<DWORD_PTR>(app));
            }
            SetFocus(GetDlgItem(dialog,IDOK));return FALSE;
        }
        if(!data)return FALSE;
        if(message==WM_CTLCOLORSTATIC||message==WM_CTLCOLOREDIT){SetTextColor(reinterpret_cast<HDC>(wparam),RGB(31,43,55));SetBkColor(reinterpret_cast<HDC>(wparam),RGB(255,255,255));return reinterpret_cast<INT_PTR>(data->app->white);}
        if(message==WM_ERASEBKGND){RECT rect{};GetClientRect(dialog,&rect);FillRect(reinterpret_cast<HDC>(wparam),&rect,data->app->white);return TRUE;}
        if(message==WM_DRAWITEM){data->app->drawItem(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam));return TRUE;}
        if(message==WM_COMMAND){
            if(LOWORD(wparam)==302){SendDlgItemMessageW(dialog,301,EM_SETSEL,0,-1);SendDlgItemMessageW(dialog,301,WM_COPY,0,0);return TRUE;}
            if(LOWORD(wparam)==IDOK||LOWORD(wparam)==IDCANCEL){EndDialog(dialog,0);return TRUE;}
        }
        if(message==WM_CLOSE){EndDialog(dialog,0);return TRUE;}return FALSE;
    }
    void showDetails(int index) {
        if(index<0||static_cast<size_t>(index)>=documents.size())return;
        DetailsView data{this,detailsText(index)};
        struct DialogTemplate {DLGTEMPLATE dialog;WORD menu{},windowClass{},title{};} layout{};
        layout.dialog.style=WS_POPUP|WS_CAPTION|WS_SYSMENU|DS_MODALFRAME;layout.dialog.cx=340;layout.dialog.cy=225;
        require(DialogBoxIndirectParamW(instance,&layout.dialog,window,detailsProc,reinterpret_cast<LPARAM>(&data))!=-1,"Open conversion details");
    }
    void queueMenu(LPARAM location) {
        POINT point{static_cast<short>(LOWORD(location)),static_cast<short>(HIWORD(location))};int index=-1;
        if(point.x==-1&&point.y==-1){index=ListView_GetNextItem(control(QueueList),-1,LVNI_SELECTED);if(index>=0){RECT rect{};ListView_GetItemRect(control(QueueList),index,&rect,LVIR_BOUNDS);point={rect.left,rect.bottom};ClientToScreen(control(QueueList),&point);}}
        else{LVHITTESTINFO hit{};hit.pt=point;ScreenToClient(control(QueueList),&hit.pt);index=ListView_SubItemHitTest(control(QueueList),&hit);}
        if(index<0||static_cast<size_t>(index)>=documents.size())return;
        ListView_SetItemState(control(QueueList),index,LVIS_SELECTED|LVIS_FOCUSED,LVIS_SELECTED|LVIS_FOCUSED);
        const auto menu=CreatePopupMenu();require(menu!=nullptr,"Create document context menu");AppendMenuW(menu,MF_STRING,1,L"转换详情");
        const auto command=TrackPopupMenu(menu,TPM_RETURNCMD|TPM_RIGHTBUTTON,point.x,point.y,0,window,nullptr);DestroyMenu(menu);
        if(command==1)showDetails(index);
    }
    static EquationNumbering advancedNumbering(HWND dialog) {
        EquationNumbering settings;settings.custom=IsDlgButtonChecked(dialog,221)==BST_CHECKED;
        auto integerValue=[&](int id){const auto value=settingNumber(windowText(GetDlgItem(dialog,id)),1,999999);require(value==std::floor(value),"Integer number required");return static_cast<long>(value);};
        settings.start=integerValue(232);settings.chapter=integerValue(230);settings.section=integerValue(231);
        if(IsDlgButtonChecked(dialog,221)==BST_CHECKED)settings.format=windowText(GetDlgItem(dialog,229));
        else{
            const wchar_t styles[]=L"1aAiI";const auto separator=windowText(GetDlgItem(dialog,228));
            require(separator.size()<=8,"Separator is at most eight characters");settings.format.clear();
            for(int row=0;row<3;++row)if(row==2||IsDlgButtonChecked(dialog,222+row)==BST_CHECKED){
                if(!settings.format.empty())settings.format+=separator;
                settings.format+=L'#';settings.format+=L"CSE"[row];settings.format+=styles[SendDlgItemMessageW(dialog,224+row,CB_GETCURSEL,0,0)];
            }
            const wchar_t* enclosures[]={L"()",L"[]",L"{}",L"<>",L"（）",L""};
            const auto enclosure=std::wstring(enclosures[SendDlgItemMessageW(dialog,227,CB_GETCURSEL,0,0)]);
            if(!enclosure.empty())settings.format=enclosure.substr(0,1)+settings.format+enclosure.substr(1,1);
        }
        numberParts(settings,settings.start);return settings;
    }
    static void refreshAdvanced(HWND dialog) {
        const bool custom=IsDlgButtonChecked(dialog,221)==BST_CHECKED;
        for(int id:{222,223,224,225,226,227,228})EnableWindow(GetDlgItem(dialog,id),!custom);
        EnableWindow(GetDlgItem(dialog,229),custom);
        for(int row=0;row<2;++row){const auto mode=SendDlgItemMessageW(dialog,201+row,CB_GETCURSEL,0,0);EnableWindow(GetDlgItem(dialog,203+row),mode>=4);SetDlgItemTextW(dialog,205+row,mode==6?L"倍":L"磅");}
        try{
            const auto settings=advancedNumbering(dialog);const auto parts=numberParts(settings,settings.start);
            SetDlgItemTextW(dialog,233,(L"预览："+parts.before+parts.digits+parts.after).c_str());
            if(!custom)SetDlgItemTextW(dialog,229,settings.format.c_str());
        }catch(const std::exception&){SetDlgItemTextW(dialog,233,L"预览：请检查编号格式或数值");}
    }
    static INT_PTR CALLBACK advancedProc(HWND dialog,UINT message,WPARAM wparam,LPARAM lparam) {
        auto app=reinterpret_cast<App*>(GetWindowLongPtrW(dialog,DWLP_USER));
        if(message==WM_INITDIALOG){
            app=reinterpret_cast<App*>(lparam);SetWindowLongPtrW(dialog,DWLP_USER,reinterpret_cast<LONG_PTR>(app));SetWindowTextW(dialog,L"高级设置");
            RECT outer{0,0,app->px(440),app->px(562)},parent{};AdjustWindowRectExForDpi(&outer,WS_POPUP|WS_CAPTION|WS_SYSMENU|DS_MODALFRAME,FALSE,WS_EX_DLGMODALFRAME,app->dpi);GetWindowRect(app->window,&parent);
            SetWindowPos(dialog,nullptr,(parent.left+parent.right-outer.right+outer.left)/2,(parent.top+parent.bottom-outer.bottom+outer.top)/2,outer.right-outer.left,outer.bottom-outer.top,SWP_NOZORDER);
            auto add=[&](int id,const wchar_t* type,const wchar_t* text,DWORD style,int x,int y,int width,int height){
                if(id==IDOK||id==IDCANCEL||id==235)style=(style&~BS_TYPEMASK)|BS_OWNERDRAW;
                if(!wcscmp(type,L"EDIT"))style&=~WS_BORDER;
                if(!wcscmp(type,L"COMBOBOX"))style|=CBS_OWNERDRAWFIXED|CBS_HASSTRINGS;
                auto child=CreateWindowExW(0,type,text,WS_CHILD|WS_VISIBLE|style,app->px(x),app->px(y),app->px(width),app->px(height),dialog,reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),app->instance,nullptr);
                SendMessageW(child,WM_SETFONT,reinterpret_cast<WPARAM>(app->body),TRUE);
                if(!wcscmp(type,L"COMBOBOX")){SetWindowSubclass(child,paintControl,1,reinterpret_cast<DWORD_PTR>(app));SendMessageW(child,CB_SETITEMHEIGHT,0,app->px(28));SendMessageW(child,CB_SETITEMHEIGHT,static_cast<WPARAM>(-1),app->px(28));}
                else if(!wcscmp(type,L"EDIT")){SetWindowSubclass(child,paintInput,1,reinterpret_cast<DWORD_PTR>(app));SetWindowPos(child,nullptr,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_FRAMECHANGED);}
                else if(!wcscmp(type,L"STATIC"))SetWindowSubclass(child,paintLabel,1,reinterpret_cast<DWORD_PTR>(app));
                if(!wcscmp(type,L"BUTTON"))SetWindowSubclass(child,paintControl,1,reinterpret_cast<DWORD_PTR>(app));return child;};
            add(200,L"STATIC",L"公式行距",0,24,16,392,24);
            for(int row=0;row<2;++row){
                const auto& spacing=row?app->displaySpacing:app->inlineSpacing;
                add(210+row,L"STATIC",row?L"行外公式":L"行内公式",0,24,48+row*40,76,34);
                auto mode=add(201+row,L"COMBOBOX",L"",CBS_DROPDOWNLIST|WS_VSCROLL|WS_TABSTOP,104,48+row*40,172,220);
                for(const auto* label:{L"跟随原文",L"单倍行距",L"1.5 倍行距",L"2 倍行距",L"固定值",L"最小值",L"多倍行距"})SendMessageW(mode,CB_ADDSTRING,0,reinterpret_cast<LPARAM>(label));
                SendMessageW(mode,CB_SETCURSEL,spacing.mode,0);wchar_t amount[40]{};swprintf_s(amount,L"%.6g",spacing.value);add(203+row,L"EDIT",amount,WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL,288,48+row*40,80,34);
                add(205+row,L"STATIC",L"磅",0,380,48+row*40,36,34);
            }
            add(208,L"STATIC",L"行内行距作用于整段；默认跟随原文。",0,24,128,392,24);
            add(237,L"STATIC",L"行外右编号",0,24,164,392,24);
            add(212,L"STATIC",L"手动指定章、节和起始序号。",0,24,190,392,24);
            add(220,L"BUTTON",L"简单格式",BS_AUTORADIOBUTTON|WS_GROUP|WS_TABSTOP,24,218,100,28);
            add(221,L"BUTTON",L"自定义格式",BS_AUTORADIOBUTTON|WS_TABSTOP,140,218,120,28);
            CheckRadioButton(dialog,220,221,app->numbering.custom?221:220);
            const wchar_t* styleLabels[]={L"1, 2, 3…",L"a, b, c…",L"A, B, C…",L"i, ii, iii…",L"I, II, III…"};
            for(int row=0;row<3;++row){
                const int y=254+row*38;
                if(row<2){add(222+row,L"BUTTON",row?L"节号":L"章号",BS_AUTOCHECKBOX|WS_TABSTOP,24,y,76,34);CheckDlgButton(dialog,222+row,(row?app->sectionNumber:app->chapterNumber)?BST_CHECKED:BST_UNCHECKED);}
                else add(213,L"STATIC",L"公式序号",0,24,y,76,34);
                auto style=add(224+row,L"COMBOBOX",L"",CBS_DROPDOWNLIST|WS_TABSTOP,104,y,132,180);
                for(const auto* label:styleLabels)SendMessageW(style,CB_ADDSTRING,0,reinterpret_cast<LPARAM>(label));
                SendMessageW(style,CB_SETCURSEL,row==0?app->chapterStyle:row==1?app->sectionStyle:app->equationStyle,0);
                add(214+row,L"STATIC",L"起始值",0,252,y,54,34);
                add(230+row,L"EDIT",std::to_wstring(row==0?app->numbering.chapter:row==1?app->numbering.section:app->numbering.start).c_str(),WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL,320,y,96,34);
            }
            add(217,L"STATIC",L"括号",0,24,368,76,34);auto enclosure=add(227,L"COMBOBOX",L"",CBS_DROPDOWNLIST|WS_TABSTOP,104,368,132,180);
            for(const auto* label:{L"( )",L"[ ]",L"{ }",L"< >",L"（ ）",L"无"})SendMessageW(enclosure,CB_ADDSTRING,0,reinterpret_cast<LPARAM>(label));SendMessageW(enclosure,CB_SETCURSEL,app->numberEnclosure,0);
            add(218,L"STATIC",L"分隔符",0,252,368,54,34);add(228,L"EDIT",app->numberSeparator.c_str(),WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL,320,368,96,34);
            add(219,L"STATIC",L"格式",0,24,410,76,34);add(229,L"EDIT",app->numbering.format.c_str(),WS_BORDER|WS_TABSTOP|ES_AUTOHSCROLL,104,410,312,34);
            SendDlgItemMessageW(dialog,229,EM_SETLIMITTEXT,128,0);
            add(236,L"STATIC",L"#C1 章 · #S1 节 · #E1 序号；样式 1 a A i I",0,24,446,392,22);
            add(233,L"STATIC",L"",SS_ENDELLIPSIS,24,474,392,34);
            add(235,L"BUTTON",L"恢复默认",BS_PUSHBUTTON|WS_TABSTOP,24,518,96,32);
            add(IDOK,L"BUTTON",L"确定",BS_DEFPUSHBUTTON|WS_TABSTOP,232,518,86,32);add(IDCANCEL,L"BUTTON",L"取消",BS_PUSHBUTTON|WS_TABSTOP,330,518,86,32);
            refreshAdvanced(dialog);return TRUE;
        }
        if(message==WM_ERASEBKGND){RECT rect{};GetClientRect(dialog,&rect);FillRect(reinterpret_cast<HDC>(wparam),&rect,app->white);return TRUE;}
        if(message==WM_CTLCOLORSTATIC||message==WM_CTLCOLOREDIT){auto dc=reinterpret_cast<HDC>(wparam);SetTextColor(dc,IsWindowEnabled(reinterpret_cast<HWND>(lparam))?RGB(48,65,79):DisabledTextColor);SetBkColor(dc,RGB(255,255,255));return reinterpret_cast<INT_PTR>(app->white);}
        if(message==WM_DRAWITEM){app->drawItem(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam));return TRUE;}
        if(message==WM_COMMAND){const int id=LOWORD(wparam);
            if(!GetDlgItem(dialog,IDOK))return TRUE;
            if(id==IDCANCEL){EndDialog(dialog,0);return TRUE;}
            if(id==235){for(int mode:{201,202,224,225,226,227})SendDlgItemMessageW(dialog,mode,CB_SETCURSEL,0,0);
                for(int value:{203,204})SetDlgItemTextW(dialog,value,L"20");for(int value:{230,231,232})SetDlgItemTextW(dialog,value,L"1");
                CheckRadioButton(dialog,220,221,220);CheckDlgButton(dialog,222,BST_UNCHECKED);CheckDlgButton(dialog,223,BST_UNCHECKED);SetDlgItemTextW(dialog,228,L".");refreshAdvanced(dialog);return TRUE;}
            if(id==IDOK){
                try{
                    EquationSpacing spacing[2];for(int row=0;row<2;++row){auto& value=spacing[row];value.mode=static_cast<int>(SendDlgItemMessageW(dialog,201+row,CB_GETCURSEL,0,0));if(value.mode>=4)value.value=settingNumber(windowText(GetDlgItem(dialog,203+row)),value.mode==6?0.5:1,value.mode==6?10:999);}
                    const auto numbers=advancedNumbering(dialog);app->inlineSpacing=spacing[0];app->displaySpacing=spacing[1];app->numbering=numbers;
                    app->numbering.custom=IsDlgButtonChecked(dialog,221)==BST_CHECKED;app->chapterNumber=IsDlgButtonChecked(dialog,222)==BST_CHECKED;app->sectionNumber=IsDlgButtonChecked(dialog,223)==BST_CHECKED;
                    app->chapterStyle=static_cast<int>(SendDlgItemMessageW(dialog,224,CB_GETCURSEL,0,0));app->sectionStyle=static_cast<int>(SendDlgItemMessageW(dialog,225,CB_GETCURSEL,0,0));app->equationStyle=static_cast<int>(SendDlgItemMessageW(dialog,226,CB_GETCURSEL,0,0));app->numberEnclosure=static_cast<int>(SendDlgItemMessageW(dialog,227,CB_GETCURSEL,0,0));app->numberSeparator=windowText(GetDlgItem(dialog,228));
                    EndDialog(dialog,IDOK);
                }catch(const std::exception&){MessageBoxW(dialog,L"请检查设置：固定/最小行距为 1–999 磅，多倍为 0.5–10 倍；编号数值为 1–999999 的整数，格式必须包含且仅包含一个 #E 序号。",L"设置有误",MB_OK|MB_ICONWARNING);}return TRUE;
            }
            if(HIWORD(wparam)==CBN_SELCHANGE||HIWORD(wparam)==BN_CLICKED||(HIWORD(wparam)==EN_CHANGE&&id!=229))refreshAdvanced(dialog);
            else if(id==229&&HIWORD(wparam)==EN_CHANGE&&IsDlgButtonChecked(dialog,221)==BST_CHECKED)refreshAdvanced(dialog);
        }
        if(message==WM_CLOSE){EndDialog(dialog,0);return TRUE;}return FALSE;
    }
    void advancedSettings() {
        struct DialogTemplate { DLGTEMPLATE dialog; WORD menu{},windowClass{},title{}; } layout{};
        layout.dialog.style=WS_POPUP|WS_CAPTION|WS_SYSMENU|DS_MODALFRAME;layout.dialog.cx=260;layout.dialog.cy=140;
        require(DialogBoxIndirectParamW(instance,&layout.dialog,window,advancedProc,reinterpret_cast<LPARAM>(this))!=-1,"Open advanced settings");
    }
    void testAdvance() {
        require(!workerFailed,"GUI worker acceptance");
        if(testStep==0){std::string trace;appendWorkerLog(trace,std::string(MaximumLogBytes+32,'x'));appendWorkerLog(trace,"LAST_FAILURE");require(trace.size()<=MaximumLogBytes&&trace.find("LAST_FAILURE")!=std::string::npos&&trace.find("truncated")!=std::string::npos,"Worker log is bounded and retains its final error");}
        if(testStep==-1){testStep=0;SendMessageW(control(Reverse),BM_CLICK,0,0);require(reverseMode&&!busy,"Direction switch does not launch Office");return;}
        if(testStep==0){
            RECT client{};GetClientRect(window,&client);
            for(HWND child=GetWindow(window,GW_CHILD);child;child=GetWindow(child,GW_HWNDNEXT)){
                if(!(GetWindowLongPtrW(child,GWL_STYLE)&WS_VISIBLE))continue;
                RECT bounds{};GetWindowRect(child,&bounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&bounds),2);
                require(bounds.left>=0&&bounds.top>=0&&bounds.right<=client.right&&bounds.bottom<=client.bottom,"Visible controls fit client area");
            }
            testLayout=true;
            require(!busy&&task.empty()&&!scanned,"File selection does not launch a scan");
            require(GetNextDlgTabItem(window,control(Choose),FALSE)==control(AddFolder),"Keyboard tab navigation");
            if(!reverseMode){
                // 先完成初次绘制，再确认参数交互不会使父窗口重画或覆盖子控件。
                ShowWindow(window,SW_SHOWNOACTIVATE);RedrawWindow(window,nullptr,nullptr,RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW);
                require((GetWindowLongPtrW(window,GWL_STYLE)&WS_CLIPCHILDREN)!=0,"Parent drawing excludes child controls");
                RECT sizeBounds{};GetWindowRect(control(SizeBox),&sizeBounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&sizeBounds),2);
                HDC parentDc=GetDC(window);const bool excluded=!PtVisible(parentDc,(sizeBounds.left+sizeBounds.right)/2,(sizeBounds.top+sizeBounds.bottom)/2);ReleaseDC(window,parentDc);
                require(excluded,"Parent DC cannot overwrite the font size input");
                // 按实际命中控件发送鼠标消息，覆盖输入框遮挡按钮和增减方向的问题。
                auto arrow=[&](bool up){RECT bounds{};GetWindowRect(control(SizeSpin),&bounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&bounds),2);
                    POINT point{(bounds.left+bounds.right)/2,up?bounds.top+(bounds.bottom-bounds.top)/4:bounds.top+3*(bounds.bottom-bounds.top)/4};
                    HWND hit=ChildWindowFromPointEx(window,point,CWP_SKIPINVISIBLE|CWP_SKIPDISABLED);require(hit==control(SizeSpin),"Font size arrow receives mouse click");
                    MapWindowPoints(window,hit,&point,1);const LPARAM position=MAKELPARAM(point.x,point.y);
                    SendMessageW(hit,WM_LBUTTONDOWN,MK_LBUTTON,position);SendMessageW(hit,WM_LBUTTONUP,0,position);};
                auto sizeIs=[&](double expected){require(wcstod(windowText(control(SizeBox)).c_str(),nullptr)==expected,"Font size arrow value");};
                require(SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==0&&!IsWindowEnabled(control(SizeBox)),"Default size follows source");
                SendMessageW(control(FollowSize),CB_SETCURSEL,1,0);layout();refresh();
                set(SizeBox,L"12");const auto refreshesBefore=testRefreshes;arrow(true);sizeIs(13);
                require(testRefreshes==refreshesBefore+1,"Font size arrow refreshes state only once");arrow(false);sizeIs(12);
                set(SizeBox,L"11.5");arrow(true);sizeIs(12.5);arrow(false);sizeIs(11.5);
                SendMessageW(control(SizeBox),WM_KEYDOWN,VK_UP,0);SendMessageW(control(SizeBox),WM_KEYUP,VK_UP,0);sizeIs(12.5);
                SendMessageW(control(SizeBox),WM_KEYDOWN,VK_DOWN,0);SendMessageW(control(SizeBox),WM_KEYUP,VK_DOWN,0);sizeIs(11.5);
                set(SizeBox,L"1");arrow(false);sizeIs(1);set(SizeBox,L"256");arrow(true);sizeIs(256);
                SendMessageW(control(ModeBox),CB_SETCURSEL,1,0);refresh();require(!IsWindowEnabled(control(AddNumbers)),"Inline mode has no right numbers");
                SendMessageW(control(ModeBox),CB_SETCURSEL,0,0);set(SizeBox,L"0");refresh();require(!IsWindowEnabled(control(Start)),"Reject zero size");
                set(SizeBox,L"NaN");refresh();require(!IsWindowEnabled(control(Start)),"Reject nonfinite size");
                set(SizeBox,L"12");
                RedrawWindow(window,nullptr,nullptr,RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW);const auto paintsBefore=testBackgroundPaints;
                for(int change=0;change<20;++change){SendMessageW(control(ModeBox),CB_SETCURSEL,change%2?0:2,0);
                    SendMessageW(window,WM_COMMAND,MAKEWPARAM(ModeBox,CBN_SELCHANGE),reinterpret_cast<LPARAM>(control(ModeBox)));
                    arrow(change%2==0);UpdateWindow(window);}
                require(testBackgroundPaints==paintsBefore,"Mode and size interactions do not repaint the parent");testParameterPaint=true;
            }else{require(!(GetWindowLongPtrW(control(ModeBox),GWL_STYLE)&WS_VISIBLE)&&(GetWindowLongPtrW(control(Bracket),GWL_STYLE)&WS_VISIBLE),"Reverse format controls");
                const bool dollar=SendMessageW(control(Dollar),BM_GETCHECK,0,0)==BST_CHECKED;
                CheckRadioButton(window,Bracket,Dollar,0);refresh();require(!IsWindowEnabled(control(Start)),"Reject missing delimiter");
                CheckRadioButton(window,Bracket,Dollar,dollar?Dollar:Bracket);set(SizeBox,L"NaN");refresh();require(IsWindowEnabled(control(Start)),"Ignore reverse hidden size settings");}
            set(SizeBox,L"18");set(NameBox,L":");refresh();require(!IsWindowEnabled(control(Start)),"Reject invalid output suffix");
            set(NameBox,L"_"+testMode);refresh();require(IsWindowEnabled(control(Start)),"Enable without pre-scan");testControls=true;
            saveNew(testDirectory+L"\\gui-ready.json","{\"NoPreScan\":true}\n");testStep=1;SetTimer(window,2,1500,nullptr);return;
        }
        require(summary&&(outcome==L"COMPLETE"||outcome==L"PARTIAL"),"GUI outcome");
        require(testScanProgress&&testSavedProgress&&testReopenMarquee&&!progressMarquee,"Scan and saved verification counters with marquee Word reopen");
        refresh();RECT client{};GetClientRect(window,&client);
        for(int id:{OpenOutput,OpenFolder,OpenReport})if(GetWindowLongPtrW(control(id),GWL_STYLE)&WS_VISIBLE){
            RECT bounds{};GetWindowRect(control(id),&bounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&bounds),2);
            require(bounds.bottom<=client.bottom-px(96),"Result buttons stay inside their card");}
        if(!reverseMode){
            // 完成摘要显示后改字号同样只更新输入框，保留公式计数和结果按钮。
            refresh();
            RedrawWindow(window,nullptr,nullptr,RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW);const auto paintsBefore=testBackgroundPaints;
            const auto savedSize=windowText(control(SizeBox));set(SizeBox,L"18.5");set(SizeBox,savedSize);UpdateWindow(window);
            require(testBackgroundPaints==paintsBefore&&scanned&&!documents.front().output.empty(),"Changing size preserves completed summary without background repaint");
        }
        std::string json="{\"Status\":\"PASSED\",\"MathTypeDirectory\":"+quote(install)+",\"Input\":"+quote(input)+",\"Output\":"+quote(output)+",\"Report\":"+quote(testDirectory+L"\\report.json")+
            ",\"Outcome\":"+quote(outcome)+(reverseMode?",\"Direction\":\"REVERSE\",\"DisplayDelimiter\":"+quote(testMode==L"reverse-dollar"?L"dollar":L"bracket"):",\"DefaultFont\":true,\"RequestedSizePt\":18")+
            ",\"ControlValidation\":true,\"ScanProgress\":true,\"SavedProgress\":true,\"ReopenMarquee\":true,\"SizeArrows\":"+(reverseMode?std::string("false"):std::string("true"))+",\"ParameterRepaint\":"+(testParameterPaint?std::string("true"):std::string("false"))+",\"KeyboardTabNavigation\":true,\"Layout\":true,\"NoPreScan\":true,\"Succeeded\":"+std::to_string(succeeded)+",\"FailedPreserved\":"+std::to_string(failed)+",\"NotAttempted\":"+std::to_string(unattempted)+"}\n";
        saveNew(testDirectory+L"\\gui-check.json",json);testStep=2;SetTimer(window,3,1500,nullptr);
    }
    static LRESULT CALLBACK paintInput(HWND child,UINT message,WPARAM wparam,LPARAM lparam,UINT_PTR,DWORD_PTR data) {
        auto app=reinterpret_cast<App*>(data);
        // 禁用 EDIT 的系统灰字不采用父窗口色值，改用与其他禁用控件一致的绘制。
        if(!IsWindowEnabled(child)&&(message==WM_PAINT||message==WM_PRINTCLIENT)){
            PAINTSTRUCT painting{};HDC dc=message==WM_PAINT?BeginPaint(child,&painting):reinterpret_cast<HDC>(wparam);RECT rect{};GetClientRect(child,&rect);
            FillRect(dc,&rect,app->white);app->drawText(dc,rect,windowText(child),DisabledTextColor,app->bodyText.Get());
            if(message==WM_PAINT)EndPaint(child,&painting);return 0;
        }
        if(message==WM_ENABLE){const auto result=DefSubclassProc(child,message,wparam,lparam);RedrawWindow(child,nullptr,nullptr,RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);return result;}
        if(message==WM_NCCALCSIZE){
            auto rect=wparam?&reinterpret_cast<NCCALCSIZE_PARAMS*>(lparam)->rgrc[0]:reinterpret_cast<RECT*>(lparam);
            rect->left+=app->px(10);rect->right-=app->px(10);rect->top+=app->px(8);rect->bottom-=app->px(8);return 0;
        }
        if(message==WM_NCPAINT){
            HDC dc=GetWindowDC(child);RECT rect{};GetWindowRect(child,&rect);OffsetRect(&rect,-rect.left,-rect.top);
            ExcludeClipRect(dc,app->px(10),app->px(8),rect.right-app->px(10),rect.bottom-app->px(8));
            FillRect(dc,&rect,app->white);
            // 字号输入与箭头共用一个外轮廓，接缝处不画第二条边框。
            if(GetDlgCtrlID(child)==SizeBox)rect.right+=app->px(6);
            app->rounded(dc,rect,RGB(255,255,255),GetFocus()==child?RGB(13,128,131):RGB(207,216,224),6);ReleaseDC(child,dc);return 0;
        }
        if(message==WM_SETFOCUS||message==WM_KILLFOCUS){const auto result=DefSubclassProc(child,message,wparam,lparam);
            RedrawWindow(child,nullptr,nullptr,RDW_INVALIDATE|RDW_FRAME);
            if(GetDlgCtrlID(child)==SizeBox)InvalidateRect(app->control(SizeSpin),nullptr,FALSE);return result;}
        return DefSubclassProc(child,message,wparam,lparam);
    }
    void chevron(HDC dc,int x,int y,bool up,COLORREF color) const {
        Gdiplus::Graphics graphics(dc);graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        Gdiplus::Pen pen(Gdiplus::Color(255,GetRValue(color),GetGValue(color),GetBValue(color)),static_cast<float>(px(1)));
        pen.SetStartCap(Gdiplus::LineCapRound);pen.SetEndCap(Gdiplus::LineCapRound);
        const float step=static_cast<float>(px(3)),cx=static_cast<float>(x),cy=static_cast<float>(y),sign=up?-1.0f:1.0f;
        graphics.DrawLine(&pen,cx-step,cy-sign*step/2,cx,cy+sign*step/2);graphics.DrawLine(&pen,cx,cy+sign*step/2,cx+step,cy-sign*step/2);
    }
    static LRESULT CALLBACK paintQueueHeader(HWND child,UINT message,WPARAM wparam,LPARAM lparam,UINT_PTR,DWORD_PTR data) {
        auto app=reinterpret_cast<App*>(data);
        if(message==HDM_LAYOUT){
            const auto result=DefSubclassProc(child,message,wparam,lparam);auto layout=reinterpret_cast<HDLAYOUT*>(lparam);
            layout->prc->top+=app->px(34)-layout->pwpos->cy;layout->pwpos->cy=app->px(34);return result;
        }
        if(message==WM_ERASEBKGND)return 1;
        if(message==WM_PAINT||message==WM_PRINTCLIENT){
            PAINTSTRUCT painting{};HDC target=message==WM_PAINT?BeginPaint(child,&painting):reinterpret_cast<HDC>(wparam);RECT rect{};GetClientRect(child,&rect);
            HDC dc=CreateCompatibleDC(target);HBITMAP bitmap=CreateCompatibleBitmap(target,std::max(rect.right,1L),std::max(rect.bottom,1L));auto old=SelectObject(dc,bitmap);
            FillRect(dc,&rect,app->white);app->rounded(dc,rect,RGB(244,247,249),RGB(244,247,249),6);
            for(int column=0;column<Header_GetItemCount(child);++column){
                RECT label{};Header_GetItemRect(child,column,&label);wchar_t text[40]{};HDITEMW item{};item.mask=HDI_TEXT|HDI_FORMAT;item.pszText=text;item.cchTextMax=40;
                Header_GetItem(child,column,&item);const bool centered=(item.fmt&HDF_JUSTIFYMASK)==HDF_CENTER;
                if(!centered){label.left+=app->px(10);label.right-=app->px(4);}
                app->drawText(dc,label,text,RGB(104,120,133),app->bodyText.Get(),centered);
            }
            BitBlt(target,0,0,rect.right,rect.bottom,dc,0,0,SRCCOPY);SelectObject(dc,old);DeleteObject(bitmap);DeleteDC(dc);
            if(message==WM_PAINT)EndPaint(child,&painting);return 0;
        }
        return DefSubclassProc(child,message,wparam,lparam);
    }
    static LRESULT CALLBACK paintControl(HWND child,UINT message,WPARAM wparam,LPARAM lparam,UINT_PTR,DWORD_PTR data) {
        auto app=reinterpret_cast<App*>(data);const int id=GetDlgCtrlID(child);
        const bool dropdown=id==ModeBox||id==FollowSize||id==AddNumbers||(GetParent(child)!=app->window&&(id==201||id==202||(id>=224&&id<=227)));
        const bool choice=id==Bracket||id==Dollar||id==OriginalDirectory||id==CustomDirectory||(GetParent(child)!=app->window&&id>=220&&id<=223);
        if(id==QueueList&&message==WM_DROPFILES)return SendMessageW(app->window,message,wparam,lparam);
        if(dropdown&&message==WM_NCCALCSIZE)return 0;
        if((dropdown||id==SizeSpin)&&message==WM_NCPAINT)return 0;
        if(message==WM_MOUSEMOVE){RECT rect{};GetClientRect(child,&rect);const int half=HIWORD(lparam)<rect.bottom/2?0:1;
            if(app->hotControl!=child||app->hotHalf!=half){app->hotControl=child;app->hotHalf=half;InvalidateRect(child,nullptr,FALSE);}
            // 开发预览固定合成鼠标状态，避免隐藏窗口的实际指针位置清除截图状态。
            if(app->testMode==L"preview"&&!dropdown&&id!=SizeSpin)return 0;
            TRACKMOUSEEVENT track{sizeof(track),TME_LEAVE,child,0};TrackMouseEvent(&track);}
        if(message==WM_MOUSELEAVE){if(app->hotControl==child)app->hotControl=nullptr;InvalidateRect(child,nullptr,FALSE);}
        if(message==WM_LBUTTONDOWN){app->pressedControl=child;InvalidateRect(child,nullptr,FALSE);}
        if(message==WM_LBUTTONUP||message==WM_CAPTURECHANGED){if(app->pressedControl==child)app->pressedControl=nullptr;InvalidateRect(child,nullptr,FALSE);}
        // 箭头的绘制和点击共用右侧 36 px 区域，不沿用原生旧按钮的窄命中范围。
        if(dropdown&&IsWindowEnabled(child)&&(message==WM_LBUTTONDOWN||message==WM_LBUTTONUP)){
            RECT rect{};GetClientRect(child,&rect);
            if(static_cast<short>(LOWORD(lparam))>=rect.right-app->px(36)){
                if(message==WM_LBUTTONDOWN){SetFocus(child);SendMessageW(child,CB_SHOWDROPDOWN,!SendMessageW(child,CB_GETDROPPEDSTATE,0,0),0);}
                return 0;
            }
        }
        if(id==Progress&&message==PBM_SETMARQUEE){const auto result=DefSubclassProc(child,message,wparam,lparam);
            if(wparam)SetTimer(child,91,40,nullptr);else KillTimer(child,91);InvalidateRect(child,nullptr,FALSE);return result;}
        if(id==Progress&&message==WM_TIMER&&wparam==91){InvalidateRect(child,nullptr,FALSE);return 0;}
        if((dropdown||id==SizeSpin||id==Progress||choice)&&message==WM_ERASEBKGND)return 1;
        if((dropdown||id==SizeSpin||id==Progress||choice)&&(message==WM_PAINT||message==WM_PRINTCLIENT)){
            PAINTSTRUCT painting{};HDC target=message==WM_PAINT?BeginPaint(child,&painting):reinterpret_cast<HDC>(wparam);RECT rect{};GetClientRect(child,&rect);
            HDC dc=CreateCompatibleDC(target);HBITMAP bitmap=CreateCompatibleBitmap(target,std::max(rect.right,1L),std::max(rect.bottom,1L));auto old=SelectObject(dc,bitmap);
            FillRect(dc,&rect,app->white);const bool enabled=IsWindowEnabled(child)!=FALSE;
            if(id==Progress){
                app->rounded(dc,rect,RGB(230,235,239),RGB(230,235,239),3);RECT value=rect;
                if(app->progressMarquee){const int length=std::max(static_cast<int>(rect.right)/4,app->px(24));const auto tick=GetTickCount64()%1600;
                    value.left=static_cast<LONG>((rect.right+length)*tick/1600)-length;value.right=value.left+length;
                }else{PBRANGE range{};SendMessageW(child,PBM_GETRANGE,FALSE,reinterpret_cast<LPARAM>(&range));
                    const auto position=SendMessageW(child,PBM_GETPOS,0,0);value.right=range.iHigh>range.iLow?MulDiv(rect.right,static_cast<int>(position)-range.iLow,range.iHigh-range.iLow):0;}
                if(value.right>value.left)app->rounded(dc,value,RGB(19,155,158),RGB(19,155,158),3);
            }else if(dropdown){
                const bool focus=GetFocus()==child||GetParent(GetFocus())==child;
                app->rounded(dc,rect,RGB(255,255,255),focus?RGB(13,128,131):app->hotControl==child&&enabled?RGB(153,184,188):RGB(207,216,224),6);
                RECT label=rect;label.left+=app->px(10);label.right-=app->px(36);
                app->drawText(dc,label,windowText(child),enabled?RGB(48,65,79):DisabledTextColor,app->bodyText.Get());
                app->chevron(dc,rect.right-app->px(18),rect.bottom/2,false,enabled?RGB(67,85,98):DisabledTextColor);
            }else if(choice){
                const bool selected=SendMessageW(child,BM_GETCHECK,0,0)==BST_CHECKED;
                const bool delimiter=id==Bracket||id==Dollar,checkbox=(GetWindowLongPtrW(child,GWL_STYLE)&BS_TYPEMASK)==BS_AUTOCHECKBOX;
                if(delimiter)app->rounded(dc,rect,selected&&enabled?RGB(239,249,249):RGB(255,255,255),selected&&enabled?RGB(151,196,200):RGB(207,216,224),6);
                const auto color=enabled?RGB(13,128,131):DisabledTextColor;const int offset=delimiter?10:2;
                {Gdiplus::Graphics graphics(dc);graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
                    Gdiplus::Pen ring(Gdiplus::Color(255,GetRValue(color),GetGValue(color),GetBValue(color)),static_cast<float>(app->px(1)));
                    const auto x=static_cast<float>(app->px(offset)),y=static_cast<float>(rect.bottom/2-app->px(6)),size=static_cast<float>(app->px(12));
                    if(checkbox){graphics.DrawRectangle(&ring,x,y,size,size);if(selected){graphics.DrawLine(&ring,x+app->px(2),y+app->px(6),x+app->px(5),y+app->px(9));graphics.DrawLine(&ring,x+app->px(5),y+app->px(9),x+app->px(10),y+app->px(3));}}
                    else{graphics.DrawEllipse(&ring,x,y,size,size);if(selected){Gdiplus::SolidBrush dot(Gdiplus::Color(255,GetRValue(color),GetGValue(color),GetBValue(color)));
                        const auto inset=static_cast<float>(app->px(3));graphics.FillEllipse(&dot,x+inset,y+inset,size-2*inset,size-2*inset);}}}
                RECT text=rect;text.left+=app->px(offset+20);app->drawText(dc,text,windowText(child),enabled?RGB(48,65,79):DisabledTextColor,app->bodyText.Get());
                if(GetFocus()==child){RECT focus=rect;InflateRect(&focus,-app->px(3),-app->px(3));if(!delimiter)focus.left=app->px(19);app->rounded(dc,focus,CLR_INVALID,RGB(13,128,131),4);}
            }else{
                RECT frame=rect;frame.left-=app->px(6);
                app->rounded(dc,frame,RGB(250,252,253),GetFocus()==app->control(SizeBox)?RGB(13,128,131):RGB(207,216,224),6);
                if(enabled&&(app->hotControl==child||app->pressedControl==child)){RECT hot=rect;hot.left+=app->px(3);hot.right-=app->px(3);
                    if(app->hotHalf==0)hot.bottom=rect.bottom/2;else hot.top=rect.bottom/2;
                    InflateRect(&hot,0,-app->px(2));const auto fill=app->pressedControl==child?RGB(211,231,233):RGB(231,243,244);app->rounded(dc,hot,fill,fill,3);}
                app->chevron(dc,rect.right/2,rect.bottom/4,true,enabled?RGB(67,85,98):DisabledTextColor);
                app->chevron(dc,rect.right/2,rect.bottom*3/4,false,enabled?RGB(67,85,98):DisabledTextColor);
            }
            BitBlt(target,0,0,rect.right,rect.bottom,dc,0,0,SRCCOPY);SelectObject(dc,old);DeleteObject(bitmap);DeleteDC(dc);
            if(message==WM_PAINT)EndPaint(child,&painting);return 0;
        }
        const auto result=DefSubclassProc(child,message,wparam,lparam);
        if(message==WM_ENABLE||message==WM_SETFOCUS||message==WM_KILLFOCUS||message==CB_SETEDITSEL||(choice&&message==BM_SETCHECK)||
           (id==Progress&&(message==PBM_SETPOS||message==PBM_SETRANGE32)))InvalidateRect(child,nullptr,FALSE);
        if((message==WM_SETFOCUS||message==WM_KILLFOCUS)&&GetParent(child)==app->control(ModeBox))InvalidateRect(app->control(ModeBox),nullptr,FALSE);
        return result;
    }
    void drawText(HDC dc,const RECT& rect,const std::wstring& value,COLORREF color,IDWriteTextFormat* format,bool centered=false,bool multiline=false) const {
        check(textTarget->BindDC(dc,&rect),"Bind native text context");textTarget->SetDpi(static_cast<float>(dpi),static_cast<float>(dpi));
        format->SetTextAlignment(centered?DWRITE_TEXT_ALIGNMENT_CENTER:DWRITE_TEXT_ALIGNMENT_LEADING);
        format->SetParagraphAlignment(multiline?DWRITE_PARAGRAPH_ALIGNMENT_NEAR:DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        format->SetWordWrapping(multiline?DWRITE_WORD_WRAPPING_WRAP:DWRITE_WORD_WRAPPING_NO_WRAP);
        ComPtr<ID2D1SolidColorBrush> brush;check(textTarget->CreateSolidColorBrush(D2D1::ColorF(GetRValue(color)/255.0f,GetGValue(color)/255.0f,GetBValue(color)/255.0f),&brush),"Create native text brush");
        const float scale=static_cast<float>(dpi)/96;const auto box=D2D1::RectF(0,0,(rect.right-rect.left)/scale,(rect.bottom-rect.top)/scale);
        textTarget->BeginDraw();textTarget->DrawTextW(value.c_str(),static_cast<UINT32>(value.size()),format,box,brush.Get(),D2D1_DRAW_TEXT_OPTIONS_CLIP);check(textTarget->EndDraw(),"Draw native interface text");
    }
    static LRESULT CALLBACK paintLabel(HWND child,UINT message,WPARAM wparam,LPARAM lparam,UINT_PTR,DWORD_PTR data) {
        auto app=reinterpret_cast<App*>(data);
        // STATIC 在更新文字时会直接用 GDI 绘制；立即覆盖为统一字体和垂直居中的完整帧。
        if(message==WM_SETTEXT||message==WM_ENABLE){const auto result=DefSubclassProc(child,message,wparam,lparam);RedrawWindow(child,nullptr,nullptr,RDW_INVALIDATE|RDW_UPDATENOW);return result;}
        if(message==WM_ERASEBKGND)return 1;
        if(message==WM_PAINT||message==WM_PRINTCLIENT){
            PAINTSTRUCT painting{};HDC dc=message==WM_PAINT?BeginPaint(child,&painting):reinterpret_cast<HDC>(wparam);RECT rect{};GetClientRect(child,&rect);
            // 状态计数经常更新，先画完整文字帧再贴回控件，避免露出清空后的背景。
            HDC buffer=CreateCompatibleDC(dc);HBITMAP bitmap=CreateCompatibleBitmap(dc,std::max(rect.right,1L),std::max(rect.bottom,1L));auto oldBitmap=SelectObject(buffer,bitmap);
            const int id=GetDlgCtrlID(child);const bool outer=id==Title||id==Status||id==Elapsed||id==Environment;
            FillRect(buffer,&rect,outer?app->background:app->white);
            const bool muted=id==StyleNote||id==OutputNote||id==Existing||id==Elapsed||id==Environment||id==InlineDelimiterNote||
                (GetParent(child)!=app->window&&(id==208||id==212||id==236));
            auto format=id==Title?app->titleText.Get():(id==FileHeading||id==StyleHeading||id==OutputHeading||id==200||id==237)?app->headingText.Get():app->bodyText.Get();
            const bool preview=GetParent(child)!=app->window&&id==233;
            if(preview){app->rounded(buffer,rect,RGB(239,249,249),RGB(151,196,200),6);rect.left+=app->px(12);rect.right-=app->px(12);}
            app->drawText(buffer,rect,windowText(child),!IsWindowEnabled(child)?DisabledTextColor:preview?RGB(9,126,132):muted?MutedTextColor:RGB(31,43,55),format,false,id==Detail||id==StyleNote);
            if(preview)GetClientRect(child,&rect);
            BitBlt(dc,0,0,rect.right,rect.bottom,buffer,0,0,SRCCOPY);SelectObject(buffer,oldBitmap);DeleteObject(bitmap);DeleteDC(buffer);
            if(message==WM_PAINT)EndPaint(child,&painting);return 0;
        }
        return DefSubclassProc(child,message,wparam,lparam);
    }
    // 抗锯齿圆角统一用于卡片和按钮；不引入额外界面框架。
    void rounded(HDC dc,const RECT& rect,COLORREF fill,COLORREF border,int radius=8,COLORREF bottom=CLR_INVALID) const {
        Gdiplus::Graphics graphics(dc);graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        const float x=static_cast<float>(rect.left)+0.5f,y=static_cast<float>(rect.top)+0.5f;
        const float w=static_cast<float>(rect.right-rect.left)-1,h=static_cast<float>(rect.bottom-rect.top)-1,r=std::min(static_cast<float>(px(radius)*2),std::min(w,h));
        Gdiplus::GraphicsPath path;path.AddArc(x,y,r,r,180,90);path.AddArc(x+w-r,y,r,r,270,90);path.AddArc(x+w-r,y+h-r,r,r,0,90);path.AddArc(x,y+h-r,r,r,90,90);path.CloseFigure();
        const auto color=[](COLORREF c){return Gdiplus::Color(255,GetRValue(c),GetGValue(c),GetBValue(c));};
        Gdiplus::LinearGradientBrush brush(Gdiplus::PointF(x,y),Gdiplus::PointF(x,y+h),color(fill),color(bottom==CLR_INVALID?fill:bottom));
        Gdiplus::Pen pen(Gdiplus::Color(255,GetRValue(border),GetGValue(border),GetBValue(border)),1);
        if(fill!=CLR_INVALID)graphics.FillPath(&brush,&path);graphics.DrawPath(&pen,&path);
    }
    void drawItem(const DRAWITEMSTRUCT& draw) const {
        if(draw.CtlType==ODT_COMBOBOX){
            // 系统会在焦点/选择变化时直接重画选中项；统一接管，避免蓝底覆盖圆角外观。
            const bool closed=(draw.itemState&ODS_COMBOBOXEDIT)!=0,enabled=IsWindowEnabled(draw.hwndItem)!=FALSE;
            const auto fill=CreateSolidBrush(!closed&&(draw.itemState&ODS_SELECTED)?RGB(231,243,244):RGB(255,255,255));
            FillRect(draw.hDC,&draw.rcItem,fill);DeleteObject(fill);
            if(draw.itemID!=static_cast<UINT>(-1)){
                const auto length=SendMessageW(draw.hwndItem,CB_GETLBTEXTLEN,draw.itemID,0);
                std::wstring value(static_cast<size_t>(length)+1,L'\0');SendMessageW(draw.hwndItem,CB_GETLBTEXT,draw.itemID,reinterpret_cast<LPARAM>(value.data()));value.resize(static_cast<size_t>(length));
                RECT label=draw.rcItem;
                if(closed){RECT client{};GetClientRect(draw.hwndItem,&client);label.left=px(10);label.right=client.right-px(36);}
                else{label.left+=px(10);label.right-=px(10);}
                drawText(draw.hDC,label,value,enabled?RGB(48,65,79):DisabledTextColor,bodyText.Get());
            }
            return;
        }
        const int id=static_cast<int>(draw.CtlID);const bool enabled=IsWindowEnabled(draw.hwndItem)!=0;
        const bool mode=id==Forward||id==Reverse,selected=(id==Forward&&!reverseMode)||(id==Reverse&&reverseMode),primary=id==Start||(id==IDOK&&GetParent(draw.hwndItem)!=window);
        const bool pressed=(draw.itemState&ODS_SELECTED)!=0,hot=hotControl==draw.hwndItem&&enabled;
        RECT r=draw.rcItem;HDC dc=CreateCompatibleDC(draw.hDC);HBITMAP bitmap=CreateCompatibleBitmap(draw.hDC,r.right,r.bottom);auto old=SelectObject(dc,bitmap);
        FillRect(dc,&r,(id==Forward||id==Reverse||id==Start||id==Stop)?background:white);
        COLORREF top,bottom,border,textColor;
        if(!enabled){top=bottom=RGB(234,239,242);border=RGB(222,229,234);textColor=DisabledTextColor;}
        else if(primary){top=pressed?RGB(10,113,118):hot?RGB(24,169,174):RGB(20,157,162);bottom=pressed?RGB(8,102,108):hot?RGB(10,139,145):RGB(8,126,133);
            border=pressed?RGB(8,102,108):RGB(14,144,149);textColor=RGB(255,255,255);}
        else{top=pressed?RGB(225,239,241):hot?RGB(244,251,251):RGB(255,255,255);bottom=pressed?RGB(225,239,241):hot?RGB(235,246,247):RGB(246,248,250);
            border=hot?RGB(153,190,194):RGB(207,216,224);textColor=selected?RGB(9,126,132):RGB(48,65,79);}
        if(mode){top=selected?RGB(255,255,255):hot?RGB(227,235,239):RGB(237,242,245);bottom=top;border=selected?RGB(213,223,229):top;
            if(!enabled)textColor=DisabledTextColor;rounded(dc,r,top,border,6,bottom);}
        else{
            if(enabled&&!pressed){RECT shadow=r;shadow.top+=px(2);rounded(dc,shadow,primary?RGB(217,234,236):RGB(230,235,239),primary?RGB(217,234,236):RGB(230,235,239),7);}
            r.bottom-=px(2);rounded(dc,r,top,border,7,bottom);
        }
        drawText(dc,r,windowText(draw.hwndItem),textColor,primary?headingText.Get():bodyText.Get(),true);
        if(draw.itemState&ODS_FOCUS){RECT focus=r;InflateRect(&focus,-px(3),-px(3));rounded(dc,focus,CLR_INVALID,primary?RGB(210,246,246):RGB(13,128,131),4);}
        BitBlt(draw.hDC,0,0,draw.rcItem.right,draw.rcItem.bottom,dc,0,0,SRCCOPY);SelectObject(dc,old);DeleteObject(bitmap);DeleteDC(dc);
    }
    void paint(HDC destination=nullptr) {
        if(!destination&&!testDirectory.empty())++testBackgroundPaints;
        PAINTSTRUCT painting{};HDC target=destination?destination:BeginPaint(window,&painting);RECT client{};GetClientRect(window,&client);
        HDC dc=CreateCompatibleDC(target);HBITMAP bitmap=CreateCompatibleBitmap(target,std::max(client.right,1L),std::max(client.bottom,1L));auto oldBitmap=SelectObject(dc,bitmap);
        FillRect(dc,&client,background);const int width=MulDiv(client.right,96,static_cast<int>(dpi)),height=MulDiv(client.bottom,96,static_cast<int>(dpi)),middle=width*56/100;
        auto card=[&](int x,int y,int w,int h){RECT rect{px(x),px(y),px(x+w),px(y+h)};rounded(dc,rect,RGB(255,255,255),RGB(221,228,233),9);};
        card(16,76,middle-24,height-156);card(middle+8,76,width-middle-24,height-156);
        RECT toggle{px(width-318),px(13),px(width-14),px(51)};rounded(dc,toggle,RGB(237,242,245),RGB(222,229,235),7);
        auto line=[&](int x,int y,int x2,int y2){HPEN pen=CreatePen(PS_SOLID,1,RGB(224,231,236));auto old=SelectObject(dc,pen);MoveToEx(dc,px(x),px(y),nullptr);LineTo(dc,px(x2),px(y2));SelectObject(dc,old);DeleteObject(pen);};
        line(0,62,width,62);line(middle+24,274,width-32,274);line(32,height-126,middle-24,height-126);line(0,height-72,width,height-72);
        HPEN dashed=CreatePen(PS_DASH,1,RGB(141,194,201));auto oldPen=SelectObject(dc,dashed);auto oldBrush=SelectObject(dc,GetStockObject(NULL_BRUSH));
        RoundRect(dc,px(32),px(124),px(middle-24),px(180),px(10),px(10));SelectObject(dc,oldPen);SelectObject(dc,oldBrush);DeleteObject(dashed);
        BitBlt(target,0,0,client.right,client.bottom,dc,0,0,SRCCOPY);SelectObject(dc,oldBitmap);DeleteObject(bitmap);DeleteDC(dc);
        if(!destination)EndPaint(window,&painting);
    }
};

static LRESULT CALLBACK windowProc(HWND window,UINT message,WPARAM wparam,LPARAM lparam) {
    auto app=reinterpret_cast<App*>(GetWindowLongPtrW(window,GWLP_USERDATA));
    if(message==WM_NCCREATE){app=reinterpret_cast<App*>(reinterpret_cast<CREATESTRUCTW*>(lparam)->lpCreateParams);app->window=window;SetWindowLongPtrW(window,GWLP_USERDATA,reinterpret_cast<LONG_PTR>(app));}
    if(!app)return DefWindowProcW(window,message,wparam,lparam);
    try {
        switch(message) {
        case WM_CREATE:app->create();return 0;
        case WM_SIZE:app->layout();return 0;
        case WM_DPICHANGED:{app->dpi=HIWORD(wparam);app->createFonts();const auto rect=reinterpret_cast<RECT*>(lparam);SetWindowPos(window,nullptr,rect->left,rect->top,rect->right-rect->left,rect->bottom-rect->top,SWP_NOZORDER|SWP_NOACTIVATE);app->layout();return 0;}
        case WM_GETMINMAXINFO:{auto info=reinterpret_cast<MINMAXINFO*>(lparam);RECT rect{0,0,app->px(800),app->px(550)};AdjustWindowRectExForDpi(&rect,WS_OVERLAPPEDWINDOW,FALSE,WS_EX_CONTROLPARENT,app->dpi);info->ptMinTrackSize={rect.right-rect.left,rect.bottom-rect.top};return 0;}
        case WM_PAINT:app->paint();return 0;
        case WM_PRINTCLIENT:app->paint(reinterpret_cast<HDC>(wparam));return 0;
        case WM_CTLCOLORSTATIC:case WM_CTLCOLOREDIT:{HDC dc=reinterpret_cast<HDC>(wparam);const int id=GetDlgCtrlID(reinterpret_cast<HWND>(lparam));
            const bool outer=id==Title||id==Forward||id==Reverse||id==Status||id==Elapsed||id==Environment;
            const bool muted=id==StyleNote||id==OutputNote||id==Existing||id==Environment;
            SetTextColor(dc,IsWindowEnabled(reinterpret_cast<HWND>(lparam))?(muted?MutedTextColor:RGB(31,43,55)):DisabledTextColor);SetBkColor(dc,outer?RGB(246,248,250):RGB(255,255,255));
            return reinterpret_cast<LRESULT>(outer?app->background:app->white);}
        case WM_DRAWITEM:app->drawItem(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam));return TRUE;
        case WM_ERASEBKGND:return TRUE;
        case WM_COMMAND:{const int id=LOWORD(wparam),notification=HIWORD(wparam);
            if((id==Forward||id==Reverse)&&notification==BN_CLICKED)app->setDirection(id==Reverse);
            else if((id==Bracket||id==Dollar)&&notification==BN_CLICKED)app->refresh();
            else if(id==Choose&&notification==BN_CLICKED)app->selectFile();
            else if(id==AddFolder&&notification==BN_CLICKED)app->selectDirectory(true);
            else if(id==Advanced&&notification==BN_CLICKED)app->advancedSettings();
            else if(id==ClearQueue&&notification==BN_CLICKED&&!app->batchRunning){app->documents.clear();ListView_DeleteAllItems(app->control(QueueList));app->input.clear();app->summary=false;app->batchStarted=app->batchElapsed=0;app->updateElapsed();app->redrawCounts();app->refresh();}
            else if(id==ChooseDirectory&&notification==BN_CLICKED)app->selectDirectory();
            else if(id==Start&&notification==BN_CLICKED)app->startConversion();else if(id==Stop&&notification==BN_CLICKED)app->stop();
            else if(((id==ModeBox||id==FollowSize||id==AddNumbers)&&notification==CBN_SELCHANGE)||((id==OriginalDirectory||id==CustomDirectory)&&notification==BN_CLICKED)){if(id==FollowSize)app->layout();app->refresh();}
            else if(((id==ModeBox||id==FollowSize||id==AddNumbers)&&notification==CBN_EDITCHANGE)||((id==SizeBox||id==NameBox)&&notification==EN_CHANGE))app->refresh();
            else if(id==OpenReport&&notification==BN_CLICKED){int index=ListView_GetNextItem(app->control(QueueList),-1,LVNI_SELECTED);if(index<0&&!app->documents.empty())index=0;app->showDetails(index);}
            else if(notification==BN_CLICKED&&(id==OpenOutput||id==OpenFolder)){const auto path=id==OpenOutput?app->output:parentPath(app->output);
                require(reinterpret_cast<INT_PTR>(ShellExecuteW(window,L"open",path.c_str(),nullptr,nullptr,SW_SHOWNORMAL))>32,"Open verified result");}return 0;}
        case WM_NOTIFY:{auto header=reinterpret_cast<NMHDR*>(lparam);
            if(header->idFrom==QueueList){
                if(header->code==NM_CLICK){const auto click=reinterpret_cast<NMITEMACTIVATE*>(lparam);if(click->iSubItem==3)app->removeDocument(click->iItem);else if(click->iSubItem==4)app->showDetails(click->iItem);}
                if(header->code==NM_DBLCLK){const auto click=reinterpret_cast<NMITEMACTIVATE*>(lparam);if(click->iSubItem<3)app->openDocument(click->iItem);}
                if(header->code==LVN_KEYDOWN){const auto key=reinterpret_cast<NMLVKEYDOWN*>(lparam);if(key->wVKey==VK_DELETE){for(int i=static_cast<int>(app->documents.size())-1;i>=0;--i)if(ListView_GetItemState(app->control(QueueList),i,LVIS_SELECTED))app->removeDocument(i);}else if(key->wVKey==VK_RETURN)app->openDocument(ListView_GetNextItem(app->control(QueueList),-1,LVNI_SELECTED));}
                if(header->code==LVN_GETINFOTIPW){auto tip=reinterpret_cast<NMLVGETINFOTIPW*>(lparam);if(tip->iItem>=0&&static_cast<size_t>(tip->iItem)<app->documents.size()){const auto& item=app->documents[static_cast<size_t>(tip->iItem)];const auto text=item.path+(item.error.empty()?L"":L"\r\n"+item.error);wcsncpy_s(tip->pszText,static_cast<size_t>(tip->cchTextMax),text.c_str(),_TRUNCATE);}return 0;}
                if(header->code==NM_CUSTOMDRAW){auto draw=reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);if(draw->nmcd.dwDrawStage==CDDS_PREPAINT)return CDRF_NOTIFYITEMDRAW;if(draw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT)return CDRF_NOTIFYSUBITEMDRAW;
                    if(draw->nmcd.dwDrawStage==(CDDS_ITEMPREPAINT|CDDS_SUBITEM)&&draw->iSubItem==4){draw->clrText=RGB(9,126,132);return CDRF_NEWFONT;}
                    if(draw->nmcd.dwDrawStage==(CDDS_ITEMPREPAINT|CDDS_SUBITEM)&&draw->iSubItem==2&&draw->nmcd.dwItemSpec<app->documents.size()){const auto state=app->documents[draw->nmcd.dwItemSpec].state;draw->clrText=state==DocumentState::Failed?RGB(184,63,63):state==DocumentState::Partial?RGB(163,113,35):state==DocumentState::Complete?RGB(17,129,124):RGB(104,120,133);return CDRF_NEWFONT;}}
            }
            if(header->idFrom==SizeSpin&&header->code==UDN_DELTAPOS){const auto change=reinterpret_cast<NMUPDOWN*>(lparam);double size=wcstod(windowText(app->control(SizeBox)).c_str(),nullptr);
                if(!std::isfinite(size))size=12;size=std::clamp(size+change->iDelta,1.0,256.0);wchar_t value[40]{};swprintf_s(value,L"%.2f",size);app->set(SizeBox,value);return TRUE;}return 0;}
        case WM_DROPFILES:{HDROP drop=reinterpret_cast<HDROP>(wparam);std::vector<std::wstring> paths;const UINT count=DragQueryFileW(drop,0xFFFFFFFF,nullptr,0);
            for(UINT i=0;i<count;++i){const auto length=DragQueryFileW(drop,i,nullptr,0);std::wstring path(static_cast<size_t>(length)+1,L'\0');DragQueryFileW(drop,i,path.data(),length+1);path.resize(length);paths.push_back(std::move(path));}DragFinish(drop);
            if(app->batchRunning)app->set(Status,L"当前批次完成后可添加文档。");else app->addInputs(paths);return 0;}
        case WM_CONTEXTMENU:if(reinterpret_cast<HWND>(wparam)==app->control(QueueList)){app->queueMenu(lparam);return 0;}break;
        case WorkerLine:{std::unique_ptr<std::string> line(reinterpret_cast<std::string*>(lparam));app->line(*line);return 0;}
        case WorkerLog:{std::unique_ptr<std::string> log(reinterpret_cast<std::string*>(lparam));app->documents.at(app->batchOrder.at(app->batchPosition)).trace=std::move(*log);return 0;}
        case WorkerDone:app->done(static_cast<DWORD>(wparam));return 0;
        case BeginInput:
            if(app->batchTest){app->addInputs(app->batchTestInputs);app->set(NameBox,app->reverseMode?L"_LaTeX":L"_MathType");app->saveNewBatchReady();}
            else app->chooseInput(app->input);return 0;
        case NextDocument:app->nextDocument();return 0;
        case WM_TIMER:

            if(wparam==2){KillTimer(window,2);SendMessageW(window,WM_COMMAND,MAKEWPARAM(Start,BN_CLICKED),reinterpret_cast<LPARAM>(app->control(Start)));}
            if(wparam==3){KillTimer(window,3);DestroyWindow(window);}
            if(wparam==4)app->updateElapsed();return 0;
        case WM_CLOSE:app->saveSettings();if(app->batchRunning){app->closePending=true;app->stop();app->set(Status,L"正在结束当前任务并清理…");}else DestroyWindow(window);return 0;
        case WM_DESTROY:PostQuitMessage(app->exitCode);return 0;
        }
    } catch(const std::exception& error) {
        app->error=fromUtf8(error.what());app->set(Status,L"操作未完成："+app->error);
        if(!app->testDirectory.empty()){
            // 开发验收保留 GUI 自身的失败原因，不能只记录子工作进程而丢失控件检查错误。
            try{saveNew(app->testDirectory+L"\\gui-error.json","{\"Error\":"+quote(app->error)+"}\n");}catch(...){}
            app->exitCode=1;if(app->busy){app->closePending=true;app->stop();}else DestroyWindow(window);
        }
        else if(message==WorkerLine)app->workerFailed=true;
        if(message==WM_CREATE)return -1;
    }
    return DefWindowProcW(window,message,wparam,lparam);
}

static int scanWorker(int argc,wchar_t** argv) {
    if(argc!=4&&argc!=5&&argc!=6)return 2;if(argc>=5&&std::wstring(argv[4])!=L"reverse"&&std::wstring(argv[4])!=L"forward")return 2;
    setvbuf(stdout,nullptr,_IONBF,0);const auto initialized=CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED);if(FAILED(initialized))return 1;
    int code=1;
    try {
        const auto result=scanReadOnlyFile(absolutePath(argv[2]),absolutePath(argv[3]),argc>=5&&std::wstring(argv[4])==L"reverse"?scanReverseDocument:scanDocument);
        emitGuiScan(result);code=0;
    } catch(const std::exception& error){std::printf("UI_ERROR %s\n",wireHex(error.what()).c_str());}
    CoUninitialize();return code;
}
int WINAPI wWinMain(HINSTANCE instance,HINSTANCE,PWSTR,int show) {
    int argc{};wchar_t** argv=CommandLineToArgvW(GetCommandLineW(),&argc);if(!argv)return 1;
    if(argc>3&&std::wstring(argv[1])==L"--task-run") {
        int code=1;try {TaskMapping mapping(argv[2],false);taskShared=mapping.state;
            code=std::wstring(argv[3])==L"--worker-recover"&&argc==5?recoverPreferencesMain(argv[4]):std::wstring(argv[3])==L"--worker-scan"?scanWorker(argc-2,argv+2):convertMain(argc-2,argv+2);
            taskShared=nullptr;
        }catch(const std::exception& error){taskError(error.what());}LocalFree(argv);return code;
    }
    if(argc>1&&(std::wstring(argv[1])==L"--worker-scan"||std::wstring(argv[1])==L"--worker-convert"||std::wstring(argv[1])==L"--worker-reverse"||std::wstring(argv[1])==L"--worker-recover")){const int code=superviseTask(argc,argv);LocalFree(argv);return code;}
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    const auto initialized=CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED);if(FAILED(initialized)){LocalFree(argv);return 1;}
    int code=1;
    try {
        // 图形服务只在主界面启动，工作进程不加载界面绘制资源。
        struct GraphicsSession {
            ULONG_PTR token{};
            GraphicsSession(){Gdiplus::GdiplusStartupInput input;require(Gdiplus::GdiplusStartup(&token,&input,nullptr)==Gdiplus::Ok,"Start native graphics");}
            ~GraphicsSession(){Gdiplus::GdiplusShutdown(token);}
        } graphics;
        App app;app.instance=instance;wchar_t executable[32768]{};require(GetModuleFileNameW(nullptr,executable,32768)>0,"GUI executable path");app.executable=executable;
        if(argc==5&&std::wstring(argv[1])==L"--ui-test") {
            app.input=absolutePath(argv[2]);app.testDirectory=absolutePath(argv[3]);app.testMode=argv[4];
            app.reverseMode=app.testMode.rfind(L"reverse-",0)==0;
            if(app.testMode==L"reverse-bracket"){app.reverseMode=false;app.testStep=-1;}
            require(app.testMode==L"normal"||app.testMode==L"stop"||app.testMode==L"reverse-bracket"||app.testMode==L"reverse-dollar"||app.testMode==L"reverse-stop"||app.testMode==L"preview","Known GUI test mode");
        }
        else if(argc>=5&&std::wstring(argv[1])==L"--ui-batch-test"){
            app.batchTest=true;app.testDirectory=absolutePath(argv[2]);app.testMode=L"batch-"+std::wstring(argv[3]);app.reverseMode=std::wstring(argv[3])==L"reverse";
            for(int i=4;i<argc;++i)app.batchTestInputs.push_back(absolutePath(argv[i]));app.input=app.batchTestInputs.front();
        }else require(argc==1,"Launch GUI without arguments, or use its documented worker/test entry");
        INITCOMMONCONTROLSEX controls{sizeof(controls),ICC_PROGRESS_CLASS|ICC_UPDOWN_CLASS|ICC_LISTVIEW_CLASSES};require(InitCommonControlsEx(&controls)!=0,"Windows common controls");
        WNDCLASSEXW cls{sizeof(cls)};cls.style=CS_HREDRAW|CS_VREDRAW;cls.lpfnWndProc=windowProc;cls.hInstance=instance;cls.hCursor=LoadCursorW(nullptr,IDC_ARROW);
        cls.hIcon=LoadIconW(instance,MAKEINTRESOURCEW(1));cls.hIconSm=static_cast<HICON>(LoadImageW(instance,MAKEINTRESOURCEW(1),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_SHARED));cls.lpszClassName=L"DocxMathTypeBatch.B";require(RegisterClassExW(&cls)!=0,"Register native GUI window");
        const UINT dpi=GetDpiForSystem();RECT rect{0,0,MulDiv(824,static_cast<int>(dpi),96),MulDiv(560,static_cast<int>(dpi),96)};
        AdjustWindowRectExForDpi(&rect,WS_OVERLAPPEDWINDOW,FALSE,WS_EX_CONTROLPARENT,dpi);
        HWND window=CreateWindowExW(WS_EX_CONTROLPARENT,cls.lpszClassName,L"DocxMathTypeBatch · 文档公式转换",WS_OVERLAPPEDWINDOW|WS_CLIPCHILDREN,
            CW_USEDEFAULT,CW_USEDEFAULT,rect.right-rect.left,rect.bottom-rect.top,nullptr,nullptr,instance,&app);require(window!=nullptr,"Create B dual-column workbench");
        ShowWindow(window,show);UpdateWindow(window);if(!app.input.empty())PostMessageW(window,BeginInput,0,0);
        MSG message{};int retrieved{};
        while((retrieved=GetMessageW(&message,nullptr,0,0))>0)if(!IsDialogMessageW(window,&message)){TranslateMessage(&message);DispatchMessageW(&message);}
        require(retrieved==0,"GUI message loop");code=static_cast<int>(message.wParam);
    }catch(const std::exception& error){MessageBoxW(nullptr,fromUtf8(error.what()).c_str(),L"DocxMathTypeBatch",MB_OK|MB_ICONERROR);}
    CoUninitialize();LocalFree(argv);return code;
}
