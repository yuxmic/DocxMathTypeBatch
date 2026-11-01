// B 双栏原生界面：同一 EXE 的独立 STA 工作进程执行扫描与转换，UI 不持有 Word 对象。
#include "scan_file.h"
#include "body_conversion.h"
#include "worker_client.h"
#include "gui_acceptance.h"
#include "mathtype_api.h"
#include "experimental_conversion.h"
#include "agent_interface.h"
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <cmath>
#include <set>
#include "layout_options.h"
#include <gdiplus.h>
#include <d2d1.h>
#include <dwrite.h>
// 启用系统新版控件外观；绘制仍使用原生 Windows API。
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

int convertMain(int argc, wchar_t** argv);
int recoverPreferencesMain(const std::wstring& originalName);
constexpr UINT BeginInput = WM_APP + 3, NextDocument = WM_APP + 4;
// 辅助文字和禁用文字分别使用统一色值，避免系统灰字与正文混淆。
constexpr COLORREF MutedTextColor=RGB(145,158,169),DisabledTextColor=RGB(178,188,197);

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

static std::wstring trim(std::wstring value) {
    const auto first = value.find_first_not_of(L" \t\r\n");
    return first == std::wstring::npos ? L"" : value.substr(first, value.find_last_not_of(L" \t\r\n") - first + 1);
}
static std::wstring parentPath(const std::wstring& path) { return path.substr(0, path.find_last_of(L"\\/")); }
static std::wstring fileName(const std::wstring& path) { return path.substr(path.find_last_of(L"\\/") + 1); }

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
// 环境检测只读注册表，避免 CLSIDFromProgID 对未知类产生注册信息。
static bool registeredClass(const wchar_t* progId) {
    wchar_t value[64]{};DWORD bytes=sizeof(value),type{};CLSID cls{};
    const auto key=std::wstring(progId)+L"\\CLSID";
    return RegGetValueW(HKEY_CLASSES_ROOT,key.c_str(),nullptr,RRF_RT_REG_SZ,&type,value,&bytes)==ERROR_SUCCESS&&SUCCEEDED(CLSIDFromString(value,&cls));
}
struct App {
    ComPtr<ID2D1Factory> drawingFactory;ComPtr<ID2D1DCRenderTarget> textTarget;ComPtr<IDWriteFactory> textFactory;
    ComPtr<IDWriteTextFormat> bodyText,headingText,titleText,helpText;
    HWND window{},fastTooltip{}; HINSTANCE instance{}; UINT dpi{96}; HFONT body{}, heading{}, title{}; HBRUSH white{}, background{};HICON documentIcon{};
    std::wstring fastHelpText{L"舍弃版式检查，提高转换速度，但存在错位风险"};
    std::wstring executable, input, task, stopFile, output, resultReport;
    std::wstring install, mathInstallError;
    WorkerClient worker;WorkerState run;
    GuiAcceptance acceptance;
    bool reverseMode{},progressMarquee{},closePending{},wordDetected{},mathDetected{};
    int exitCode{};
    std::vector<QueuedDocument> documents;
    std::vector<size_t> batchOrder;
    size_t batchPosition{};
    bool batchRunning{}, batchStopping{};
    ULONGLONG batchStarted{}, batchElapsed{};
    EquationSpacing inlineSpacing,displaySpacing;
    EquationNumbering numbering;
    ExperimentalOptions experimental;
    bool convertOmml{true};
    bool fastConversion{};
    float fastPosition{},fastStart{},fastTarget{};ULONGLONG fastStarted{};
    bool chapterNumber{},sectionNumber{};
    int chapterStyle{},sectionStyle{},equationStyle{},numberEnclosure{};
    std::wstring numberSeparator{L"."};
    HIMAGELIST queueImages{};
    HWND hotControl{},pressedControl{};int hotHalf{};
    std::wstring interfaceFont{L"Microsoft YaHei UI"};
    ~App() {
        worker.join(); cleanupTask();
        for (auto font : {body, heading, title}) if (font) DeleteObject(font);
        if (white) DeleteObject(white); if (background) DeleteObject(background);
        if(documentIcon)DestroyIcon(documentIcon);
        if(queueImages)ImageList_Destroy(queueImages);
    }
    HWND control(int id) const { return GetDlgItem(window, id); }
    void set(int id, const std::wstring& value) const { if(windowText(control(id))!=value)SetWindowTextW(control(id), value.c_str()); }
    void enable(int id,bool enabled) const {if((IsWindowEnabled(control(id))!=FALSE)!=enabled)EnableWindow(control(id),enabled);}
    void show(int id,bool visible) const {if(((GetWindowLongPtrW(control(id),GWL_STYLE)&WS_VISIBLE)!=0)!=visible)ShowWindow(control(id),visible?SW_SHOW:SW_HIDE);}
    void updateFastTransition(HWND child) {
        if(!fastStarted)return;
        const float time=std::min(static_cast<float>(GetTickCount64()-fastStarted)/180.0f,1.0f),remaining=1-time;
        fastPosition=fastStart+(fastTarget-fastStart)*(1-remaining*remaining*remaining);
        if(time>=1){fastStarted=0;KillTimer(child,92);}
    }
    void startFastTransition(HWND child) {
        const float target=SendMessageW(child,BM_GETCHECK,0,0)==BST_CHECKED?1.0f:0.0f;
        if(target==fastTarget)return;
        // 连续点击从当前画面反向过渡；系统关闭界面动画时立即显示最终状态。
        updateFastTransition(child);fastStart=fastPosition;fastTarget=target;
        BOOL animate=TRUE;SystemParametersInfoW(SPI_GETCLIENTAREAANIMATION,0,&animate,0);
        if(animate){fastStarted=GetTickCount64();SetTimer(child,92,15,nullptr);}
        else{fastPosition=target;fastStarted=0;KillTimer(child,92);}
        InvalidateRect(child,nullptr,FALSE);
    }
    void animateProgress(bool active) {
        if(progressMarquee==active)return;
        auto bar=control(Progress);if(!active)SendMessageW(bar,PBM_SETMARQUEE,FALSE,0);
        const auto style=GetWindowLongPtrW(bar,GWL_STYLE);SetWindowLongPtrW(bar,GWL_STYLE,active?style|PBS_MARQUEE:style&~PBS_MARQUEE);
        if(active)SendMessageW(bar,PBM_SETMARQUEE,TRUE,40);progressMarquee=active;
    }
    void countProgress(size_t current,size_t total) {
        animateProgress(false);SendMessageW(control(Progress),PBM_SETRANGE32,0,static_cast<LPARAM>(total));SendMessageW(control(Progress),PBM_SETPOS,current,0);
        acceptance.progress(window,current);
    }
    int px(int value) const { return MulDiv(value, static_cast<int>(dpi), 96); }
    bool cleanupTask() {
        // 临时目录由本实例唯一创建，仅清理登记的任务文件，不递归删除或枚举用户目录。
        if (task.empty()) return true;
        bool clean=true;
        for (const auto& path : {stopFile, resultReport,stopFile+L".commit"}) if (!path.empty() && !DeleteFileW(path.c_str()) && GetLastError()!=ERROR_FILE_NOT_FOUND) clean=false;
        if(!RemoveDirectoryW(task.c_str())&&GetLastError()!=ERROR_PATH_NOT_FOUND)clean=false;
        if(clean){task.clear(); stopFile.clear(); resultReport.clear();}return clean;
    }
    void newTask() {
        require(cleanupTask(),"Previous GUI task files are removed"); wchar_t temporary[32768]{}; require(GetTempPathW(32768, temporary) > 0, "Windows task temporary directory");
        GUID id{}; check(CoCreateGuid(&id), "Create GUI task identity"); wchar_t guid[40]{}; StringFromGUID2(id, guid, 40);
        const auto path = std::wstring(temporary) + L"DocxMathTypeBatch-" + guid;
        require(CreateDirectoryW(path.c_str(), nullptr) != 0, "Create unique GUI task directory"); task = path;
        resultReport = task + L"\\result.json"; stopFile = task + L"\\stop.flag";
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
        if(fastTooltip){SendMessageW(fastTooltip,WM_SETFONT,reinterpret_cast<WPARAM>(nextBody),TRUE);SendMessageW(fastTooltip,TTM_SETMAXTIPWIDTH,0,px(360));}
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
        format(10,DWRITE_FONT_WEIGHT_NORMAL,helpText);
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
        add(StyleHeading,L"STATIC",L"转换设置");add(ModeLabel,L"STATIC",L"正文范围");
        add(ModeBox,L"COMBOBOX",L"",CBS_DROPDOWNLIST|WS_VSCROLL|WS_TABSTOP);
        for(const auto* mode:{L"所有正文",L"仅行内",L"仅行外"})SendMessageW(control(ModeBox),CB_ADDSTRING,0,reinterpret_cast<LPARAM>(mode));
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
        add(FastLabel,L"STATIC",L"快速转换");
        add(FastConversion,L"BUTTON",L"快速转换（跳过核验）",BS_AUTOCHECKBOX|WS_TABSTOP);
        add(FastHelp,L"BUTTON",fastHelpText.c_str(),BS_OWNERDRAW|WS_TABSTOP);
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
        for(int id:{Forward,Reverse,FastConversion,FastHelp,Choose,AddFolder,ClearQueue,Advanced,ChooseDirectory,Start,Stop,ModeBox,FollowSize,AddNumbers,SizeSpin,Progress,Bracket,Dollar,QueueList})
            SetWindowSubclass(control(id),paintControl,1,reinterpret_cast<DWORD_PTR>(this));
        // 由系统管理悬停显示和离开隐藏；说明仅关联问号，不扩大快速开关的点击区域。
        fastTooltip=CreateWindowExW(WS_EX_TOPMOST,TOOLTIPS_CLASSW,nullptr,WS_POPUP|TTS_ALWAYSTIP|TTS_NOPREFIX,
            CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,window,nullptr,instance,nullptr);
        require(fastTooltip!=nullptr,"Create fast conversion tooltip");
        TOOLINFOW help{sizeof(help)};help.hwnd=window;help.uFlags=TTF_IDISHWND|TTF_SUBCLASS;
        help.uId=reinterpret_cast<UINT_PTR>(control(FastHelp));help.lpszText=fastHelpText.data();
        require(SendMessageW(fastTooltip,TTM_ADDTOOLW,0,reinterpret_cast<LPARAM>(&help))!=FALSE,"Attach fast conversion tooltip");
        SendMessageW(fastTooltip,WM_SETFONT,reinterpret_cast<WPARAM>(body),FALSE);SendMessageW(fastTooltip,TTM_SETMAXTIPWIDTH,0,px(360));
        SHFILEINFOW fileInfo{};if(SHGetFileInfoW(L".docx",FILE_ATTRIBUTE_NORMAL,&fileInfo,sizeof(fileInfo),SHGFI_ICON|SHGFI_LARGEICON|SHGFI_USEFILEATTRIBUTES))documentIcon=fileInfo.hIcon;
        queueImages=ImageList_Create(px(24),px(32),ILC_COLOR32|ILC_MASK,1,1);require(queueImages!=nullptr,"Create document queue icons");
        if(documentIcon)ImageList_AddIcon(queueImages,documentIcon);ListView_SetImageList(control(QueueList),queueImages,LVSIL_SMALL);
        SetWindowTheme(control(QueueList),L"Explorer",nullptr);
        SetWindowSubclass(ListView_GetHeader(control(QueueList)),paintQueueHeader,1,reinterpret_cast<DWORD_PTR>(this));
        for(int id:{OriginalDirectory,CustomDirectory})SetWindowSubclass(control(id),paintControl,1,reinterpret_cast<DWORD_PTR>(this));
        wordDetected=registeredClass(L"Word.Application");
        install=findMathTypeInstall(mathInstallError);mathDetected=!install.empty();
        if(!acceptance.active())loadSettings();if(acceptance.mode==L"reverse-dollar")CheckRadioButton(window,Bracket,Dollar,Dollar);
        updateDirectionControls();DragAcceptFiles(window,TRUE);DragAcceptFiles(control(QueueList),TRUE);layout();refresh();
    }
    bool validDelimiter() const {
        return (SendMessageW(control(Bracket),BM_GETCHECK,0,0)==BST_CHECKED) != (SendMessageW(control(Dollar),BM_GETCHECK,0,0)==BST_CHECKED);
    }
    void updateDirectionControls() {
        for (int id : {ModeLabel,ModeBox,SizeLabel,SizeBox,SizeSpin,FollowSize,AddNumbers,NumberLabel}) show(id,!reverseMode);
        for (int id : {InlineDelimiterLabel,InlineDelimiterNote,DisplayDelimiterLabel,Bracket,Dollar}) show(id,reverseMode);
        set(StyleHeading,reverseMode?L"导出格式":L"转换设置");
        show(Advanced,true);
    }
    void setDirection(bool reverse) {
        if(batchRunning||reverseMode==reverse)return;reverseMode=reverse;updateDirectionControls();
        set(NameBox,reverse?L"_LaTeX":L"_MathType");refresh();
        if(acceptance.interactive()&&!input.empty()&&acceptance.step==0)acceptance.advance(acceptanceView());
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
        if (acceptance.active()) return;
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
        move(SizeLabel,right,170,74,34);move(FollowSize,right+82,170,rw-82-(fixed?100:0),200);move(SizeBox,width-120,170,64,34);move(SizeSpin,width-56,170,24,34);
        move(NumberLabel,right,210,74,34);move(AddNumbers,right+82,210,rw-82,200);
        move(InlineDelimiterLabel,right,132,88,30);move(InlineDelimiterNote,right+90,132,rw-90,30);
        move(DisplayDelimiterLabel,right,172,88,30);move(Bracket,right+90,170,88,32);move(Dollar,right+182,170,80,32);
        // 收紧设置行与输出区间距，为最后一行的快速开关留出空间。
        // 标签与滑块分别占用独立控件，点击文字或同行空白不会切换状态。
        move(FastLabel,right,250,52,30);move(FastHelp,right+56,258,16,16);move(FastConversion,right+82,254,40,22);
        move(OutputHeading,right,294,rw,24);move(DirectoryLabel,right,328,74,28);
        move(OriginalDirectory,right+82,324,90,30);move(CustomDirectory,right+180,324,90,30);
        move(DirectoryBox,right+82,360,rw-142,34);move(ChooseDirectory,width-84,360,52,34);
        move(NameLabel,right,404,78,32);move(NameBox,right+82,402,rw-82,34);move(OutputNote,right,442,rw,22);
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
        acceptance.refresh();
        std::wstring why;double points{};const bool validStyle=reverseMode?validDelimiter():style(points,why);
        const bool original=SendMessageW(control(OriginalDirectory),BM_GETCHECK,0,0)==BST_CHECKED;
        const auto attributes=GetFileAttributesW(windowText(control(DirectoryBox)).c_str());
        const auto suffix=windowText(control(NameBox));
        const bool validOutput=legalName(L"document"+suffix+L".docx")&&(original||(attributes!=INVALID_FILE_ATTRIBUTES&&(attributes&FILE_ATTRIBUTE_DIRECTORY)));
        enable(Start,!batchRunning&&!documents.empty()&&wordDetected&&mathDetected&&validStyle&&validOutput);
        for(int id:{Choose,AddFolder,ClearQueue,ModeBox,FollowSize,NameBox,Forward,Reverse,FastConversion,Bracket,Dollar,OriginalDirectory,CustomDirectory,Advanced})enable(id,!batchRunning);
        enable(ChooseDirectory,!batchRunning&&!original);enable(DirectoryBox,!batchRunning&&!original);
        const bool fixed=SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==1;
        show(SizeBox,!reverseMode&&fixed);show(SizeSpin,!reverseMode&&fixed);enable(SizeBox,!batchRunning&&fixed);enable(SizeSpin,!batchRunning&&fixed);
        enable(AddNumbers,!batchRunning&&SendMessageW(control(ModeBox),CB_GETCURSEL,0,0)!=1);
        show(Stop,batchRunning);enable(Stop,batchRunning&&!closePending&&!batchStopping);
        show(Progress,batchRunning);
        size_t pending{};for(const auto& item:documents)pending+=item.state==DocumentState::Pending;
        set(Start,(pending||!batchStarted?L"开始转换":L"重新转换")+(documents.empty()?std::wstring():L"（"+std::to_wstring(pending?pending:documents.size())+L"）"));
        if(!batchRunning&&!run.summary){
            if(!mathDetected||!wordDetected)set(Status,environmentProblem());else if(documents.empty())set(Status,L"就绪 · 请添加文档");
            else if(!validStyle)set(Status,reverseMode?L"请选择行外定界符":why);else if(!validOutput)set(Status,L"请设置有效的输出位置和名称");else set(Status,L"准备就绪");
        }
    }
    void redrawCounts() {
        size_t pending{},done{},errors{};for(const auto& item:documents){pending+=item.state==DocumentState::Pending;done+=item.state==DocumentState::Complete||item.state==DocumentState::Partial;errors+=item.state==DocumentState::Failed;}
        set(Existing,std::to_wstring(documents.size())+L" 个文件");set(StageHeading,L"待转换 "+std::to_wstring(pending)+L"    完成 "+std::to_wstring(done)+L"    失败 "+std::to_wstring(errors));
    }
    void launch(const std::vector<std::wstring>& arguments) {
        const bool hadCounts=run.scanned;
        worker.join();run.reset(fastConversion);
        require(SetEnvironmentVariableW(L"DOCX_MATH_GUI_COMMIT",(stopFile+L".commit").c_str())!=0,"Set current GUI publication acknowledgement");
        if(hadCounts)redrawCounts();
        animateProgress(true);refresh();worker.start(window,executable,arguments,acceptance.active());
    }
    void updateRow(size_t index) {
        const auto& document=documents.at(index);
        auto value=fileName(document.path);ListView_SetItemText(control(QueueList),static_cast<int>(index),0,value.data());
        wchar_t bytes[48]{};swprintf_s(bytes,document.bytes>=1024*1024?L"%.1f MB":L"%.0f KB",document.bytes>=1024*1024?document.bytes/1048576.0:std::max(1.0,document.bytes/1024.0));
        ListView_SetItemText(control(QueueList),static_cast<int>(index),1,bytes);
        const auto state=document.fast&&document.state==DocumentState::Complete?L"完成·快速":document.fast&&document.state==DocumentState::Partial?L"部分·快速":document.state==DocumentState::Complete&&!document.warnings.empty()?L"完成·提醒":DocumentStates[static_cast<int>(document.state)];
        ListView_SetItemText(control(QueueList),static_cast<int>(index),2,const_cast<wchar_t*>(state));
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
        run.summary=false;redrawCounts();refresh();if(skipped)set(Status,L"已跳过 "+std::to_wstring(skipped)+L" 个重复或非 DOCX 文件");
    }
    void chooseInput(const std::wstring& path) {
        addInputs({path});
        if(acceptance.interactive())acceptance.advance(acceptanceView());
    }
    void removeDocument(int index) {
        if(batchRunning||index<0||static_cast<size_t>(index)>=documents.size())return;
        documents.erase(documents.begin()+index);ListView_DeleteItem(control(QueueList),index);
        input=documents.empty()?L"":documents.front().path;run.summary=false;redrawCounts();refresh();
    }
    void updateElapsed() {
        const auto milliseconds=batchRunning?GetTickCount64()-batchStarted:batchElapsed;
        // 用系统单调时钟的差值计时，界面漏掉刷新也不会使计时变慢。
        wchar_t duration[64]{};swprintf_s(duration,L"已用时 %02llu:%02llu",milliseconds/60000,(milliseconds/1000)%60);set(Elapsed,duration);
    }

    void finishBatch() {
        batchElapsed=GetTickCount64()-batchStarted;batchRunning=false;KillTimer(window,4);updateElapsed();animateProgress(false);
        run.summary=true;redrawCounts();refresh();
        const bool hasFailures=std::any_of(documents.begin(),documents.end(),[](const QueuedDocument& item){return item.state==DocumentState::Failed;});
        set(Status,batchStopping?L"已停止":hasFailures?L"批次完成 · 双击失败文档查看详情":fastConversion?L"批次完成 · 快速输出未经核验":L"批次完成");
        acceptance.finishedBatch(window,batchElapsed,batchStopping,documents,environmentProblem());
        if(closePending)DestroyWindow(window);
    }
    void nextDocument() {
        if(!batchRunning)return;
        if(batchStopping||batchPosition>=batchOrder.size()){finishBatch();return;}
        auto& document=documents.at(batchOrder.at(batchPosition));input=document.path;
        document.fast=fastConversion;
        document.state=DocumentState::Running;document.trace.clear();document.problems.clear();document.interrupted={};document.warnings.clear();document.layoutChecked=false;document.exitCode=STILL_ACTIVE;run.lastStage=L"准备任务";document.stage.clear();
        SYSTEMTIME started{};GetLocalTime(&started);wchar_t time[40]{};swprintf_s(time,L"%04d-%02d-%02d %02d:%02d:%02d",started.wYear,started.wMonth,started.wDay,started.wHour,started.wMinute,started.wSecond);
        document.context=L"程序："+executable+L"\r\n开始时间："+time+L"\r\n方向："+(reverseMode?L"MathType → LaTeX":L"LaTeX → MathType")+L"\r\n";
        if(fastConversion)document.context+=L"模式：实验性快速转换（跳过核验，可能漏转、错位或内容错误）\r\n";
        if(!reverseMode)document.context+=L"公式模式："+windowText(control(ModeBox))+L"\r\n字号："+windowText(control(FollowSize))+(SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==1?L" "+windowText(control(SizeBox))+L" pt":L"")+L"\r\n右编号："+windowText(control(AddNumbers))+L"\r\n行内行距："+spacingArgument(inlineSpacing)+L"；行外行距："+spacingArgument(displaySpacing)+L"\r\n";
        if(!reverseMode)document.context+=L"Word 原生公式（OMML）："+std::wstring(convertOmml?L"转换":L"保留")+L"\r\n";
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
            if(experimental.enabled())document.context+=L"实验性功能："+std::wstring(experimental.tables?L"表格 ":L"")+(experimental.textboxes?L"文本框 ":L"")+(experimental.headers?L"页眉页脚 ":L"")+(experimental.notes?L"脚注尾注":L"")+L"（仅转换公式）\r\n";
            std::vector<std::wstring> arguments{reverseMode?L"--worker-reverse":L"--worker-convert",input,output,resultReport,styleArgument,stopFile,install};
            if(!reverseMode){const auto mode=SendMessageW(control(ModeBox),CB_GETCURSEL,0,0);arguments.push_back(mode==1?L"--scope=inline":mode==2?L"--scope=display":L"--scope=all");if(mode!=1&&SendMessageW(control(AddNumbers),CB_GETCURSEL,0,0)==1){
                    arguments.push_back(L"--numbers");if(numbering.custom)arguments.push_back(L"--number-custom");arguments.push_back(L"--number-format="+numbering.format);
                    arguments.push_back(L"--number-start="+std::to_wstring(numbering.start));arguments.push_back(L"--number-chapter="+std::to_wstring(numbering.chapter));arguments.push_back(L"--number-section="+std::to_wstring(numbering.section));
                }
                if(inlineSpacing.mode)arguments.push_back(L"--inline-spacing="+spacingArgument(inlineSpacing));
                if(displaySpacing.mode)arguments.push_back(L"--display-spacing="+spacingArgument(displaySpacing));}
            if(!reverseMode&&!convertOmml)arguments.push_back(L"--skip-omml");
            if(experimental.tables)arguments.push_back(L"--experimental-tables");
            if(experimental.textboxes)arguments.push_back(L"--experimental-textboxes");
            if(experimental.headers)arguments.push_back(L"--experimental-headers");
            if(experimental.notes)arguments.push_back(L"--experimental-notes");
            if(fastConversion)arguments.push_back(L"--experimental-fast");
            acceptance.batchStarted(batchPosition);
            launch(arguments);
        }catch(const std::exception& exception){
            document.state=DocumentState::Failed;document.error=failureText(fromUtf8(exception.what()));document.stage=run.lastStage;document.exitCode=1;run.busy=false;updateRow(batchOrder[batchPosition]);
            document.problems={{0,L"",L"",L"",fromUtf8(exception.what()),L"TASK_FAILED",L""}};
            if(!cleanupTask()){document.error+=L"；临时文件未清理："+task;batchStopping=true;}
            ++batchPosition;PostMessageW(window,NextDocument,0,0);
        }
    }
    void startConversion() {
        if(!IsWindowEnabled(control(Start)))return;
        batchOrder.clear();for(size_t i=0;i<documents.size();++i)if(documents[i].state==DocumentState::Pending)batchOrder.push_back(i);
        if(batchOrder.empty())for(size_t i=0;i<documents.size();++i){documents[i].state=DocumentState::Pending;documents[i].output.clear();documents[i].error.clear();documents[i].problems.clear();documents[i].interrupted={};documents[i].warnings.clear();documents[i].layoutChecked=false;documents[i].context.clear();documents[i].trace.clear();documents[i].stage.clear();documents[i].exitCode=STILL_ACTIVE;updateRow(i);batchOrder.push_back(i);}
        batchPosition=0;batchRunning=true;batchStopping=false;batchStarted=GetTickCount64();batchElapsed=0;SetTimer(window,4,1000,nullptr);updateElapsed();nextDocument();
    }
    void stop() {
        if(!batchRunning||batchStopping)return;batchStopping=true;
        if(!run.busy||run.acceptingResult){refresh();return;}run.cancelling=true;
        if(GetFileAttributesW(stopFile.c_str())==INVALID_FILE_ATTRIBUTES)saveNew(stopFile,"");
        set(Status,L"正在停止并清理当前文档…");EnableWindow(control(Stop),FALSE);
    }
    void line(const std::string& value) {
        acceptance.trace(value);
        const auto update=run.read(value);
        switch(update.event) {
        case WorkerEvent::CommitReady:
            saveNew(stopFile+L".commit","");run.acceptingResult=true;show(Stop,false);break;
        case WorkerEvent::ScanCounts:
            acceptance.scanned(run,batchPosition);
            redrawCounts();break;
        case WorkerEvent::BatchStatus:case WorkerEvent::Phase:
            set(Status,std::to_wstring(batchPosition+1)+L"/"+std::to_wstring(batchOrder.size())+L" · "+update.label);
            if(update.marquee)animateProgress(true);else if(update.event==WorkerEvent::Phase)countProgress(update.current,update.total);
            break;
        case WorkerEvent::Progress:case WorkerEvent::ConversionProgress:
            countProgress(update.current,update.total);set(Status,update.label);
            break;
        case WorkerEvent::Status:set(Status,update.label);break;
        case WorkerEvent::None:break;
        }
        if(acceptance.updated(update,window,progressMarquee,batchPosition,run.completed))stop();
    }
    void done(DWORD code) {
        worker.join();animateProgress(false);
        const bool cancelled=run.finish(code,output,resultReport);
        auto& document=documents.at(batchOrder.at(batchPosition));
        document.state=cancelled?DocumentState::Cancelled:run.workerFailed?DocumentState::Failed:run.outcome==L"PARTIAL"?DocumentState::Partial:DocumentState::Complete;
        document.error=run.error;document.stage=run.lastStage;document.exitCode=code;
        document.problems=std::move(run.problems);document.interrupted=std::move(run.interrupted);document.warnings=std::move(run.warnings);document.layoutChecked=run.layoutChecked;
        if(!run.workerFailed&&!cancelled){
            document.output=output;
            acceptance.keepReport(resultReport,batchPosition);
        }
        const auto taskDirectory=task;
        const bool clean=cleanupTask();
        if(!clean){document.state=DocumentState::Failed;document.error+=L"\r\n临时文件未清理："+task;document.problems.push_back({0,L"",L"",L"",L"任务临时文件未能清理。",L"TASK_FAILED",L""});run.workerFailed=true;batchStopping=true;}
        updateRow(batchOrder[batchPosition]);redrawCounts();
        if(cancelled)batchStopping=true;
        acceptance.finishedItem(acceptanceView(),code,taskDirectory,clean,cancelled,closePending,batchPosition);
        if(acceptance.active()&&!acceptance.batch&&run.workerFailed){exitCode=1;DestroyWindow(window);return;}
        ++batchPosition;
        if(batchStopping||batchPosition>=batchOrder.size())finishBatch();else PostMessageW(window,NextDocument,0,0);
        if(acceptance.interactive()&&!cancelled&&!run.workerFailed&&!closePending)acceptance.advance(acceptanceView());
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
        return documentDetails(documents.at(static_cast<size_t>(index)),environmentProblem());
    }
    struct DetailsView {App* app;std::wstring summary,raw;bool showRaw{};};
    static INT_PTR CALLBACK detailsProc(HWND dialog,UINT message,WPARAM wparam,LPARAM lparam) {
        auto data=reinterpret_cast<DetailsView*>(GetWindowLongPtrW(dialog,DWLP_USER));
        if(message==WM_INITDIALOG){
            data=reinterpret_cast<DetailsView*>(lparam);SetWindowLongPtrW(dialog,DWLP_USER,reinterpret_cast<LONG_PTR>(data));auto app=data->app;
            SetWindowTextW(dialog,L"转换详情");
            RECT outer{0,0,app->px(680),app->px(450)},parent{};AdjustWindowRectExForDpi(&outer,WS_POPUP|WS_CAPTION|WS_SYSMENU|DS_MODALFRAME,FALSE,WS_EX_DLGMODALFRAME,app->dpi);GetWindowRect(app->window,&parent);
            SetWindowPos(dialog,nullptr,(parent.left+parent.right-outer.right+outer.left)/2,(parent.top+parent.bottom-outer.bottom+outer.top)/2,outer.right-outer.left,outer.bottom-outer.top,SWP_NOZORDER);
            // 人读摘要自动换行；底层日志通过独立视图查看，保留原生文本选择和键盘操作。
            auto edit=CreateWindowExW(0,L"EDIT",L"",WS_CHILD|WS_VISIBLE|WS_TABSTOP|WS_BORDER|WS_VSCROLL|ES_MULTILINE|ES_AUTOVSCROLL|ES_READONLY,app->px(16),app->px(16),app->px(648),app->px(374),dialog,reinterpret_cast<HMENU>(static_cast<INT_PTR>(301)),app->instance,nullptr);
            require(edit!=nullptr,"Create conversion log view");SendMessageW(edit,WM_SETFONT,reinterpret_cast<WPARAM>(app->body),TRUE);SendMessageW(edit,EM_SETLIMITTEXT,8*1024*1024,0);SetWindowTextW(edit,data->summary.c_str());
            for(const auto& button:{std::make_pair(303,L"原始日志"),std::make_pair(302,L"复制全部"),std::make_pair(IDOK,L"关闭")}){
                auto child=CreateWindowExW(0,L"BUTTON",button.second,WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_OWNERDRAW,app->px(button.first==303?16:button.first==302?456:564),app->px(404),app->px(100),app->px(32),dialog,reinterpret_cast<HMENU>(static_cast<INT_PTR>(button.first)),app->instance,nullptr);
                require(child!=nullptr,"Create log dialog button");SendMessageW(child,WM_SETFONT,reinterpret_cast<WPARAM>(app->body),TRUE);SetWindowSubclass(child,paintControl,1,reinterpret_cast<DWORD_PTR>(app));
            }
            SetFocus(GetDlgItem(dialog,IDOK));return FALSE;
        }
        if(!data)return FALSE;
        if(message==WM_CTLCOLORSTATIC||message==WM_CTLCOLOREDIT){SetTextColor(reinterpret_cast<HDC>(wparam),RGB(31,43,55));SetBkColor(reinterpret_cast<HDC>(wparam),RGB(255,255,255));return reinterpret_cast<INT_PTR>(data->app->white);}
        if(message==WM_ERASEBKGND){RECT rect{};GetClientRect(dialog,&rect);FillRect(reinterpret_cast<HDC>(wparam),&rect,data->app->white);return TRUE;}
        if(message==WM_DRAWITEM){data->app->drawItem(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam));return TRUE;}
        if(message==WM_COMMAND){
            if(LOWORD(wparam)==303&&HIWORD(wparam)==BN_CLICKED){
                data->showRaw=!data->showRaw;SetDlgItemTextW(dialog,301,(data->showRaw?data->raw:data->summary).c_str());SetDlgItemTextW(dialog,303,data->showRaw?L"返回摘要":L"原始日志");
                SendDlgItemMessageW(dialog,301,EM_SETSEL,0,0);SendDlgItemMessageW(dialog,301,EM_SCROLLCARET,0,0);return TRUE;
            }
            if(LOWORD(wparam)==302){SendDlgItemMessageW(dialog,301,EM_SETSEL,0,-1);SendDlgItemMessageW(dialog,301,WM_COPY,0,0);return TRUE;}
            if(LOWORD(wparam)==IDOK||LOWORD(wparam)==IDCANCEL){EndDialog(dialog,0);return TRUE;}
        }
        if(message==WM_CLOSE){EndDialog(dialog,0);return TRUE;}return FALSE;
    }
    void showDetails(int index) {
        if(index<0||static_cast<size_t>(index)>=documents.size())return;
        DetailsView data{this,documentSummary(documents.at(static_cast<size_t>(index)),environmentProblem()),detailsText(index)};
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
        auto app=reinterpret_cast<App*>(GetWindowLongPtrW(dialog,DWLP_USER));
        const bool custom=IsDlgButtonChecked(dialog,221)==BST_CHECKED;
        for(int id:{222,223,224,225,226,227,228})EnableWindow(GetDlgItem(dialog,id),!custom);
        EnableWindow(GetDlgItem(dialog,229),custom);
        for(int row=0;row<2;++row){const auto mode=SendDlgItemMessageW(dialog,201+row,CB_GETCURSEL,0,0);EnableWindow(GetDlgItem(dialog,203+row),mode>=4);SetDlgItemTextW(dialog,205+row,mode==6?L"倍":L"磅");}
        try{
            const auto settings=advancedNumbering(dialog);const auto parts=numberParts(settings,settings.start);
            SetDlgItemTextW(dialog,233,(L"预览："+parts.before+parts.digits+parts.after).c_str());
            if(!custom)SetDlgItemTextW(dialog,229,settings.format.c_str());
        }catch(const std::exception&){SetDlgItemTextW(dialog,233,L"预览：请检查编号格式或数值");}
        EnableWindow(GetDlgItem(dialog,246),!app->reverseMode);
        if(app->reverseMode)for(int id:{201,202,203,204,220,221,222,223,224,225,226,227,228,229,230,231,232})EnableWindow(GetDlgItem(dialog,id),FALSE);
    }
    // 高缩放或较小工作区时滚动整个设置窗口，保留系统字号和键盘访问。
    static void scrollAdvanced(HWND dialog,int bar,int position) {
        SCROLLINFO info{sizeof(info),SIF_ALL};if(!GetScrollInfo(dialog,bar,&info))return;
        position=std::clamp(position,info.nMin,std::max(info.nMin,info.nMax-static_cast<int>(info.nPage)+1));const int delta=info.nPos-position;if(!delta)return;
        info.fMask=SIF_POS;info.nPos=position;SetScrollInfo(dialog,bar,&info,TRUE);
        for(auto child=GetWindow(dialog,GW_CHILD);child;child=GetWindow(child,GW_HWNDNEXT)){RECT rect{};GetWindowRect(child,&rect);MapWindowPoints(nullptr,dialog,reinterpret_cast<POINT*>(&rect),2);SetWindowPos(child,nullptr,rect.left+(bar==SB_HORZ?delta:0),rect.top+(bar==SB_VERT?delta:0),0,0,SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);}
        InvalidateRect(dialog,nullptr,TRUE);
    }
    static void revealAdvancedControl(HWND dialog,HWND child) {
        RECT rect{},client{};GetWindowRect(child,&rect);MapWindowPoints(nullptr,dialog,reinterpret_cast<POINT*>(&rect),2);GetClientRect(dialog,&client);
        for(const int bar:{SB_HORZ,SB_VERT}){SCROLLINFO info{sizeof(info),SIF_POS};if(!GetScrollInfo(dialog,bar,&info))continue;const int start=bar==SB_HORZ?rect.left:rect.top,end=bar==SB_HORZ?rect.right:rect.bottom,limit=bar==SB_HORZ?client.right:client.bottom;scrollAdvanced(dialog,bar,info.nPos+(start<0?start:end>limit?end-limit:0));}
    }
    static INT_PTR CALLBACK advancedProc(HWND dialog,UINT message,WPARAM wparam,LPARAM lparam) {
        auto app=reinterpret_cast<App*>(GetWindowLongPtrW(dialog,DWLP_USER));
        if(message==WM_INITDIALOG){
            app=reinterpret_cast<App*>(lparam);SetWindowLongPtrW(dialog,DWLP_USER,reinterpret_cast<LONG_PTR>(app));SetWindowTextW(dialog,L"高级设置");
            RECT outer{0,0,app->px(792),app->px(562)},parent{};AdjustWindowRectExForDpi(&outer,WS_POPUP|WS_CAPTION|WS_SYSMENU|DS_MODALFRAME,FALSE,WS_EX_DLGMODALFRAME,app->dpi);GetWindowRect(app->window,&parent);
            MONITORINFO monitor{sizeof(monitor)};require(GetMonitorInfoW(MonitorFromWindow(app->window,MONITOR_DEFAULTTONEAREST),&monitor)!=FALSE,"Advanced settings monitor work area");
            const int availableWidth=monitor.rcWork.right-monitor.rcWork.left-16,availableHeight=monitor.rcWork.bottom-monitor.rcWork.top-16;
            const int width=std::min(static_cast<int>(outer.right-outer.left),availableWidth),height=std::min(static_cast<int>(outer.bottom-outer.top),availableHeight);
            const bool scroll=width<outer.right-outer.left||height<outer.bottom-outer.top;
            if(scroll)SetWindowLongPtrW(dialog,GWL_STYLE,GetWindowLongPtrW(dialog,GWL_STYLE)|WS_HSCROLL|WS_VSCROLL);
            const int dialogX=std::clamp(static_cast<int>((parent.left+parent.right-width)/2),static_cast<int>(monitor.rcWork.left)+8,static_cast<int>(monitor.rcWork.right)-width-8),dialogY=std::clamp(static_cast<int>((parent.top+parent.bottom-height)/2),static_cast<int>(monitor.rcWork.top)+8,static_cast<int>(monitor.rcWork.bottom)-height-8);
            SetWindowPos(dialog,nullptr,dialogX,dialogY,width,height,SWP_NOZORDER|SWP_FRAMECHANGED);
            auto add=[&](int id,const wchar_t* type,const wchar_t* text,DWORD style,int x,int y,int width,int height){
                if(id==IDOK||id==IDCANCEL||id==235)style=(style&~BS_TYPEMASK)|BS_OWNERDRAW;
                if(!wcscmp(type,L"BUTTON"))style|=BS_NOTIFY;
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
            // 原生公式开关与位置选择分开；仅在确定后提交到本次会话。
            add(247,L"STATIC",L"转换内容",0,464,16,304,24);
            add(246,L"BUTTON",L"转换 Word 原生公式（OMML）",BS_AUTOCHECKBOX|WS_TABSTOP,464,48,304,32);
            CheckDlgButton(dialog,246,app->convertOmml?BST_CHECKED:BST_UNCHECKED);
            add(248,L"STATIC",app->reverseMode?L"反向导出时保留 Word 原生公式。":L"关闭后保留原生公式，仅转换 LaTeX。",0,464,88,304,24);
            add(249,L"STATIC",L"适用于正文及已启用的额外位置。",0,464,112,304,24);
            add(238,L"STATIC",L"实验性功能",0,464,164,304,24);
            const wchar_t* locations[]={L"表格中的公式",L"文本框中的公式",L"页眉与页脚中的公式",L"脚注与尾注中的公式"};
            const bool selected[]={app->experimental.tables,app->experimental.textboxes,app->experimental.headers,app->experimental.notes};
            for(int row=0;row<4;++row){add(241+row,L"BUTTON",locations[row],BS_AUTOCHECKBOX|WS_TABSTOP,464,196+row*40,304,32);CheckDlgButton(dialog,241+row,selected[row]?BST_CHECKED:BST_UNCHECKED);}
            add(235,L"BUTTON",L"恢复默认",BS_PUSHBUTTON|WS_TABSTOP,24,518,96,32);
            add(IDOK,L"BUTTON",L"确定",BS_DEFPUSHBUTTON|WS_TABSTOP,584,518,86,32);add(IDCANCEL,L"BUTTON",L"取消",BS_PUSHBUTTON|WS_TABSTOP,682,518,86,32);
            if(scroll)for(int pass=0;pass<2;++pass){RECT client{};GetClientRect(dialog,&client);for(const int bar:{SB_HORZ,SB_VERT}){SCROLLINFO info{sizeof(info),SIF_RANGE|SIF_PAGE|SIF_POS};info.nMax=app->px(bar==SB_HORZ?768:550)-1;info.nPage=bar==SB_HORZ?client.right:client.bottom;SetScrollInfo(dialog,bar,&info,TRUE);}}
            refreshAdvanced(dialog);return TRUE;
        }
        if(message==WM_ERASEBKGND){RECT rect{};GetClientRect(dialog,&rect);FillRect(reinterpret_cast<HDC>(wparam),&rect,app->white);return TRUE;}
        if(message==WM_CTLCOLORSTATIC||message==WM_CTLCOLOREDIT){auto dc=reinterpret_cast<HDC>(wparam);SetTextColor(dc,IsWindowEnabled(reinterpret_cast<HWND>(lparam))?RGB(48,65,79):DisabledTextColor);SetBkColor(dc,RGB(255,255,255));return reinterpret_cast<INT_PTR>(app->white);}
        if(message==WM_DRAWITEM){app->drawItem(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam));return TRUE;}
        if(message==WM_VSCROLL||message==WM_HSCROLL){const int bar=message==WM_VSCROLL?SB_VERT:SB_HORZ;SCROLLINFO info{sizeof(info),SIF_ALL};GetScrollInfo(dialog,bar,&info);int position=info.nPos;
            switch(LOWORD(wparam)){case SB_LINEUP:position-=app->px(32);break;case SB_LINEDOWN:position+=app->px(32);break;case SB_PAGEUP:position-=info.nPage;break;case SB_PAGEDOWN:position+=info.nPage;break;case SB_TOP:position=info.nMin;break;case SB_BOTTOM:position=info.nMax;break;case SB_THUMBTRACK:case SB_THUMBPOSITION:position=info.nTrackPos;break;default:return TRUE;}
            scrollAdvanced(dialog,bar,position);return TRUE;}
        if(message==WM_MOUSEWHEEL){SCROLLINFO info{sizeof(info),SIF_POS};GetScrollInfo(dialog,SB_VERT,&info);scrollAdvanced(dialog,SB_VERT,info.nPos-GET_WHEEL_DELTA_WPARAM(wparam)*app->px(40)/WHEEL_DELTA);return TRUE;}
        if(message==WM_COMMAND){const int id=LOWORD(wparam);
            if(!GetDlgItem(dialog,IDOK))return TRUE;
            const auto notification=HIWORD(wparam);if(lparam&&(notification==BN_SETFOCUS||notification==CBN_SETFOCUS||notification==EN_SETFOCUS))revealAdvancedControl(dialog,reinterpret_cast<HWND>(lparam));
            if(id==IDCANCEL){EndDialog(dialog,0);return TRUE;}
            if(id==235){
                for(int option=241;option<=244;++option)CheckDlgButton(dialog,option,BST_UNCHECKED);
                // 恢复默认仍在确定后关闭首页快速开关，取消不提交修改。
                SetPropW(dialog,L"ResetFastConversion",reinterpret_cast<HANDLE>(1));
                if(app->reverseMode){refreshAdvanced(dialog);return TRUE;}
                CheckDlgButton(dialog,246,BST_CHECKED);
                for(int mode:{201,202,224,225,226,227})SendDlgItemMessageW(dialog,mode,CB_SETCURSEL,0,0);
                for(int value:{203,204})SetDlgItemTextW(dialog,value,L"20");for(int value:{230,231,232})SetDlgItemTextW(dialog,value,L"1");
                CheckRadioButton(dialog,220,221,220);CheckDlgButton(dialog,222,BST_UNCHECKED);CheckDlgButton(dialog,223,BST_UNCHECKED);SetDlgItemTextW(dialog,228,L".");refreshAdvanced(dialog);return TRUE;}
            if(id==IDOK){
                try{
                    ExperimentalOptions options{IsDlgButtonChecked(dialog,241)==BST_CHECKED,IsDlgButtonChecked(dialog,242)==BST_CHECKED,IsDlgButtonChecked(dialog,243)==BST_CHECKED,IsDlgButtonChecked(dialog,244)==BST_CHECKED};
                    const bool fast=app->fastConversion&&!GetPropW(dialog,L"ResetFastConversion");
                    if(app->reverseMode){app->experimental=options;app->fastConversion=fast;EndDialog(dialog,IDOK);return TRUE;}
                    EquationSpacing spacing[2];for(int row=0;row<2;++row){auto& value=spacing[row];value.mode=static_cast<int>(SendDlgItemMessageW(dialog,201+row,CB_GETCURSEL,0,0));if(value.mode>=4)value.value=settingNumber(windowText(GetDlgItem(dialog,203+row)),value.mode==6?0.5:1,value.mode==6?10:999);}
                    const auto numbers=advancedNumbering(dialog);app->inlineSpacing=spacing[0];app->displaySpacing=spacing[1];app->numbering=numbers;
                    app->numbering.custom=IsDlgButtonChecked(dialog,221)==BST_CHECKED;app->chapterNumber=IsDlgButtonChecked(dialog,222)==BST_CHECKED;app->sectionNumber=IsDlgButtonChecked(dialog,223)==BST_CHECKED;
                    app->chapterStyle=static_cast<int>(SendDlgItemMessageW(dialog,224,CB_GETCURSEL,0,0));app->sectionStyle=static_cast<int>(SendDlgItemMessageW(dialog,225,CB_GETCURSEL,0,0));app->equationStyle=static_cast<int>(SendDlgItemMessageW(dialog,226,CB_GETCURSEL,0,0));app->numberEnclosure=static_cast<int>(SendDlgItemMessageW(dialog,227,CB_GETCURSEL,0,0));app->numberSeparator=windowText(GetDlgItem(dialog,228));
                    app->experimental=options;app->convertOmml=IsDlgButtonChecked(dialog,246)==BST_CHECKED;app->fastConversion=fast;
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
        SendMessageW(control(FastConversion),BM_SETCHECK,fastConversion?BST_CHECKED:BST_UNCHECKED,0);
    }
    GuiAcceptanceView acceptanceView() const {
        return {window,dpi,reverseMode,progressMarquee,task.empty(),run,documents,install,input,output};
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
        const bool choice=id==FastConversion||id==Bracket||id==Dollar||id==OriginalDirectory||id==CustomDirectory||(GetParent(child)!=app->window&&((id>=220&&id<=223)||(id>=241&&id<=244)||id==246));
        if(id==QueueList&&message==WM_DROPFILES)return SendMessageW(app->window,message,wparam,lparam);
        if(dropdown&&message==WM_NCCALCSIZE)return 0;
        if((dropdown||id==SizeSpin)&&message==WM_NCPAINT)return 0;
        if(message==WM_MOUSEMOVE){RECT rect{};GetClientRect(child,&rect);const int half=HIWORD(lparam)<rect.bottom/2?0:1;
            if(app->hotControl!=child||app->hotHalf!=half){app->hotControl=child;app->hotHalf=half;InvalidateRect(child,nullptr,FALSE);}
            // 开发预览固定合成鼠标状态，避免隐藏窗口的实际指针位置清除截图状态。
            if(app->acceptance.mode==L"preview"&&!dropdown&&id!=SizeSpin&&id!=FastHelp)return 0;
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
        if(id==FastConversion&&message==WM_TIMER&&wparam==92){app->updateFastTransition(child);InvalidateRect(child,nullptr,FALSE);return 0;}
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
            }else if(id==FastConversion){
                // 保留原生复选框的键盘及可访问状态，只将外观绘制成圆角开关。
                app->updateFastTransition(child);
                const bool hot=app->hotControl==child&&enabled,pressed=app->pressedControl==child&&enabled;
                // 用轻微渐变、细边和滑块投影呈现厚度，保持现有青灰配色。
                auto blend=[position=app->fastPosition](COLORREF off,COLORREF on){return RGB(
                    static_cast<int>(GetRValue(off)+static_cast<float>(GetRValue(on)-GetRValue(off))*position),
                    static_cast<int>(GetGValue(off)+static_cast<float>(GetGValue(on)-GetGValue(off))*position),
                    static_cast<int>(GetBValue(off)+static_cast<float>(GetBValue(on)-GetBValue(off))*position));};
                const auto top=!enabled?RGB(222,229,234):blend(pressed?RGB(145,162,176):hot?RGB(172,186,198):RGB(166,180,193),pressed?RGB(10,119,125):hot?RGB(25,159,164):RGB(20,150,155));
                const auto bottom=!enabled?RGB(215,223,230):blend(pressed?RGB(132,151,166):RGB(149,165,179),pressed?RGB(8,109,115):RGB(8,126,133));
                const auto border=!enabled?RGB(209,219,226):blend(RGB(141,158,172),RGB(10,121,128));
                app->rounded(dc,rect,top,border,11,bottom);
                RECT thumb{app->px(2)+static_cast<int>(static_cast<float>(rect.right-app->px(22))*app->fastPosition),app->px(2),0,rect.bottom-app->px(2)};thumb.right=thumb.left+app->px(18);
                RECT shadow=thumb;OffsetRect(&shadow,0,app->px(1));const auto shade=!enabled?RGB(198,208,217):blend(RGB(124,143,160),RGB(8,111,118));
                app->rounded(dc,shadow,shade,shade,9);
                app->rounded(dc,thumb,enabled?RGB(255,255,255):RGB(242,245,247),enabled?RGB(213,223,231):RGB(220,228,234),9,pressed?RGB(226,234,240):RGB(239,244,248));
                if(GetFocus()==child&&!(SendMessageW(child,WM_QUERYUISTATE,0,0)&UISF_HIDEFOCUS)){RECT focus=rect;InflateRect(&focus,-app->px(1),-app->px(1));app->rounded(dc,focus,CLR_INVALID,RGB(8,84,90),10);}
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
        if(id==FastConversion&&message==BM_SETCHECK)app->startFastTransition(child);
        // 原生复选框更新焦点提示时可能直接画矩形，立即用圆角开关的完整帧覆盖。
        if(id==FastConversion&&(message==WM_UPDATEUISTATE||message==WM_SETFOCUS||message==WM_KILLFOCUS)){
            RedrawWindow(child,nullptr,nullptr,RDW_INVALIDATE|RDW_UPDATENOW);return result;
        }
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
                (GetParent(child)!=app->window&&(id==208||id==212||id==236||id==248||id==249));
            auto format=id==Title?app->titleText.Get():(id==FileHeading||id==StyleHeading||id==OutputHeading||id==200||id==237||id==238||id==247)?app->headingText.Get():app->bodyText.Get();
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
        if(id==FastHelp){
            RECT icon=draw.rcItem;FillRect(draw.hDC,&icon,white);InflateRect(&icon,-px(2),-px(2));
            const bool focused=(draw.itemState&ODS_FOCUS)&&!(draw.itemState&ODS_NOFOCUSRECT),active=hotControl==draw.hwndItem||focused;
            const auto ink=active?RGB(13,128,131):RGB(104,120,133);
            rounded(draw.hDC,icon,active?RGB(239,249,249):RGB(255,255,255),ink,6);
            drawText(draw.hDC,draw.rcItem,L"?",ink,helpText.Get(),true);return;
        }
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
        acceptance.painted(destination);
        PAINTSTRUCT painting{};HDC target=destination?destination:BeginPaint(window,&painting);RECT client{};GetClientRect(window,&client);
        HDC dc=CreateCompatibleDC(target);HBITMAP bitmap=CreateCompatibleBitmap(target,std::max(client.right,1L),std::max(client.bottom,1L));auto oldBitmap=SelectObject(dc,bitmap);
        FillRect(dc,&client,background);const int width=MulDiv(client.right,96,static_cast<int>(dpi)),height=MulDiv(client.bottom,96,static_cast<int>(dpi)),middle=width*56/100;
        auto card=[&](int x,int y,int w,int h){RECT rect{px(x),px(y),px(x+w),px(y+h)};rounded(dc,rect,RGB(255,255,255),RGB(221,228,233),9);};
        card(16,76,middle-24,height-156);card(middle+8,76,width-middle-24,height-156);
        RECT toggle{px(width-318),px(13),px(width-14),px(51)};rounded(dc,toggle,RGB(237,242,245),RGB(222,229,235),7);
        auto line=[&](int x,int y,int x2,int y2){HPEN pen=CreatePen(PS_SOLID,1,RGB(224,231,236));auto old=SelectObject(dc,pen);MoveToEx(dc,px(x),px(y),nullptr);LineTo(dc,px(x2),px(y2));SelectObject(dc,old);DeleteObject(pen);};
        line(0,62,width,62);line(middle+24,284,width-32,284);line(32,height-126,middle-24,height-126);line(0,height-72,width,height-72);
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
            else if(id==FastConversion&&notification==BN_CLICKED&&!app->batchRunning){app->fastConversion=SendMessageW(app->control(FastConversion),BM_GETCHECK,0,0)==BST_CHECKED;app->startFastTransition(app->control(FastConversion));app->refresh();}
            else if((id==Bracket||id==Dollar)&&notification==BN_CLICKED)app->refresh();
            else if(id==Choose&&notification==BN_CLICKED)app->selectFile();
            else if(id==AddFolder&&notification==BN_CLICKED)app->selectDirectory(true);
            else if(id==Advanced&&notification==BN_CLICKED)app->advancedSettings();
            else if(id==ClearQueue&&notification==BN_CLICKED&&!app->batchRunning){app->documents.clear();ListView_DeleteAllItems(app->control(QueueList));app->input.clear();app->run.summary=false;app->batchStarted=app->batchElapsed=0;app->updateElapsed();app->redrawCounts();app->refresh();}
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
            if(app->acceptance.batch){app->addInputs(app->acceptance.inputs);app->set(NameBox,app->reverseMode?L"_LaTeX":L"_MathType");app->acceptance.readyBatch(app->acceptanceView());}
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
        app->run.error=fromUtf8(error.what());app->set(Status,L"操作未完成："+app->run.error);
        if(app->acceptance.active()){
            // 开发验收保留 GUI 自身的失败原因，不能只记录子工作进程而丢失控件检查错误。
            try{app->acceptance.failed(app->run.error);}catch(...){}
            app->exitCode=1;if(app->run.busy){app->closePending=true;app->stop();}else DestroyWindow(window);
        }
        else if(message==WorkerLine)app->run.workerFailed=true;
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
    if(argc>1&&std::wstring(argv[1])==L"--describe"){const int code=agentDescribe(argc,argv);LocalFree(argv);return code;}
    if(argc==2&&std::wstring(argv[1])==L"--help"){const int code=agentHelp();LocalFree(argv);return code;}
    if(argc>3&&std::wstring(argv[1])==L"--task-run") {
        int code=1;try {TaskMapping mapping(argv[2],false);taskShared=mapping.state;
            code=std::wstring(argv[3])==L"--worker-recover"&&argc==5?recoverPreferencesMain(argv[4]):std::wstring(argv[3])==L"--worker-scan"?scanWorker(argc-2,argv+2):convertMain(argc-2,argv+2);
            taskShared=nullptr;
        }catch(const std::exception& error){taskError(error.what());}LocalFree(argv);return code;
    }
    if(argc>1&&(std::wstring(argv[1])==L"--worker-scan"||std::wstring(argv[1])==L"--worker-convert"||std::wstring(argv[1])==L"--worker-reverse"||std::wstring(argv[1])==L"--worker-recover")){const int code=superviseTask(argc,argv);LocalFree(argv);return code;}
    if(argc>1&&std::wstring(argv[1])!=L"--ui-test"&&std::wstring(argv[1])!=L"--ui-batch-test"){
        std::puts("{\"Status\":\"FAILED\",\"ErrorCode\":\"INVALID_ARGUMENT\",\"Message\":\"Unknown command; use --help\"}");LocalFree(argv);return 2;
    }
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
        app.acceptance.configure(argc,argv);
        if(app.acceptance.active()){app.input=app.acceptance.inputs.front();app.reverseMode=app.acceptance.initialReverse;}
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
