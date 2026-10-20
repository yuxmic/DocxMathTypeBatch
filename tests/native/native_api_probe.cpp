// 阶段 1 的原生接口小样：仅操作自建文档，不扫描或转换用户文件。
#include "mathtype_api.h"
#include "docx_scan.h"

// Find 只定位本轮固定小样文本，不是阶段 2 的通用定界符扫描器。
static ComPtr<IDispatch> findSampleRange(IDispatch* document, const wchar_t* source) {
    auto range = get(document, L"Content");
    auto find = get(range.Get(), L"Find");
    invoke(find.Get(), L"ClearFormatting", DISPATCH_METHOD);
    // Word Find 即使关闭通配符仍把 ^ 当特殊字符前缀；LaTeX 上标必须按 ^^ 搜索。
    std::wstring needle;
    for (const wchar_t* character = source; *character; ++character) {
        needle += *character;
        if (*character == L'^') needle += L'^';
    }
    invoke(find.Get(), L"Text", DISPATCH_PROPERTYPUT, {needle.c_str()});
    invoke(find.Get(), L"Forward", DISPATCH_PROPERTYPUT, {true});
    invoke(find.Get(), L"Wrap", DISPATCH_PROPERTYPUT, {0L});
    invoke(find.Get(), L"MatchWildcards", DISPATCH_PROPERTYPUT, {false});
    invoke(find.Get(), L"MatchWholeWord", DISPATCH_PROPERTYPUT, {false});
    invoke(find.Get(), L"Format", DISPATCH_PROPERTYPUT, {false});
    const auto found = invoke(find.Get(), L"Execute", DISPATCH_METHOD);
    require(found.v.vt == VT_BOOL && found.v.boolVal != VARIANT_FALSE, "Find fixed sample source");
    require(text(range.Get(), L"Text") == source, "Source range before insertion");
    return range;
}

struct ProbeCase {
    const char* name;
    const wchar_t* source;
    const wchar_t* tex;
    const wchar_t* eqp;
    const char* font;
    const char* points;
    const char* expected;
    bool inlineEquation;
    long tableRow{};
};

// 临时对象只在已确认归属的 Word 实例里创建，不保存文档。
// 无 item 时读取空对象偏好；有 item 时从空对象生成规范原生内容，识别重复插入。
static std::vector<unsigned char> scratchNative(MathTypeApi& api, WordSession& word,
                                                const ProbeCase* item, const std::wstring& directory) {
    auto scratch = asObject(invoke(get(word.application.Get(), L"Documents").Get(), L"Add", DISPATCH_METHOD));
    ComPtr<IDispatch> object;
    bool active = false;
    try {
        if (item) {
            const auto prefs = loadPrefs(api, directory + L"\\" + item->eqp);
            require(api.setPrefs(static_cast<short>(1 + (item->inlineEquation ? 4 : 0)), prefs.c_str(), 10) == 0,
                    "Set reference equation preferences");
        }
        auto range = asObject(invoke(scratch.Get(), L"Range", DISPATCH_METHOD, {0L, 0L}));
        auto shape = asObject(invoke(get(scratch.Get(), L"InlineShapes").Get(), L"AddOLEObject", DISPATCH_METHOD,
            {L"Equation.DSMT4", Value(), false, false, Value(), Value(), Value(), range.Get()}));
        auto ole = get(shape.Get(), L"OLEFormat");
        invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L});
        object = get(ole.Get(), L"Object");
        active = true;
        if (item) {
            const std::wstring tex = item->tex;
            require(api.write(object.Get(), 1, tex.c_str(), static_cast<long>(tex.size())) == 0, "Write empty reference equation");
        }
        // 偏好检查保持空对象；写入 $q$ 会主动改成行内，从而掩盖 NEXT_EQN 的内外标志。
        auto native = readNative(api, object.Get());
        const auto style = inspectStyle(native);
        std::printf("SCRATCH_STYLE FONT=%s FULL=%s INLINE=%d\n", style.styles.at(2).first.c_str(),
                    style.sizes.at(0).second.c_str(), style.inlineEquation);
        require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close scratch OLE");
        active = false;
        invoke(scratch.Get(), L"Close", DISPATCH_METHOD, {0L});
        return native;
    } catch (...) {
        if (active) require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close failed scratch OLE");
        invoke(scratch.Get(), L"Close", DISPATCH_METHOD, {0L});
        throw;
    }
}

static EquationStyle scratchStyle(MathTypeApi& api, WordSession& word) {
    return inspectStyle(scratchNative(api, word, nullptr, L""));
}

static void verifyDefaultStyle(MathTypeApi& api, WordSession& word, const EquationStyle& before) {
    const auto after = scratchStyle(api, word);
    require(after.inlineEquation == before.inlineEquation && after.sizes == before.sizes && after.styles == before.styles,
            "Next ordinary equation uses unchanged defaults");
}

// 仅对比自建小样的文件字节；使用已有 C 运行库的宽路径读取，避免引入 iostream 体积。
static std::vector<char> sampleFileBytes(const std::wstring& path) {
    FILE* raw{};
    require(_wfopen_s(&raw, path.c_str(), L"rb") == 0 && raw, "Open sample for byte comparison");
    std::unique_ptr<FILE, decltype(&fclose)> file(raw, &fclose);
    std::vector<char> bytes;
    char buffer[4096];
    for (size_t count; (count = fread(buffer, 1, sizeof(buffer), file.get())) != 0;)
        bytes.insert(bytes.end(), buffer, buffer + count);
    require(ferror(file.get()) == 0 && !bytes.empty(), "Read sample for byte comparison");
    require(fclose(file.release()) == 0, "Close sample byte comparison");
    return bytes;
}

// 仅用于固定求和/上下标样本的等价命令比较，不改写用户 TeX。
static std::string sampleTexSignature(std::string tex) {
    tex.erase(std::remove_if(tex.begin(), tex.end(), [](char c) { return c == ' ' || c == '\r' || c == '\n' || c == '\t'; }), tex.end());
    for (const auto& replacement : {std::make_pair("\\limits", ""), std::make_pair("{n}", "n"), std::make_pair("{2}", "2")}) {
        for (size_t position = tex.find(replacement.first); position != std::string::npos; position = tex.find(replacement.first))
            tex.replace(position, std::char_traits<char>::length(replacement.first), replacement.second);
    }
    return tex;
}

// 仅检查本轮固定样本的未转义花括号；不是完整 LaTeX 语法验证器。
static bool sampleBracesBalanced(const std::wstring& tex) {
    long depth = 0;
    for (size_t index = 0; index < tex.size(); ++index) {
        if (tex[index] == L'\\' && index + 1 < tex.size()) { ++index; continue; }
        if (tex[index] == L'{') ++depth;
        if (tex[index] == L'}' && --depth < 0) return false;
    }
    return depth == 0;
}

// 只查找可见 MathType 窗口；按进程映像核对，不依赖可能缓存旧宿主名的标题。
static std::vector<HWND> mathTypeWindows() {
    struct Search { std::vector<HWND> windows; bool failed{}; } search;
    EnumWindows([](HWND hwnd, LPARAM parameter) -> BOOL {
        auto& found = *reinterpret_cast<Search*>(parameter);
        if (!IsWindowVisible(hwnd)) return TRUE;
        DWORD pid{};
        GetWindowThreadProcessId(hwnd, &pid);
        HANDLE process = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
        if (!process) return TRUE;
        wchar_t path[1024]{};
        DWORD length = 1024;
        const BOOL queried = QueryFullProcessImageNameW(process, 0, path, &length);
        CloseHandle(process);
        const auto name = wcsrchr(path, L'\\');
        if (!queried || !name || _wcsicmp(name + 1, L"MathType.exe") != 0) return TRUE;
        try { found.windows.push_back(hwnd); }
        catch (...) { found.failed = true; return FALSE; }
        return TRUE;
    }, reinterpret_cast<LPARAM>(&search));
    require(!search.failed, "Collect editor windows");
    return search.windows;
}

// 从当前编辑器菜单查找“全选”或“清除”，不能硬编码厂商菜单编号。
static void editCommands(HWND editor, HMENU menu, bool wantSelectAll, std::vector<UINT>& commands) {
    const int count = GetMenuItemCount(menu);
    require(count >= 0, "Read editor menu");
    for (int index = 0; index < count; ++index) {
        wchar_t caption[512]{};
        MENUITEMINFOW item{sizeof(item)};
        item.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
        item.dwTypeData = caption; item.cch = 512;
        require(GetMenuItemInfoW(menu, static_cast<UINT>(index), TRUE, &item) != 0, "Read editor menu item");
        if (item.hSubMenu) {
            // 菜单状态由 WM_INITMENUPOPUP 更新；未初始化的缓存可能把有效命令标为禁用。
            DWORD_PTR result{};
            require(SendMessageTimeoutW(editor, WM_INITMENUPOPUP, reinterpret_cast<WPARAM>(item.hSubMenu),
                                       MAKELPARAM(index, 0), SMTO_ABORTIFHUNG | SMTO_BLOCK, 3000, &result) != 0,
                    "Refresh editor menu state");
            editCommands(editor, item.hSubMenu, wantSelectAll, commands);
        }
        else {
            // 快捷键必须完整匹配：Ctrl+A 前缀也会命中 Ctrl+Alt 的其他菜单项。
            const auto shortcut = wcsrchr(caption, L'\t');
            std::wstring label(caption, shortcut ? shortcut : caption + wcslen(caption));
            label.erase(std::remove(label.begin(), label.end(), L'&'), label.end());
            if (label.size() >= 3 && label[label.size() - 3] == L'(' && label.back() == L')') label.resize(label.size() - 3);
            while (!label.empty() && label.back() == L' ') label.pop_back();
            const bool selected = wantSelectAll ?
                ((shortcut && _wcsicmp(shortcut + 1, L"Ctrl+A") == 0) || _wcsicmp(label.c_str(), L"Select All") == 0 || label == L"选择全部" || label == L"全选") :
                (_wcsicmp(label.c_str(), L"Clear") == 0 || label == L"清除");
            if (selected && (item.fState & (MFS_DISABLED | MFS_GRAYED)) == 0) commands.push_back(item.wID);
        }
    }
}

// 已保存的新公式通过双击等效主操作重开，选择全部内容后写入同一个 OLE。
static void editSample(MathTypeApi& api, IDispatch* shape, const ProbeCase& edited,
                       const std::vector<unsigned char>& nativeReference) {
    auto ole = get(shape, L"OLEFormat");
    require(text(ole.Get(), L"ClassType") == L"Equation.DSMT4", "Editable sample OLE type");
    const auto before = mathTypeWindows();
    // 自检要求没有其他可见 MathType 编辑器；有用户编辑窗口时不发送任何菜单命令。
    require(before.empty(), "Refuse editing test with another visible MathType window");
    invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {0L});
    auto object = get(ole.Get(), L"Object");
    bool active = true;
    try {
        const auto windows = mathTypeWindows();
        std::printf("EDIT_WINDOWS BEFORE=%zu MATCHED=%zu\n", before.size(), windows.size());
        require(windows.size() == 1 && std::find(before.begin(), before.end(), windows[0]) == before.end(),
                "Own newly visible editor window");
        for (bool selectAll : {true, false}) {
            std::vector<UINT> commands;
            editCommands(windows[0], GetMenu(windows[0]), selectAll, commands);
            require(commands.size() == 1 && commands[0] > 0 && commands[0] <= 65535, "Unique enabled editor command");
            // 命令只发给本条 OLE 主操作后唯一新可见的窗口，写入仍绑定本任务的具体 OLE。
            require(mathTypeWindows() == windows, "Editor identity before menu command");
            DWORD_PTR result{};
            require(SendMessageTimeoutW(windows[0], WM_COMMAND, MAKEWPARAM(commands[0], 0), 0,
                                       SMTO_ABORTIFHUNG | SMTO_BLOCK, 3000, &result) != 0, "Execute owned editor command");
            std::printf("EDIT_MENU_ACTION=%s ID=%u\n", selectAll ? "SELECT_ALL" : "CLEAR", commands[0]);
        }
        // 菜单清除实际旧内容并回存；SDK 导入在空对象中写入，避免其忽略交互选择范围。
        // 交互清除属于脏状态，必须 SAVEIFDIRTY=0；NOSAVE=1 会丢弃清除并造成旧式残留。
        require(api.close(OLECLOSE_SAVEIFDIRTY, object.Get()) == 0, "Save cleared editor contents");
        active = false;
        invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L});
        object = get(ole.Get(), L"Object");
        active = true;
        require(!nativeReference.empty() && nativeReference.size() < 1000000, "Edit native reference bounds");
        const std::wstring tex = edited.tex;
        require(api.write(object.Get(), 1, tex.c_str(), static_cast<long>(tex.size())) == 0, "Replace selected equation");
        require(sampleTexSignature(readTex(api, object.Get())) == edited.expected, "Edited TeX input cache");
        const auto editedNative = readNative(api, object.Get());
        if (editedNative != nativeReference) {
            std::printf("EDIT_NATIVE_ACTUAL=");
            for (auto byte : editedNative) std::printf("%02X", byte);
            std::printf("\nEDIT_NATIVE_REFERENCE=");
            for (auto byte : nativeReference) std::printf("%02X", byte);
            std::printf("\n");
        }
        require(editedNative == nativeReference, "Live edit native equals fresh empty reference");
        verifyStyle(inspectStyle(editedNative), edited.font, edited.points, edited.inlineEquation);
        require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close imported edited OLE");
        active = false;
        std::printf("EDIT_CASE=%s PRIMARY_VERB=0 CLEAR_SAVE_MODE=0 PASSED\n", edited.name);
    } catch (...) {
        if (active) require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close failed edit OLE");
        throw;
    }
}

static ComPtr<IDispatch> convertSample(MathTypeApi& api, WordSession& word, const ProbeCase& item,
                                     const std::wstring& directory, std::vector<unsigned char>& native,
                                     bool rejectCreation = false) {
    auto document = word.document.Get();
    auto range = findSampleRange(document, item.source);
    auto shapes = get(document, L"InlineShapes");
    const long countBefore = integer(shapes.Get(), L"Count");
    const auto bodyBefore = text(get(document, L"Content").Get(), L"Text");
    const auto sourceFont = get(range.Get(), L"Font");
    const auto fontBefore = text(sourceFont.Get(), L"Name");
    const long boldBefore = integer(sourceFont.Get(), L"Bold");
    auto undo = get(word.application.Get(), L"UndoRecord");
    invoke(undo.Get(), L"StartCustomRecord", DISPATCH_METHOD, {L"原生公式小样"});
    ComPtr<IDispatch> activeObject;
    bool active = false;
    bool pendingPrefs = false;
    try {
        // NEXT_EQN 必须在新对象创建前设置；激活后再设会留下偏好，污染下一次插入。
        const auto prefs = loadPrefs(api, directory + L"\\" + item.eqp);
        // 返回错误时也可能已设置部分状态；保守地纳入后续清理。
        pendingPrefs = true;
        require(api.setPrefs(static_cast<short>(1 + (item.inlineEquation ? 4 : 0)), prefs.c_str(), 10) == 0,
                "Set next equation preferences");
        invoke(range.Get(), L"Text", DISPATCH_PROPERTYPUT, {L""});
        invoke(range.Get(), L"Collapse", DISPATCH_METHOD, {1L});
        // 自检在真实创建调用前注入异常，覆盖“偏好已设、源式已清除、对象还未创建”的边界。
        // 无效 OLE 类名会触发 Word 模态对话框，不能用它做无人值守失败测试。
        require(!rejectCreation, "Injected creation failure");
        auto shape = asObject(invoke(shapes.Get(), L"AddOLEObject", DISPATCH_METHOD,
            {L"Equation.DSMT4", Value(), false, false, Value(), Value(), Value(), range.Get()}));
        auto ole = get(shape.Get(), L"OLEFormat");
        invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L});
        activeObject = get(ole.Get(), L"Object");
        active = true;
        pendingPrefs = false;
        const std::wstring tex = item.tex;
        const long status = api.write(activeObject.Get(), 1, tex.c_str(), static_cast<long>(tex.size()));
        std::printf("WRITE_CASE=%s STATUS=%ld\n", item.name, status);
        require(status == 0, "Write TeX");
        const auto cachedTex = readTex(api, activeObject.Get());
        // 小样故意送入破损式验证恢复：本机 API 可返回 0，不能仅凭返回码宣布转换成功。
        // 正式转换入口须将这类验证前置到扫描阶段，本函数目前只处理自建测试样本。
        require(sampleBracesBalanced(tex), "Malformed TeX returned as success");
        require(sampleTexSignature(cachedTex) == item.expected, "Sample TeX input cache");
        native = readNative(api, activeObject.Get());
        verifyStyle(inspectStyle(native), item.font, item.points, item.inlineEquation);
        require(api.close(OLECLOSE_NOSAVE, activeObject.Get()) == 0, "Close converted OLE");
        active = false;
        if (!item.inlineEquation) {
            auto paragraph = get(get(shape.Get(), L"Range").Get(), L"ParagraphFormat");
            invoke(paragraph.Get(), L"Alignment", DISPATCH_PROPERTYPUT, {1L});
            invoke(paragraph.Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {0L});
        }
        invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD);
        return shape;
    } catch (...) {
        // 先关闭本条新对象，再撤销整个记录；恢复无法确认时抛错终止，不能继续发布。
        if (active) require(api.close(OLECLOSE_NOSAVE, activeObject.Get()) == 0, "Close failed new OLE");
        invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD);
        // 准备阶段失败可能没有任何文档改动；空记录不能 Undo，否则会撤销上一条成功式。
        if (integer(shapes.Get(), L"Count") != countBefore || text(get(document, L"Content").Get(), L"Text") != bodyBefore) {
            const auto undone = invoke(document, L"Undo", DISPATCH_METHOD, {1L});
            require(undone.v.vt == VT_BOOL && undone.v.boolVal != VARIANT_FALSE, "Undo failed equation");
        }
        require(integer(shapes.Get(), L"Count") == countBefore &&
                text(get(document, L"Content").Get(), L"Text") == bodyBefore, "Recovered source and object count");
        auto recoveredFont = get(findSampleRange(document, item.source).Get(), L"Font");
        require(text(recoveredFont.Get(), L"Name") == fontBefore && integer(recoveredFont.Get(), L"Bold") == boldBefore,
                "Recovered source font");
        if (pendingPrefs) {
            // 文档恢复后，在独立临时文档消费待用样式；清理失败直接中止，不能发布候选。
            const auto consumed = scratchStyle(api, word);
            if (rejectCreation) verifyStyle(consumed, item.font, item.points, item.inlineEquation);
            std::puts("PENDING_PREFERENCES_CONSUMED_PASSED");
        }
        std::puts("FAILED_ITEM_RECOVERY_PASSED");
        throw;
    }
}

// 只向本任务唯一新打开的编辑器发送字符，不调用全局按键或剪贴板。
static void typeSampleCharacters(MathTypeApi& api, IDispatch* shape, const wchar_t* characters) {
    require(mathTypeWindows().empty(), "Refuse character test with visible user editor");
    auto ole = get(shape, L"OLEFormat");
    invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {0L});
    auto object = get(ole.Get(), L"Object");
    try {
        const auto windows = mathTypeWindows();
        require(windows.size() == 1, "Unique owned editor for character input");
        GUITHREADINFO info{sizeof(info)};
        require(GetGUIThreadInfo(GetWindowThreadProcessId(windows[0], nullptr), &info) != 0,
                "Read owned editor focus");
        require(info.hwndFocus && GetAncestor(info.hwndFocus, GA_ROOT) == windows[0], "Focus belongs to owned editor");
        wchar_t focusClass[256]{};
        GetClassNameW(info.hwndFocus, focusClass, 256);
        std::wprintf(L"TYPED_FOCUS=%ls CHARACTERS=%ls\n", focusClass, characters);
        DWORD_PTR result{};
        require(SendMessageTimeoutW(info.hwndFocus, WM_KEYDOWN, VK_END, 0,
                                   SMTO_ABORTIFHUNG | SMTO_BLOCK, 3000, &result) != 0, "Move sample caret to end");
        for (const wchar_t* character = characters; *character; ++character) {
            require(mathTypeWindows() == windows, "Editor identity before character input");
            // 定向入队，让编辑器自己的消息循环处理字符；不修改全局键盘状态。
            require(PostMessageW(info.hwndFocus, WM_CHAR, *character, 1) != 0, "Queue owned sample character");
            pump(100);
        }
        require(api.close(OLECLOSE_SAVEIFDIRTY, object.Get()) == 0, "Save character-edited native contents");
    } catch (...) {
        api.close(OLECLOSE_NOSAVE, object.Get());
        throw;
    }
}

// 阶段 5.1 只验证导出，不替换 DOCX 对象，不启用尚未验收的反向 GUI。
static void exportSelfTest(const std::wstring& directory, const std::wstring& editedInput) {
    std::wstring problem; const auto install = findMathTypeInstall(problem); require(!install.empty(), "Installed MathType required");
    MathTypeApi api(install);
    require(api.init(0, 10) > 0, "Initialize native export API");
    api.initialized = true;
    const auto defaults = readDefaultPrefs(api);
    WordSession word;
    auto documents = get(word.application.Get(), L"Documents");
    word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD));
    word.verifyOwnership();
    invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
    invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
    invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
    word.disableLinkUpdates();
    const auto editedBytes = sampleFileBytes(editedInput);
    auto content = get(word.document.Get(), L"Content");
    invoke(content.Get(), L"Text", DISPATCH_PROPERTYPUT,
           {L"原生字符输入行内：\r原生字符输入行外：\r输入后交互修改：\r输入缓存不完整：\r"});
    auto font = get(content.Get(), L"Font");
    invoke(font.Get(), L"Name", DISPATCH_PROPERTYPUT, {L"Times New Roman"});
    invoke(font.Get(), L"NameFarEast", DISPATCH_PROPERTYPUT, {L"宋体"});
    auto paragraphs = get(content.Get(), L"ParagraphFormat");
    invoke(paragraphs.Get(), L"Alignment", DISPATCH_PROPERTYPUT, {3L});
    invoke(paragraphs.Get(), L"SpaceBefore", DISPATCH_PROPERTYPUT, {0L});
    invoke(paragraphs.Get(), L"SpaceAfter", DISPATCH_PROPERTYPUT, {0L});
    invoke(paragraphs.Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {2L});
    struct Fixture { const wchar_t* label; const wchar_t* seed; const wchar_t* typed; bool inlineEquation; };
    const Fixture fixtures[] = {
        {L"原生字符输入行内：", L"", L"u+9", true},
        {L"原生字符输入行外：", L"", L"v+8", false},
        {L"输入后交互修改：", L"$x_1$", L"+7", true},
        {L"输入缓存不完整：", L"$a+1$", L"", true}
    };
    const auto prefs = loadPrefs(api, directory + L"\\Manual.eqp");
    for (const auto& fixture : fixtures) {
        auto range = findSampleRange(word.document.Get(), fixture.label);
        invoke(range.Get(), L"Collapse", DISPATCH_METHOD, {0L});
        require(api.setPrefs(fixture.inlineEquation ? 5 : 1, prefs.c_str(), -1) == 0, "Set sample type only for next equation");
        auto shape = asObject(invoke(get(word.document.Get(), L"InlineShapes").Get(), L"AddOLEObject", DISPATCH_METHOD,
            {L"Equation.DSMT4", Value(), false, false, Value(), Value(), Value(), range.Get()}));
        auto ole = get(shape.Get(), L"OLEFormat");
        invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L});
        auto object = get(ole.Get(), L"Object");
        if (*fixture.seed) require(api.write(object.Get(), 1, fixture.seed, static_cast<long>(wcslen(fixture.seed))) == 0,
                                   "Write fixed sample seed");
        if (fixture.typed[0] == 0) {
            // 故意复现已有对象追加导入：缓存仅有新式，真实数学结构包含旧式和新式。
            require(api.write(object.Get(), 1, L"$b+2$", 5) == 0, "Create real input-cache mismatch");
        }
        require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close API sample seed");
        if (*fixture.typed) typeSampleCharacters(api, shape.Get(), fixture.typed);
    }
    const std::wstring manual = directory + L"\\native-manual.docx";
    require(GetFileAttributesW(manual.c_str()) == INVALID_FILE_ATTRIBUTES, "Manual sample already exists");
    invoke(word.document.Get(), L"SaveAs2", DISPATCH_METHOD, {manual.c_str(), 16L});
    invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L});
    word.document.Reset();
    const auto manualBytes = sampleFileBytes(manual);
    const std::wstring copied = directory + L"\\edited-readonly.docx";
    require(CopyFileW(editedInput.c_str(), copied.c_str(), TRUE) != 0, "Copy previously edited input for read-only export");
    bool rejected{};
    try { exportNativeTex(api, std::vector<unsigned char>{5}, "DocxMathTypeBatch-missing-translator.tdl"); }
    catch (const std::exception& error) { rejected = std::string(error.what()) == "Select native TeX translator"; }
    require(rejected, "Missing translator rejected before native translation");
    long total{};
    for (const auto& input : {manual, copied}) {
        word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD, {input.c_str(), false, true, false}));
        auto shapes = get(word.document.Get(), L"InlineShapes");
        const long count = integer(shapes.Get(), L"Count");
        require(count == (input == manual ? 4 : 6), "Expected fixed export fixture count");
        const auto body = text(get(word.document.Get(), L"Content").Get(), L"Text");
        for (long index = 1; index <= count; ++index) {
            auto shape = asObject(invoke(shapes.Get(), L"Item", DISPATCH_METHOD, {index}));
            auto ole = get(shape.Get(), L"OLEFormat");
            invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L});
            auto object = get(ole.Get(), L"Object");
            try {
                const auto native = readNative(api, object.Get());
                const auto style = inspectStyle(native);
                // 精确字节串只作缓存存在证据，不解析、构造或改写 MTEF；缺失必无该扩展。
                const std::string cacheMarker = "TeX Input Language";
                const bool hasInputCache = std::search(native.begin(), native.end(), cacheMarker.begin(), cacheMarker.end()) != native.end();
                long cacheLength{};
                const auto cacheStatus = api.read(object.Get(), 1, nullptr, &cacheLength);
                const auto cache = cacheStatus == 0 && cacheLength > 0 ? readTex(api, object.Get()) : std::string();
                const auto exported = exportNativeTex(api, native);
                require(readNative(api, object.Get()) == native, "Export leaves complete native bytes unchanged");
                std::printf("EXPORT_CASE %ld %d %d %ld %s %s\n", ++total, style.inlineEquation, hasInputCache, cacheStatus,
                    cache.empty() ? "-" : wireHex(cache).c_str(), wireHex(exported).c_str());
                require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close read-only export object");
            } catch (...) {
                api.close(OLECLOSE_NOSAVE, object.Get());
                throw;
            }
        }
        require(text(get(word.document.Get(), L"Content").Get(), L"Text") == body, "Export leaves document text unchanged");
        invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L});
        word.document.Reset();
    }
    require(sampleFileBytes(editedInput) == editedBytes && sampleFileBytes(copied) == editedBytes &&
            sampleFileBytes(manual) == manualBytes, "Export leaves all input file bytes unchanged");
    require(readDefaultPrefs(api) == defaults, "Export leaves global MathType preferences unchanged");
    word.shutdown();
    require(api.term() == 0, "Terminate native export API");
    api.initialized = false;
    std::puts("EXPORT_INPUT_AND_NATIVE_UNCHANGED_PASSED");
    std::puts("EXPORT_MISSING_TRANSLATOR_REJECTED_PASSED");
    std::puts("NATIVE_EXPORT_PROBE_PASSED");
}

// 隔离测试注册表只影响本进程，绝不改写用户的 MathType/Word 注册项。
static void installationSelfTest(const std::wstring& directory) {
    std::wstring problem; const auto original=findMathTypeInstall(problem);
    require(!original.empty(),"Read original registered MathType installation");
    std::printf("REGISTERED_MATHTYPE %s\n",wireHex(utf8(original)).c_str());
    const auto temporaryKey=L"Software\\DocxMathTypeBatchInstallTest-"+std::to_wstring(GetCurrentProcessId());
    HKEY root{};require(RegCreateKeyExW(HKEY_CURRENT_USER,temporaryKey.c_str(),0,nullptr,REG_OPTION_VOLATILE,KEY_ALL_ACCESS,nullptr,&root,nullptr)==ERROR_SUCCESS,"Create private test registry");
    struct RegistryScope {
        HKEY root;const std::wstring& key;
        ~RegistryScope(){RegOverridePredefKey(HKEY_CLASSES_ROOT,nullptr);RegCloseKey(root);RegDeleteTreeW(HKEY_CURRENT_USER,key.c_str());}
    } registry{root,temporaryKey};
    require(RegOverridePredefKey(HKEY_CLASSES_ROOT,root)==ERROR_SUCCESS,"Override classes only inside test process");
    const auto serverKey=L"CLSID\\{0002CE03-0000-0000-C000-000000000046}\\LocalServer32";
    auto write=[&](const wchar_t* key,const wchar_t* name,const std::wstring& value,DWORD type=REG_SZ){
        HKEY entry{};require(RegCreateKeyExW(root,key,0,nullptr,REG_OPTION_VOLATILE,KEY_ALL_ACCESS,nullptr,&entry,nullptr)==ERROR_SUCCESS,"Create isolated registration");
        const auto status=RegSetValueExW(entry,name,0,type,reinterpret_cast<const BYTE*>(value.c_str()),static_cast<DWORD>((value.size()+1)*sizeof(wchar_t)));
        RegCloseKey(entry);require(status==ERROR_SUCCESS,"Write isolated registration");
    };
    auto verify=[&](const char* name,bool expected){
        const auto found=findMathTypeInstall(problem);
        require(expected?(found==directory&&problem.empty()):(found.empty()&&!problem.empty()),name);
        std::printf("INSTALLATION_CASE %s PASSED\n",name);
    };
    verify("missing-registration",false);
    write(L"Equation.DSMT4\\CLSID",nullptr,L"{0002CE03-0000-0000-C000-000000000046}");
    verify("missing-server",false);
    write(serverKey,nullptr,directory+L"\\MathType.exe");verify("unquoted-spaces-unicode",true);
    write(serverKey,nullptr,directory+L"\\MathType.exe /Embedding");verify("unquoted-arguments",true);
    write(serverKey,nullptr,L"\""+directory+L"\\MathType.exe\" -Embedding");verify("quoted-arguments",true);
    write(serverKey,nullptr,(directory+L"\\missing\\MathType.exe"));
    write(serverKey,L"ServerExecutable",directory+L"\\MathType.exe");verify("server-executable-priority",true);
    HKEY server{};require(RegOpenKeyExW(root,serverKey,0,KEY_SET_VALUE,&server)==ERROR_SUCCESS,"Open isolated server");
    const auto deleted=RegDeleteValueW(server,L"ServerExecutable");RegCloseKey(server);require(deleted==ERROR_SUCCESS,"Remove isolated executable value");
    const auto variable=L"DOCX_MATH_INSTALL_TEST_"+std::to_wstring(GetCurrentProcessId());
    require(SetEnvironmentVariableW(variable.c_str(),directory.c_str())!=0,"Set isolated expansion value");
    write(serverKey,nullptr,L"%"+variable+L"%\\MathType.exe -Embedding",REG_EXPAND_SZ);
    verify("expanded-registration",true);require(SetEnvironmentVariableW(variable.c_str(),nullptr)!=0,"Clear expansion value");
    write(serverKey,nullptr,L"\""+directory+L"\\Other.exe\"");verify("wrong-server",false);
    write(serverKey,nullptr,L"\""+directory+L"\\MathType.exe");verify("unterminated-quote",false);
    write(serverKey,nullptr,L"MathType\\MathType.exe");verify("relative-server",false);
    write(serverKey,nullptr,directory+L"\\MathType.exe.invalid");verify("wrong-extension",false);
    require(RegOverridePredefKey(HKEY_CLASSES_ROOT,nullptr)==ERROR_SUCCESS,"Restore real classes view");
    require(findMathTypeInstall(problem)==original&&problem.empty(),"Original registration unchanged");
    std::puts("INSTALLATION_DISCOVERY_PASSED");
}

#include "test_native_numbering.h"
#include "test_experimental_api.h"
int wmain(int argc, wchar_t** argv) {
    setvbuf(stdout, nullptr, _IONBF, 0);
    if(argc==4&&std::wstring(argv[1])==L"--experimental-native-compare"){
        if(FAILED(CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED)))return 1;int result=1;
        try{compareExperimentalNative(absolutePath(argv[2]),absolutePath(argv[3]));result=0;}catch(const std::exception& error){std::printf("EXPERIMENTAL_NATIVE_FAILED %s\n",error.what());}CoUninitialize();return result;
    }
    if(argc==4&&std::wstring(argv[1])==L"--experimental-fixture")return superviseTask(argc,argv);
    if(argc==6&&std::wstring(argv[1])==L"--task-run"&&std::wstring(argv[3])==L"--experimental-fixture"){
        if(FAILED(CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED)))return 1;int result=1;
        try{TaskMapping mapping(argv[2],false);taskShared=mapping.state;createExperimentalFixture(absolutePath(argv[4]),absolutePath(argv[5]));taskShared=nullptr;result=0;}
        catch(const std::exception& error){std::printf("EXPERIMENTAL_FIXTURE_FAILED %s\n",error.what());taskShared=nullptr;}CoUninitialize();return result;
    }
    if(argc==4&&std::wstring(argv[1])==L"--experimental-api-self-test")return superviseTask(argc,argv);
    if(argc==6&&std::wstring(argv[1])==L"--task-run"&&std::wstring(argv[3])==L"--experimental-api-self-test"){
        if(FAILED(CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED)))return 1;int result=1;
        try{TaskMapping mapping(argv[2],false);taskShared=mapping.state;checkExperimentalApi(absolutePath(argv[5]));taskShared=nullptr;result=0;}
        catch(const std::exception& error){std::printf("EXPERIMENTAL_API_FAILED %s\n",error.what());taskShared=nullptr;}CoUninitialize();return result;
    }
    if(argc==4&&(std::wstring(argv[1])==L"--numbering-self-test"||std::wstring(argv[1])==L"--numbering-fixture"||std::wstring(argv[1])==L"--display-self-test"))return superviseTask(argc,argv);
    if(argc==6&&std::wstring(argv[1])==L"--task-run"&&(std::wstring(argv[3])==L"--numbering-self-test"||std::wstring(argv[3])==L"--numbering-fixture"||std::wstring(argv[3])==L"--display-self-test")){
        if(FAILED(CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED)))return 1;int result=1;
        try{TaskMapping mapping(argv[2],false);taskShared=mapping.state;checkNativeNumbering(absolutePath(argv[4]),absolutePath(argv[5]),std::wstring(argv[3])==L"--numbering-fixture",std::wstring(argv[3])==L"--display-self-test");taskShared=nullptr;result=0;}
        catch(const std::exception& error){std::printf("NATIVE_NUMBER_CHECK_FAILED %s\n",error.what());taskShared=nullptr;}CoUninitialize();return result;
    }
    if(argc==4&&std::wstring(argv[1])==L"--link-settings-self-test"){
        if(FAILED(CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED)))return 1;int result=1;
        try{
            const auto originalBytes=sampleFileBytes(argv[2]);require(CopyFileW(argv[2],argv[3],TRUE)!=0,"Create private link settings test copy");
            const auto original=setDocxLinkUpdates(argv[3]);require(!original.empty(),"Real linked OLE fixture required");
            const auto manual=setDocxLinkUpdates(argv[3]);for(const auto& part:manual)for(const auto& mode:part.second)require(mode==L"OnCall","Private linked OLE updates disabled");
            setDocxLinkUpdates(argv[3],&original);const auto restored=setDocxLinkUpdates(argv[3]);require(restored==original,"Original linked OLE update modes restored");
            setDocxLinkUpdates(argv[3],&restored);require(sampleFileBytes(argv[2])==originalBytes,"Linked OLE source unchanged");
            std::puts("PRIVATE_LINK_SETTINGS_PASSED");result=0;
        }catch(const std::exception& error){std::printf("LINK_SETTINGS_FAILED %s\n",error.what());}CoUninitialize();return result;
    }
    if(argc==3&&(std::wstring(argv[1])==L"--installation-self-test"||std::wstring(argv[1])==L"--installation-check")) {
        try {
            if(std::wstring(argv[1])==L"--installation-self-test")installationSelfTest(absolutePath(argv[2]));
            else {const auto problem=mathTypeInstallProblem(absolutePath(argv[2]));require(problem.empty(),utf8(problem).c_str());}
            return 0;
        }catch(const std::exception& error){std::printf("INSTALLATION_FAILED %s\n",error.what());return 1;}
    }
    if (argc == 4 && std::wstring(argv[1]) == L"--export-self-test") {
        if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) return 1;
        int result = 1;
        try { exportSelfTest(argv[2], argv[3]); result = 0; }
        catch (const std::exception& error) { std::printf("NATIVE_EXPORT_PROBE_FAILED: %s\n", error.what()); }
        CoUninitialize();
        return result;
    }
    if (argc != 3 || std::wstring(argv[1]) != L"--self-test") {
        std::puts("usage: native-api-probe --self-test <new-output-directory>");
        return 2;
    }
    HRESULT initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(initialized)) return 1;
    int exitCode = 1;
    try {
        std::wstring problem; const auto install = findMathTypeInstall(problem); require(!install.empty(), "Installed MathType required");
        const std::wstring output = std::wstring(argv[2]) + L"\\native-probe.docx";
        require(GetFileAttributesW(output.c_str()) == INVALID_FILE_ATTRIBUTES, "Output already exists");
        MathTypeApi api(install);
        const long init = api.init(0, 10);
        std::printf("MTInitAPI=%ld\n", init);
        // Windows 初始化返回正数表示成功；其余写入/读取接口采用 mtOK=0。
        require(init > 0, "MathType initialization");
        api.initialized = true;
        WordSession word;
        auto documents = get(word.application.Get(), L"Documents");
        word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD));
        word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
        invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
        auto document = word.document;
        auto content = get(document.Get(), L"Content");
        invoke(content.Get(), L"Text", DISPATCH_PROPERTYPUT,
            {L"Existing: $z^2$ remains\rbefore $x_1$ after\r$$\\frac{a}{b}$$\r\\[\\sum_{i=1}^{n}i\\]\rInvalid: $\\frac{$ preserved\r表1-1 表格范围与恢复验证\r"});
        // 自建 Word 样本也遵守用户正文约定；这里设置正文，不代表公式内部样式已验收。
        auto font = get(content.Get(), L"Font");
        invoke(font.Get(), L"Name", DISPATCH_PROPERTYPUT, {L"Times New Roman"});
        invoke(font.Get(), L"NameFarEast", DISPATCH_PROPERTYPUT, {L"宋体"});
        auto paragraph = get(content.Get(), L"ParagraphFormat");
        invoke(paragraph.Get(), L"Alignment", DISPATCH_PROPERTYPUT, {3L});
        invoke(paragraph.Get(), L"SpaceBefore", DISPATCH_PROPERTYPUT, {0L});
        invoke(paragraph.Get(), L"SpaceAfter", DISPATCH_PROPERTYPUT, {0L});
        invoke(paragraph.Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {2L});
        auto caption = findSampleRange(document.Get(), L"表1-1 表格范围与恢复验证");
        invoke(get(caption.Get(), L"Font").Get(), L"Size", DISPATCH_PROPERTYPUT, {10.5});
        invoke(get(caption.Get(), L"ParagraphFormat").Get(), L"Alignment", DISPATCH_PROPERTYPUT, {1L});
        invoke(get(caption.Get(), L"ParagraphFormat").Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {0L});
        const long tableStart = integer(get(document.Get(), L"Content").Get(), L"End") - 1;
        auto tableAnchor = asObject(invoke(document.Get(), L"Range", DISPATCH_METHOD, {tableStart, tableStart}));
        auto table = asObject(invoke(get(document.Get(), L"Tables").Get(), L"Add", DISPATCH_METHOD, {tableAnchor.Get(), 5L, 2L}));
        const wchar_t* cellTexts[][2] = {
            {L"场景", L"公式与保留文本"},
            {L"行内（邻格保持）", L"cell before $t_2$ cell after"},
            {L"行外（邻格保持）", L"$$\\frac{c}{d}$$"},
            {L"破损式（邻格保持）", L"Invalid cell: ${v$ retained"},
            {L"创建失败（邻格保持）", L"Create failure: $r+1$ retained"},
        };
        for (long row = 1; row <= 5; ++row)
            for (long column = 1; column <= 2; ++column)
                invoke(cellRange(table.Get(), row, column).Get(), L"Text", DISPATCH_PROPERTYPUT, {cellTexts[row - 1][column - 1]});
        auto tableRange = get(table.Get(), L"Range");
        invoke(get(tableRange.Get(), L"Font").Get(), L"Size", DISPATCH_PROPERTYPUT, {10.5});
        auto tableParagraph = get(tableRange.Get(), L"ParagraphFormat");
        invoke(tableParagraph.Get(), L"Alignment", DISPATCH_PROPERTYPUT, {1L});
        invoke(tableParagraph.Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {0L});
        invoke(get(table.Get(), L"Borders").Get(), L"Enable", DISPATCH_PROPERTYPUT, {0L});
        blackBorder(table.Get(), -1, 12); // 上边与下边 1.5 pt。
        blackBorder(table.Get(), -3, 12);
        auto headerRow = asObject(invoke(get(table.Get(), L"Rows").Get(), L"Item", DISPATCH_METHOD, {1L}));
        blackBorder(headerRow.Get(), -3, 6); // 表头下线 0.75 pt。
        invoke(get(get(headerRow.Get(), L"Range").Get(), L"Font").Get(), L"Bold", DISPATCH_PROPERTYPUT, {true});
        const std::wstring directory = argv[2];
        const ProbeCase cases[] = {
            {"existing", L"$z^2$", L"$z^2$", L"Times12.eqp", "Times New Roman", "12", "$z^2$", true},
            {"invalid", L"$\\frac{$", L"$\\frac{$", L"Arial18.eqp", "Arial", "18", "", true},
            {"inline", L"$x_1$", L"$x_1$", L"Arial18.eqp", "Arial", "18", "$x_1$", true},
            {"dollar-display", L"$$\\frac{a}{b}$$", L"\\[\\frac{a}{b}\\]", L"Times14.eqp", "Times New Roman", "14", "\\[\\frac{a}{b}\\]", false},
            {"bracket-display", L"\\[\\sum_{i=1}^{n}i\\]", L"\\[\\sum_{i=1}^{n}i\\]", L"Arial16.eqp", "Arial", "16", "\\[\\sum_{i=1}^ni\\]", false},
            {"create-failure", L"$r+1$", L"$r+1$", L"Arial18.eqp", "Arial", "18", "$r+1$", true, 5},
            {"table-inline", L"$t_2$", L"$t_2$", L"Times12.eqp", "Times New Roman", "12", "$t_2$", true, 2},
            {"table-display", L"$$\\frac{c}{d}$$", L"\\[\\frac{c}{d}\\]", L"Arial16.eqp", "Arial", "16", "\\[\\frac{c}{d}\\]", false, 3},
            {"table-invalid", L"${v$", L"${v$", L"Times14.eqp", "Times New Roman", "14", "", true, 4},
        };
        std::vector<std::vector<unsigned char>> nativeByCase(std::size(cases));
        const auto defaultsBefore = readDefaultPrefs(api);
        const auto defaultStyleBefore = scratchStyle(api, word);
        convertSample(api, word, cases[0], directory, nativeByCase[0]);
        // 无效式加粗，用实际失败验证 Undo 同时恢复文本和字符格式。
        invoke(get(findSampleRange(document.Get(), cases[1].source).Get(), L"Font").Get(), L"Bold", DISPATCH_PROPERTYPUT, {true});
        invoke(get(findSampleRange(document.Get(), cases[5].source).Get(), L"Font").Get(), L"Bold", DISPATCH_PROPERTYPUT, {true});
        invoke(get(findSampleRange(document.Get(), cases[8].source).Get(), L"Font").Get(), L"Bold", DISPATCH_PROPERTYPUT, {true});
        const std::wstring input = directory + L"\\native-input.docx";
        require(GetFileAttributesW(input.c_str()) == INVALID_FILE_ATTRIBUTES, "Input sample already exists");
        invoke(document.Get(), L"SaveAs2", DISPATCH_METHOD, {input.c_str(), 12L});
        for (size_t index = 1; index < std::size(cases); ++index) {
            try {
                convertSample(api, word, cases[index], directory, nativeByCase[index], index == 5);
                require(index != 1 && index != 8, "Invalid TeX was unexpectedly accepted");
            } catch (const std::runtime_error& error) {
                if (index == 5 && std::string(error.what()) == "Injected creation failure") {
                    // 不调用 SetMTPrefs 的下一对象必须回到自然默认，不能仅检查全局偏好串。
                    verifyDefaultStyle(api, word, defaultStyleBefore);
                    std::puts("CREATION_FAILURE_AND_NEXT_DEFAULT_PASSED");
                } else {
                    if ((index != 1 && index != 8) || (std::string(error.what()) != "Write TeX" &&
                        std::string(error.what()) != "Malformed TeX returned as success")) throw;
                    verifyDefaultStyle(api, word, defaultStyleBefore);
                    std::puts(index == 1 ? "EXPECTED_INVALID_TEX_PASSED" : "TABLE_INVALID_TEX_RECOVERY_PASSED");
                }
            }
        }
        // 再覆盖未改动文档的准备失败，确认不会误撤销上一条成功的表格公式。
        auto preparationFailure = cases[1];
        preparationFailure.eqp = L"Missing.eqp";
        try {
            std::vector<unsigned char> unusedNative;
            convertSample(api, word, preparationFailure, directory, unusedNative);
            throw std::runtime_error("Missing EQP unexpectedly accepted");
        } catch (const std::runtime_error& error) {
            if (std::string(error.what()) != "EQP short path") throw;
            std::puts("PREPARATION_FAILURE_UNCHANGED_PASSED");
        }
        invoke(document.Get(), L"SaveAs2", DISPATCH_METHOD, {output.c_str(), 12L});
        // 仅自检导出实际 Word 版式，供人工查看表格与预览，不进入转换产品流程。
        const std::wstring previewPdf = directory + L"\\native-preview.pdf";
        invoke(document.Get(), L"ExportAsFixedFormat", DISPATCH_METHOD, {previewPdf.c_str(), 17L});
        invoke(document.Get(), L"Close", DISPATCH_METHOD, {0L});
        word.document.Reset();
        std::vector<std::vector<unsigned char>> savedBeforeEdit(std::size(cases));
        std::vector<std::vector<unsigned char>> expectedEditedNative(std::size(cases));
        std::wstring savedBodyBeforeEdit;
        // 复用同一套真实 Range、表格、原生内容和内部样式检查，避免编辑验证漏掉原有边界。
        auto verifySaved = [&](const std::wstring& path, const ProbeCase* expectedCases, bool checkEdits) {
            auto reopened = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD,
                                           {path.c_str(), false, true, false}));
            word.document = reopened;
            const auto savedBody = text(get(reopened.Get(), L"Content").Get(), L"Text");
            if (checkEdits) require(savedBody == savedBodyBeforeEdit, "Edited document surrounding text unchanged");
            else savedBodyBeforeEdit = savedBody;
            auto reopenedShapes = get(reopened.Get(), L"InlineShapes");
            require(integer(reopenedShapes.Get(), L"Count") == 6, "Reopened object count");
            const wchar_t* prefixes[] = {L"Existing: ", L"before ", L"", L"", L"cell before ", L""};
            const wchar_t* suffixes[] = {L" remains\r", L" after\r", L"\r", L"\r", L" cell after\r\a", L"\r\a"};
            require(integer(get(reopened.Get(), L"Tables").Get(), L"Count") == 1, "Saved table count");
            auto savedTable = asObject(invoke(get(reopened.Get(), L"Tables").Get(), L"Item", DISPATCH_METHOD, {1L}));
            require(integer(get(savedTable.Get(), L"Rows").Get(), L"Count") == 5 &&
                    integer(get(savedTable.Get(), L"Columns").Get(), L"Count") == 2 &&
                    integer(get(get(savedTable.Get(), L"Range").Get(), L"Cells").Get(), L"Count") == 10, "Saved table dimensions");
            for (long row = 1; row <= 5; ++row) {
                require(text(cellRange(savedTable.Get(), row, 1).Get(), L"Text") == std::wstring(cellTexts[row - 1][0]) + L"\r\a",
                        "Unchanged adjacent cell");
                for (long column = 1; column <= 2; ++column)
                    require(integer(get(cellRange(savedTable.Get(), row, column).Get(), L"Paragraphs").Get(), L"Count") == 1,
                            "No additional cell paragraph");
            }
            require(text(cellRange(savedTable.Get(), 1, 2).Get(), L"Text") == std::wstring(cellTexts[0][1]) + L"\r\a" &&
                    text(cellRange(savedTable.Get(), 4, 2).Get(), L"Text") == std::wstring(cellTexts[3][1]) + L"\r\a" &&
                    text(cellRange(savedTable.Get(), 5, 2).Get(), L"Text") == std::wstring(cellTexts[4][1]) + L"\r\a",
                    "Saved header and failed cell sources");
            long objectIndex = 0;
            for (size_t index = 0; index < std::size(cases); ++index) {
                if (index == 1 || index == 5 || index == 8) continue;
                auto shape = asObject(invoke(reopenedShapes.Get(), L"Item", DISPATCH_METHOD, {++objectIndex}));
                auto objectRange = get(shape.Get(), L"Range");
                auto paragraphItem = asObject(invoke(get(objectRange.Get(), L"Paragraphs").Get(), L"Item", DISPATCH_METHOD, {1L}));
                auto paragraphRange = get(paragraphItem.Get(), L"Range");
                auto leading = asObject(invoke(reopened.Get(), L"Range", DISPATCH_METHOD,
                    {integer(paragraphRange.Get(), L"Start"), integer(objectRange.Get(), L"Start")}));
                auto trailing = asObject(invoke(reopened.Get(), L"Range", DISPATCH_METHOD,
                    {integer(objectRange.Get(), L"End"), integer(paragraphRange.Get(), L"End")}));
                require(text(leading.Get(), L"Text") == prefixes[objectIndex - 1] &&
                        text(trailing.Get(), L"Text") == suffixes[objectIndex - 1], "Exact surrounding text");
                if (!cases[index].inlineEquation) {
                    auto savedParagraph = get(objectRange.Get(), L"ParagraphFormat");
                    require(integer(savedParagraph.Get(), L"Alignment") == 1, "Display paragraph centered");
                    const auto indent = invoke(savedParagraph.Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYGET);
                    require((indent.v.vt == VT_R4 && indent.v.fltVal == 0) ||
                            (indent.v.vt == VT_R8 && indent.v.dblVal == 0), "Display paragraph no indent");
                }
                if (cases[index].tableRow != 0) {
                    auto cell = cellRange(savedTable.Get(), cases[index].tableRow, 2);
                    require(integer(objectRange.Get(), L"Start") >= integer(cell.Get(), L"Start") &&
                            integer(objectRange.Get(), L"End") <= integer(cell.Get(), L"End") - 1 &&
                            integer(get(cell.Get(), L"InlineShapes").Get(), L"Count") == 1, "OLE remains inside original cell");
                }
                auto ole = get(shape.Get(), L"OLEFormat");
                require(text(ole.Get(), L"ClassType") == L"Equation.DSMT4", "Reopened editable type");
                invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L});
                auto object = get(ole.Get(), L"Object");
                require(sampleTexSignature(readTex(api, object.Get())) == expectedCases[index].expected, "Saved TeX input cache");
                const auto savedNative = readNative(api, object.Get());
                verifyStyle(inspectStyle(savedNative), cases[index].font, cases[index].points, cases[index].inlineEquation);
                if (index == 0) require(savedNative == nativeByCase[0], "Unchanged existing native equation");
                if (checkEdits) {
                    const bool changed = std::string(expectedCases[index].expected) != cases[index].expected;
                    require((savedNative != savedBeforeEdit[index]) == changed, "Only intended native equations changed");
                    if (changed) require(savedNative == expectedEditedNative[index], "Edited native equals fresh empty reference");
                } else savedBeforeEdit[index] = savedNative;
                require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close reopened OLE");
            }
            require(integer(get(findSampleRange(reopened.Get(), cases[1].source).Get(), L"Font").Get(), L"Bold") == -1,
                    "Saved failed source format");
            for (size_t failed : {5u, 8u})
                require(integer(get(findSampleRange(reopened.Get(), cases[failed].source).Get(), L"Font").Get(), L"Bold") == -1,
                        "Saved failed cell source format");
            require(text(get(reopened.Get(), L"Content").Get(), L"Text").find(L"Invalid: $\\frac{$ preserved") != std::wstring::npos,
                    "Saved failed source text");
            std::puts("EXISTING_EQUATION_UNCHANGED_PASSED");
            std::puts("STYLE_AND_DISPLAY_PASSED");
            std::puts("SAVED_NATIVE_CONTENT_PASSED");
            std::puts("TABLE_STRUCTURE_AND_CONTENT_PASSED");
            if (checkEdits) std::puts("EDIT_UNTARGETED_NATIVE_UNCHANGED_PASSED");
            if (checkEdits) std::puts("EDIT_NATIVE_REFERENCE_MATCH_PASSED");
            invoke(reopened.Get(), L"Close", DISPATCH_METHOD, {0L});
            word.document.Reset();
        };
        verifySaved(output, cases, false);

        // 编辑只发生在新建副本；原转换结果完整字节必须保持。
        const auto outputBeforeEdit = sampleFileBytes(output);
        const std::wstring editedOutput = directory + L"\\native-edited.docx";
        require(CopyFileW(output.c_str(), editedOutput.c_str(), TRUE) != 0, "Create independent edit sample");
        auto editDocument = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD,
            {editedOutput.c_str(), false, false, false}));
        word.document = editDocument;
        auto editedShapes = get(editDocument.Get(), L"InlineShapes");
        std::vector<ProbeCase> editedCases(cases, cases + std::size(cases));
        struct EditTarget { size_t fixture; long objectNumber; const wchar_t* tex; const char* expected; };
        const EditTarget editTargets[] = {
            {2, 2, L"$x_1+7$", "$x_1+7$"},
            {3, 3, L"\\[\\frac{a+2}{b}\\]", "\\[\\frac{a+2}{b}\\]"},
            {6, 5, L"$t_2+3$", "$t_2+3$"},
            {7, 6, L"\\[\\frac{c+1}{d}\\]", "\\[\\frac{c+1}{d}\\]"},
        };
        for (const auto& target : editTargets) {
            auto& editedCase = editedCases[target.fixture];
            editedCase.tex = target.tex;
            editedCase.expected = target.expected;
            auto shape = asObject(invoke(editedShapes.Get(), L"Item", DISPATCH_METHOD, {target.objectNumber}));
            expectedEditedNative[target.fixture] = scratchNative(api, word, &editedCase, directory);
            editSample(api, shape.Get(), editedCase, expectedEditedNative[target.fixture]);
        }
        require(integer(editedShapes.Get(), L"Count") == 6, "Edited document object count unchanged");
        require(text(get(editDocument.Get(), L"Content").Get(), L"Text") == savedBodyBeforeEdit,
                "Editing leaves surrounding text unchanged");
        invoke(editDocument.Get(), L"Save", DISPATCH_METHOD);
        const std::wstring editedPdf = directory + L"\\native-edited-preview.pdf";
        invoke(editDocument.Get(), L"ExportAsFixedFormat", DISPATCH_METHOD, {editedPdf.c_str(), 17L});
        invoke(editDocument.Get(), L"Close", DISPATCH_METHOD, {0L});
        word.document.Reset();
        verifySaved(editedOutput, editedCases.data(), true);
        require(sampleFileBytes(output) == outputBeforeEdit, "Original conversion file bytes unchanged after edit");
        std::puts("EDIT_ORIGINAL_FILE_UNCHANGED_PASSED");
        std::puts("EDIT_ROUND_TRIP_PASSED");
        verifyDefaultStyle(api, word, defaultStyleBefore);
        require(readDefaultPrefs(api) == defaultsBefore, "Unchanged global MathType preferences after editing");
        std::puts("GLOBAL_PREFERENCES_UNCHANGED_PASSED");
        // 清理也属于成功门槛，不在关闭失败时报告完整通过。
        word.shutdown();
        const long terminated = api.term();
        std::printf("MTTermAPI=%ld\n", terminated);
        require(terminated == 0, "Terminate MathType API");
        api.initialized = false;
        std::puts("NATIVE_API_PROBE_PASSED");
        exitCode = 0;
    } catch (const std::exception& error) {
        std::printf("NATIVE_API_PROBE_FAILED: %s\n", error.what());
    }
    CoUninitialize();
    return exitCode;
}
