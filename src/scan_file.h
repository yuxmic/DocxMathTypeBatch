// 只读扫描文件任务与报告发布；文档扫描器只负责读取已打开的 Word 文档。
#pragma once
#include "docx_scan.h"
#include "docx_package.h"

inline void writeReport(const std::wstring& path, const std::wstring& input, const ScanResult& result) {
    const auto ready = std::count_if(result.items.begin(), result.items.end(), [](const ScanItem& item) { return item.ready; });
    std::string json = "{\n  \"Input\": " + quote(input) + ",\n  \"Scope\": \"MAIN_STORY_BODY_ONLY\",\n";
    json += "  \"Status\": \"" + std::string(result.protectedDocument ? "SKIPPED_PROTECTED_DOCUMENT" : "SCANNED") + "\",\n";
    json += "  \"SchemaVersion\": 1, \"ReadOnly\": true, \"Positions\": \"WORD_MAIN_STORY_SNAPSHOT\",\n";
    if (result.reverse) json += "  \"Direction\":\"REVERSE\", \"MathTypeTypeSource\":\"READ_ONLY_EQUATION_NATIVE\",\n";
    json += "  \"Validation\": \"BOUNDARIES_AND_BASIC_SYNTAX_ONLY\", \"MustRescanBeforeConvert\": true,\n";
    // 整篇保护时没有扫描，计数为未知，不能以 0 冒充检测结果。
    auto count = [&](size_t value) { return result.protectedDocument ? std::string("null") : std::to_string(value); };
    json += "  \"Ready\": " + count(static_cast<size_t>(ready)) + ", \"Rejected\": " + count(result.items.size() - static_cast<size_t>(ready)) + ",\n";
    json += "  \"ExistingMathType\": " + count(result.existingMathType) + ", \"InlineShapes\": " + count(result.inlineShapes) + ",\n";
    const auto ommlTargets=std::count_if(result.items.begin(),result.items.end(),[](const ScanItem& item){return item.ommlOrdinal!=0;});
    json += "  \"Excluded\": {\"OMML\":" + count(result.nativeMath-ommlTargets) + ",\"Fields\":" + count(result.fields) +
            ",\"Revisions\":" + count(result.revisions) + ",\"ContentControls\":" + count(result.controls) +
            ",\"NoteReferences\":" + count(result.noteReferences) + "},\n  \"Items\": [\n";
    for (size_t index = 0; index < result.items.size(); ++index) {
        const auto& item = result.items[index];
        json += "    {\"Id\":" + std::to_string(index + 1) + ",\"Start\":" + std::to_string(item.span.start) + ",\"End\":" + std::to_string(item.span.end) +
            ",\"Paragraph\":" + std::to_string(item.paragraph) + ",\"Region\":\"" + (item.row ? std::string("TABLE") : std::string("BODY")) +
            "\",\"TableStart\":" + std::to_string(item.tableStart) + ",\"Row\":" + std::to_string(item.row) + ",\"Column\":" + std::to_string(item.column) +
            ",\"Delimiter\":" + quote(item.delimiter) + ",\"Display\":" + (item.display ? "true" : "false") +
            ",\"Status\":\"" + (item.ready ? "READY" : "REJECTED") + "\",\"Reason\":" + quote(item.reason) +
            ",\"Source\":" + quote(item.source) + ",\"Inner\":" + quote(item.inner) + "}" + (index + 1 == result.items.size() ? "\n" : ",\n");
    }
    json += "  ]\n}\n";
    // CREATE_NEW 防止覆盖任何已存在的清单；关闭失败或短写入不宣布通过。
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(file != INVALID_HANDLE_VALUE, "Create new scan report");
    taskTrackFile(path, file, true);
    DWORD written{};
    const bool saved = json.size() <= MAXDWORD && WriteFile(file, json.data(), static_cast<DWORD>(json.size()), &written, nullptr) && written == json.size();
    const bool closed = CloseHandle(file) != 0;
    if (!saved || !closed) { DeleteFileW(path.c_str()); throw std::runtime_error("Write complete scan report"); }
}

// GUI 与命令行共用同一只读扫描；调用方所在工作进程必须先初始化 STA。
inline ScanResult scanReadOnlyFile(const std::wstring& input, const std::wstring& report,
                                  ScanResult (*scanner)(IDispatch*,const ScanOptions&) = scanDocument, const ScanOptions& options={}) {
    require(input.size() > 5 && _wcsicmp(input.c_str() + input.size() - 5, L".docx") == 0, "DOCX input required");
    require(GetFileAttributesW(report.c_str()) == INVALID_FILE_ATTRIBUTES, "Report already exists");
    const auto copy = report + L".scan-" + std::to_wstring(GetCurrentProcessId()) + L".docx";
    HANDLE source = CreateFileW(input.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(source != INVALID_HANDLE_VALUE, "Lock source during read-only scan");
    bool owned = false;
    try {
        // 在只读共享锁内检查输入，避免检查后被替换成加密文件而触发 Word 对话框。
        requireDocxZip(input);
        taskCopy(input, copy); owned = true;setDocxLinkUpdates(copy);
        WordSession word;
        auto documents = get(word.application.Get(), L"Documents");
        word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD)); word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
        invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
        invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L}); word.document.Reset();
        taskTrackWordOwner(copy);
        word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD, {copy.c_str(), false, true, false}));
        const auto readOnly = invoke(word.document.Get(), L"ReadOnly", DISPATCH_PROPERTYGET);
        require(readOnly.v.vt == VT_BOOL && readOnly.v.boolVal != VARIANT_FALSE, "Scan document opened read-only");
        const auto before = text(get(word.document.Get(), L"Content").Get(), L"Text");
        auto result = scanner(word.document.Get(),options);
        require(text(get(word.document.Get(), L"Content").Get(), L"Text") == before, "Scanning leaves Word text unchanged");
        word.shutdown();
        require(DeleteFileW(copy.c_str()) != 0, "Remove owned scan copy"); owned = false;
        require(CloseHandle(source) != 0, "Release scan source lock"); source = INVALID_HANDLE_VALUE;
        const auto staged=report+L".work-"+std::to_wstring(GetCurrentProcessId());taskReservePrivatePath(staged);
        writeReport(staged, input, result);taskMove(staged,report);
        taskCommit();
        return result;
    } catch (...) {
        if (owned && !DeleteFileW(copy.c_str())) std::puts("OWNED_SCAN_COPY_CLEANUP_FAILED");
        if (source != INVALID_HANDLE_VALUE) CloseHandle(source);
        throw;
    }
}

// GUI 只接收统计，公式原文和定位始终留在转换进程中。
inline void emitGuiScan(const ScanResult& result) {
    const auto ready=std::count_if(result.items.begin(),result.items.end(),[](const ScanItem& item){return item.ready;});
    const auto inlines=std::count_if(result.items.begin(),result.items.end(),[](const ScanItem& item){return item.ready&&!item.display;});
    const auto rejected=std::count_if(result.items.begin(),result.items.end(),[](const ScanItem& item){return !item.ready&&!item.existingDisplay&&!item.existingReference;});
    std::printf("UI_SCAN_COUNTS %ld %td %td %td\n",result.existingMathType,ready,inlines,rejected);
}

