// 阶段 2：只读扫描正文与表格；样本创建仅用于自检。
#include "scan_file.h"

// 下方仅自检入口创建/格式化任务副本；正式 --scan 路径从不保存 Word 文档。
static void makeFixture(WordSession& word, const std::wstring& base, const std::wstring& directory) {
    const auto path = directory + L"\\scan-fixture.docx";
    require(CopyFileW(base.c_str(), path.c_str(), TRUE) != 0, "Create independent scan fixture");
    auto documents = get(word.application.Get(), L"Documents");
    word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD, {path.c_str(), false, false, false}));
    auto document = word.document.Get();
    const long start = integer(get(document, L"Content").Get(), L"End") - 1;
    const wchar_t* cases =
        L"PAIR: before $alpha_1$ then $beta_2$ after\r"
        L"DISPLAY: $$\\frac{u}{v}$$ and \\[\\sum_{j=1}^{m}j\\]\r"
        L"SPLIT: $split_1+split_2$\r"
        L"ESCAPE: price \\$25 and \\$50; $ok_1$\r"
        L"EVEN: \\\\$even_1$\r"
        L"ESC_BRACKET: \\\\[not_math\\\\]\r"
        L"EMPTY: $ $ and $$ $$ and \\[ \\]\r"
        L"BRACES: ${a$ then $b}$ and $after_1$\r"
        L"LITERAL_BRACES: $\\{v\\}$\r"
        L"UNICODE: 😀前 $\\text{汉字}$ 后\r"
        L"BROKEN: $no_end\r"
        L"CROSS_PARAGRAPH: $$left\rright$$\r"
        L"ORPHAN: \\]\r"
        L"MIXED: $x\\] and $later_1$\r"
        L"AMBIGUOUS: $$$bad$$$\r"
        L"MANUAL_BREAK: \\[a\vb\\]\r"
        L"FIELD: #FIELD#\r"
        L"FIELD_CROSS: $left #CROSS_FIELD# right$\r"
        L"OMML: $omml_only$\r"
        L"REVISION: \r"
        L"CONTROL: #CONTROL#\r"
        L"NOTE: $note_left #NOTE# note_right$\r"
        L"PICTURE: $image_left #PICTURE# image_right$\r"
        L"图1-1 普通图片边界\r"
        L"表1-2 扫描单元格边界\r";
    invoke(docRange(document, start, start).Get(), L"InsertAfter", DISPATCH_METHOD, {cases});
    auto added = docRange(document, start, integer(get(document, L"Content").Get(), L"End"));
    auto font = get(added.Get(), L"Font");
    invoke(font.Get(), L"Name", DISPATCH_PROPERTYPUT, {L"Times New Roman"});
    invoke(font.Get(), L"NameFarEast", DISPATCH_PROPERTYPUT, {L"宋体"});
    auto para = get(added.Get(), L"ParagraphFormat");
    invoke(para.Get(), L"Alignment", DISPATCH_PROPERTYPUT, {3L});
    invoke(para.Get(), L"SpaceBefore", DISPATCH_PROPERTYPUT, {0L});
    invoke(para.Get(), L"SpaceAfter", DISPATCH_PROPERTYPUT, {0L});
    invoke(para.Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {2L});
    auto locate = [&](const wchar_t* source) {
        auto found = findLiteral(ScanRangeContext{document}, start, integer(get(document, L"Content").Get(), L"End"), source);
        require(found != nullptr, "Find fixture marker"); return found;
    };
    invoke(get(locate(L"split_1").Get(), L"Font").Get(), L"Bold", DISPATCH_PROPERTYPUT, {true});
    invoke(get(locate(L"split_2").Get(), L"Font").Get(), L"Italic", DISPATCH_PROPERTYPUT, {true});
    for (const auto marker : {L"#FIELD#", L"#CROSS_FIELD#"}) {
        auto range = locate(marker);
        invoke(range.Get(), L"Text", DISPATCH_PROPERTYPUT, {L""});
        invoke(range.Get(), L"Collapse", DISPATCH_METHOD, {1L});
        auto field = asObject(invoke(get(document, L"Fields").Get(), L"Add", DISPATCH_METHOD,
            {range.Get(), -1L, L"QUOTE \"$field_hidden$\"", false}));
        invoke(field.Get(), L"Update", DISPATCH_METHOD);
    }
    auto mathRange = asObject(invoke(get(document, L"OMaths").Get(), L"Add", DISPATCH_METHOD, {locate(L"$omml_only$").Get()}));
    invoke(get(mathRange.Get(), L"OMaths").Get(), L"BuildUp", DISPATCH_METHOD);
    auto revision = locate(L"REVISION: ");
    invoke(revision.Get(), L"Collapse", DISPATCH_METHOD, {0L});
    invoke(document, L"TrackRevisions", DISPATCH_PROPERTYPUT, {true});
    invoke(revision.Get(), L"InsertAfter", DISPATCH_METHOD, {L"$revision_only$"});
    invoke(document, L"TrackRevisions", DISPATCH_PROPERTYPUT, {false});
    auto control = asObject(invoke(get(document, L"ContentControls").Get(), L"Add", DISPATCH_METHOD, {1L, locate(L"#CONTROL#").Get()}));
    invoke(get(control.Get(), L"Range").Get(), L"Text", DISPATCH_PROPERTYPUT, {L"$control_only$"});
    invoke(control.Get(), L"LockContents", DISPATCH_PROPERTYPUT, {true});
    auto note = locate(L"#NOTE#");
    invoke(note.Get(), L"Text", DISPATCH_PROPERTYPUT, {L""});
    invoke(note.Get(), L"Collapse", DISPATCH_METHOD, {1L});
    invoke(get(document, L"Footnotes").Get(), L"Add", DISPATCH_METHOD, {note.Get(), Value(), L"$footnote_only$"});
    auto section = collectionItem(get(document, L"Sections").Get(), 1);
    auto header = collectionItem(get(section.Get(), L"Headers").Get(), 1);
    invoke(get(header.Get(), L"Range").Get(), L"Text", DISPATCH_PROPERTYPUT, {L"$header_only$"});
    auto picture = locate(L"#PICTURE#");
    invoke(picture.Get(), L"Text", DISPATCH_PROPERTYPUT, {L""});
    invoke(picture.Get(), L"Collapse", DISPATCH_METHOD, {1L});
    const auto picturePath = directory + L"\\scan-picture.png";
    invoke(get(document, L"InlineShapes").Get(), L"AddPicture", DISPATCH_METHOD, {picturePath.c_str(), false, true, picture.Get()});
    auto picturePara = get(picture.Get(), L"ParagraphFormat");
    invoke(picturePara.Get(), L"Alignment", DISPATCH_PROPERTYPUT, {1L});
    invoke(picturePara.Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {0L});
    auto pictureCaption = locate(L"图1-1 普通图片边界");
    invoke(get(pictureCaption.Get(), L"Font").Get(), L"Size", DISPATCH_PROPERTYPUT, {10.5});
    invoke(get(pictureCaption.Get(), L"ParagraphFormat").Get(), L"Alignment", DISPATCH_PROPERTYPUT, {1L});
    invoke(get(pictureCaption.Get(), L"ParagraphFormat").Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {0L});
    auto caption = locate(L"表1-2 扫描单元格边界");
    invoke(get(caption.Get(), L"Font").Get(), L"Size", DISPATCH_PROPERTYPUT, {10.5});
    invoke(get(caption.Get(), L"ParagraphFormat").Get(), L"Alignment", DISPATCH_PROPERTYPUT, {1L});
    invoke(get(caption.Get(), L"ParagraphFormat").Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {0L});
    const long end = integer(get(document, L"Content").Get(), L"End") - 1;
    auto table = asObject(invoke(get(document, L"Tables").Get(), L"Add", DISPATCH_METHOD, {docRange(document, end, end).Get(), 3L, 2L}));
    const wchar_t* cells[][2] = {{L"边界左侧", L"边界右侧"}, {L"$cross_cell", L"suffix$"},
                                {L"Cell $cell_1$ / $$cell_2$$", L"Invalid ${cell$ and blank $ $"}};
    for (long row = 1; row <= 3; ++row) for (long column = 1; column <= 2; ++column)
        invoke(cellRange(table.Get(), row, column).Get(), L"Text", DISPATCH_PROPERTYPUT, {cells[row - 1][column - 1]});
    auto tableRange = get(table.Get(), L"Range");
    invoke(get(tableRange.Get(), L"Font").Get(), L"Size", DISPATCH_PROPERTYPUT, {10.5});
    invoke(get(tableRange.Get(), L"ParagraphFormat").Get(), L"Alignment", DISPATCH_PROPERTYPUT, {1L});
    invoke(get(tableRange.Get(), L"ParagraphFormat").Get(), L"CharacterUnitFirstLineIndent", DISPATCH_PROPERTYPUT, {0L});
    invoke(get(table.Get(), L"Borders").Get(), L"Enable", DISPATCH_PROPERTYPUT, {0L});
    blackBorder(table.Get(), -1, 12); blackBorder(table.Get(), -3, 12);
    blackBorder(collectionItem(get(table.Get(), L"Rows").Get(), 1).Get(), -3, 6);
    invoke(document, L"Save", DISPATCH_METHOD);
    // 额外自建整篇只读保护副本，验收跳过状态与未知计数，不解锁原文件。
    const auto protectedPath = directory + L"\\scan-protected.docx";
    invoke(document, L"SaveAs2", DISPATCH_METHOD, {protectedPath.c_str(), 12L});
    invoke(document, L"Protect", DISPATCH_METHOD, {3L, true, L""});
    invoke(document, L"Save", DISPATCH_METHOD);
    invoke(document, L"Close", DISPATCH_METHOD, {0L});
    word.document.Reset();
}

static int scanMain(int argc, wchar_t** argv) {
    setvbuf(stdout, nullptr, _IONBF, 0);
    if (argc != 4 || (std::wstring(argv[1]) != L"--scan" && std::wstring(argv[1]) != L"--self-test")) {
        std::puts("usage: docx-math-scan --scan <input.docx> <new-report.json> | --self-test <base.docx> <new-task-directory>"); return 2;
    }
    const bool selfTest = std::wstring(argv[1]) == L"--self-test";
    std::wstring base, destination, report;
    const HRESULT initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(initialized)) return 1;
    int code = 1;
    try {
        // Word 的相对路径解析可能使用自己的文档目录，先统一为本进程的绝对路径。
        base = absolutePath(argv[2]); destination = absolutePath(argv[3]);
        report = selfTest ? destination + L"\\scan.json" : destination;
        require(base.size() > 5 && _wcsicmp(base.c_str() + base.size() - 5, L".docx") == 0, "DOCX input required");
        requireDocxZip(base);
        require(GetFileAttributesW(report.c_str()) == INVALID_FILE_ATTRIBUTES, "Report already exists");
        if (selfTest) {
            WordSession word;
            auto documents = get(word.application.Get(), L"Documents");
            word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD)); word.verifyOwnership();
            invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
            invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
            invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L}); word.disableLinkUpdates();
            invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L}); word.document.Reset();
            makeFixture(word, base, destination);
            word.shutdown();
        }
        const std::wstring input = selfTest ? destination + L"\\scan-fixture.docx" : base;
        scanReadOnlyFile(input, report);
        std::puts("SCAN_DOCUMENT_PASSED"); code = 0;
    } catch (const std::exception& error) {
        std::printf("SCAN_DOCUMENT_FAILED: %s\n", error.what());
    }
    CoUninitialize(); return code;
}

// 扫描不加载 MathType；其有界恢复进程只恢复本任务修改过的 Word 链接选项。
static int recoverScanOptions(const wchar_t* originalName) {
    const auto initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED); if (FAILED(initialized)) return 1;
    int code = 1;
    try {
        TaskMapping original(originalName, false); require(!original.state->pendingPrefs, "Scan never sets MathType preferences");
        WordSession word; word.document = asObject(invoke(get(word.application.Get(), L"Documents").Get(), L"Add", DISPATCH_METHOD)); word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false}); invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        if (original.state->linksChanged) {
            invoke(get(word.application.Get(), L"Options").Get(), L"UpdateLinksAtOpen", DISPATCH_PROPERTYPUT, {original.state->originalLinks});
            InterlockedExchange(&original.state->linksChanged, 0); std::puts("WORD_OPTIONS_RECOVERY_PASSED");
        }
        word.shutdown(); code = 0;
    } catch (const std::exception& error) { taskError(error.what()); }
    CoUninitialize(); return code;
}
int wmain(int argc, wchar_t** argv) {
    if (argc > 3 && std::wstring(argv[1]) == L"--task-run") {
        try { TaskMapping mapping(argv[2], false); taskShared = mapping.state;
            const int code = std::wstring(argv[3]) == L"--worker-recover" && argc == 5 ? recoverScanOptions(argv[4]) : scanMain(argc - 2, argv + 2);
            taskShared = nullptr; return code;
        } catch (const std::exception& error) { taskError(error.what()); return 1; }
    }
    return superviseTask(argc, argv, false);
}
