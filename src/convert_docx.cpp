// 任务入口：锁定原件、创建副本、组合正文与额外区域阶段，核验后发布。
#include "body_conversion.h"
#include "equation_styles.h"
#include "app_version.h"
#include "docx_package.h"
#include "conversion_timing.h"

static void writeNewText(const std::wstring& path, const std::string& contents) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(file != INVALID_HANDLE_VALUE, "Create new batch report"); DWORD written{};
    taskTrackFile(path, file, false);
    const bool saved = contents.size() <= MAXDWORD && WriteFile(file, contents.data(), static_cast<DWORD>(contents.size()), &written, nullptr) && written == contents.size();
    const bool flushed = FlushFileBuffers(file) != 0, closed = CloseHandle(file) != 0;
    if (!saved || !flushed || !closed) { DeleteFileW(path.c_str()); throw std::runtime_error("Write complete batch report"); }
}
static std::string outcomeReport(const std::wstring& input, const std::wstring& output, const std::wstring& eqp,
                                 const BodyOptions& options,const BodyResult& body) {
    const auto& scan=body.scan;const auto& outcomes=body.outcomes;const auto& experimental=body.experimental;const auto& issues=body.issues;const bool fullReadback=options.selfTest;
    size_t ready{}, succeeded{}, failed{}, failedUnverified{}, unattempted{};
    for (size_t index = 0; index < outcomes.size(); ++index) {
        ready += scan.items[index].ready; succeeded += outcomes[index].status == L"SUCCEEDED";
        failed += failedOutcome(outcomes[index]);failedUnverified+=outcomes[index].status==L"FAILED_UNVERIFIED"; unattempted += outcomes[index].status == L"NOT_ATTEMPTED";
    }
    const auto existing=std::count_if(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.existingDisplay||item.existingReference;});
    const size_t rejected=outcomes.size()-ready-static_cast<size_t>(existing)+experimental.rejected;
    ready+=experimental.succeeded+experimental.failed+experimental.unattempted;succeeded+=experimental.succeeded;failed+=experimental.failed;unattempted+=experimental.unattempted;
    failedUnverified+=experimental.failedUnverified;
    const auto status = failed || rejected || unattempted || !issues.empty() || experimental.partial() ? L"PARTIAL" : L"COMPLETE";
    std::string json = "{\n  \"SchemaVersion\":1, \"Application\":\"DocxMathTypeBatch\", \"ApplicationVersion\":\"" DMTB_VERSION "\", \"CliContractVersion\":" DMTB_CLI_CONTRACT ", \"Scope\":"+quote(experimental.enabled?L"BODY_WITH_EXPERIMENTAL_EXTENSIONS":L"MAIN_STORY_BODY_ONLY")+", \"Positions\":"+quote(options.scan.fast?L"WORD_WORKING_COPY_UNVERIFIED":L"WORD_SNAPSHOTS_BEFORE_AND_AFTER")+",\n";
    json+="  \"FastConversion\":"+std::string(options.scan.fast?"true":"false")+", \"FailedUnverified\":"+std::to_string(failedUnverified)+",\n";
    json+="  \"ConvertOmml\":"+std::string(!options.reverse&&options.scan.convertOmml?"true":"false")+",\n";
    json += "  \"EquationMode\":"+quote(options.scan.scope==EquationScope::Inline?L"Inline":options.scan.scope==EquationScope::Display?L"Display":L"All")+", \"NumbersAdded\":"+std::to_string(body.numbering.inserted)+",\n";
    const auto displays=std::count_if(outcomes.begin(),outcomes.end(),[](const ItemOutcome& item){return item.status==L"EXISTING_FORMATTED";});
    json += "  \"NumbersUpdated\":"+std::to_string(body.numbering.updated)+", \"NumberReferencesUpdated\":"+std::to_string(body.numbering.referencesUpdated)+", \"ExistingDisplaysFormatted\":"+std::to_string(displays)+",\n";
    json += "  \"InlineSpacing\":"+quote(spacingArgument(options.inlineSpacing))+", \"DisplaySpacing\":"+quote(spacingArgument(options.displaySpacing))+", \"NumberFormat\":"+quote(options.numbering.format)+",\n";
    json += "  \"SavedNativeBatch\":"+std::string(body.savedNative?"true":"false")+",\n";
    json += "  \"Input\":" + quote(input) + ", \"Output\":" + quote(output) + ", \"Preferences\":" + quote(eqp) + ",\n";
    if (scan.reverse) {
        long exported{}, normalized{};
        for (size_t index = 0; index < outcomes.size(); ++index) if (outcomes[index].status == L"SUCCEEDED") {
            if (scan.items[index].delimiter == L"MATH_TYPE") ++exported; else ++normalized;
        }
        json += "  \"Direction\":\"REVERSE\", \"DisplayDelimiter\":" + quote(eqp) +
                ", \"ExportedMathType\":" + std::to_string(exported) + ", \"NormalizedLatex\":" + std::to_string(normalized) + ",\n";
    }
    json += "  \"Status\":" + quote(status) + ", \"Targets\":" + std::to_string(ready) + ", \"Succeeded\":" + std::to_string(succeeded) +
            ", \"FailedPreserved\":" + std::to_string(failed-failedUnverified) + ", \"NotAttempted\":" + std::to_string(unattempted) +
            ", \"RejectedPreserved\":" + std::to_string(rejected) + ", \"ExistingMathType\":" + std::to_string(scan.existingMathType) + ",\n";
    json += "  \"Validation\":\"" + std::string(options.scan.fast?"SKIPPED_EXPERIMENTAL_FAST":scan.reverse ? "NATIVE_TRANSLATOR_CONTENT_RESIDUALS_OMML_PACKAGE_AND_WORD_REOPEN" : "SAVED_NATIVE_BYTES_STYLE_RESIDUALS_PACKAGE_AND_WORD_REOPEN") + "\", \"MathematicalEquivalenceVerified\":false,\n";
    json += "  \"FullDevelopmentReadback\":" + std::string(fullReadback&&!options.scan.fast?"true":"false") + ", \"ReopenedEquationActivations\":" + std::to_string(fullReadback&&!options.scan.fast&&!scan.reverse?succeeded-experimental.succeeded:0) + ",\n";
    json += "  \"PostConversionScans\":"+std::to_string(options.scan.fast?0:(fullReadback?2:1)+(experimental.enabled&&experimental.skipped.empty()?2:0))+",\n";
    if(options.scan.fast)json+="  \"OriginalFileBytesUnchanged\":null, \"ExistingNativeStreamsUnchanged\":null, \"ExistingMediaBytesPreserved\":null, \"RetainedNativeStreamsUnchanged\":null, \"RetainedMediaBytesPreserved\":null, \"OMMLUnchanged\":null, \"WordReopenVerified\":false,\n  \"Items\":[\n";
    else json += scan.reverse ? "  \"OriginalFileBytesUnchanged\":true, \"RetainedNativeStreamsUnchanged\":true, \"RetainedMediaBytesPreserved\":true, \"OMMLUnchanged\":true, \"WordReopenVerified\":true,\n  \"Items\":[\n" :
                          "  \"OriginalFileBytesUnchanged\":true, \"ExistingNativeStreamsUnchanged\":true, \"ExistingMediaBytesPreserved\":true, \"WordReopenVerified\":true,\n  \"Items\":[\n";
    size_t written{};for (size_t index = 0; index < outcomes.size(); ++index) {
        const auto& item = scan.items[index]; const auto& result = outcomes[index];
        if(item.existingDisplay||item.existingReference)continue;
        const auto fullSize=result.native.empty()||scan.reverse?std::string():inspectStyle(result.native).sizes.at(0).second;
        json += (written++?",\n":"")+std::string("    {\"Id\":") + std::to_string(index + 1) + ",\"InitialStart\":" + std::to_string(item.span.start) + ",\"InitialEnd\":" + std::to_string(item.span.end) +
            ",\"FinalStart\":" + std::to_string(result.finalSpan.start) + ",\"FinalEnd\":" + std::to_string(result.finalSpan.end) +
            ",\"Region\":\"" + (item.row ? "TABLE" : "BODY") + std::string("\",\"Row\":") + std::to_string(item.row) + ",\"Column\":" + std::to_string(item.column) +
            ",\"Display\":" + (item.display ? "true" : "false") + ",\"Status\":" + quote(result.status) + ",\"Reason\":" + quote(result.reason) +
            ",\"NativeFullSize\":" + (fullSize.empty()?std::string("null"):quote(std::wstring(fullSize.begin(),fullSize.end()))) +
            ",\"AddedParagraphs\":" + std::to_string(result.addedParagraphs) + ",\"Source\":" + quote(item.source) +sourcePositionJson(item.position)+
            (scan.reverse ? ",\"Kind\":" + quote(item.delimiter == L"MATH_TYPE" ? L"MATH_TYPE" : L"LATEX") + ",\"OutputText\":" + quote(result.outputText) : ",\"Kind\":"+quote(item.ommlOrdinal?L"OMML":L"LATEX")+",\"FormattingApplied\":"+(result.status==L"SUCCEEDED"?"true":"false")) + "}";
    }
    for(const auto& item:experimental.items)json+=(written++?",\n":"")+item.json();
    json+="\n  ],\n  \"Issues\":[";size_t issueIndex{};for(const auto& issue:issues)json+=(issueIndex++?",":"")+issue.json(options.scan.fast);for(const auto& issue:experimental.issues)json+=(issueIndex++?",":"")+issue.json(options.scan.fast);
    json+="],\n  \"LayoutWarningsChecked\":"+std::string(!options.scan.fast&&!options.reverse?"true":"false")+",\n  \"LayoutWarningScope\":\"BODY_STANDALONE_DISPLAY_WIDTH_SINGLE_OR_EQUAL_COLUMNS_AND_CONVERTED_EQUATION_EXACT_LINE_SPACING\",\n  \"Warnings\":[";
    size_t warningIndex{};for(const auto& warning:body.warnings)json+=(warningIndex++?",":"")+warning.json();
    for(const auto& warning:experimental.warnings)json+=(warningIndex++?",":"")+warning.json();
    return json + "],\n  \"Experimental\":"+experimental.json(options.scan.fast)+"\n}\n";
}

int convertMain(int argc, wchar_t** argv) {
    setvbuf(stdout, nullptr, _IONBF, 0);
    const std::wstring mode = argc > 1 ? argv[1] : L"";
    const bool reverse = mode == L"--reverse" || mode == L"--worker-reverse" || mode == L"--self-test-reverse" ||
                         mode == L"--self-test-reverse-stop" || mode == L"--self-test-reverse-abort";
    const bool guiWorker = mode == L"--worker-convert" || mode == L"--worker-reverse";
    BodyOptions options;options.reverse=reverse;options.scan.progress=guiWorker;
    if ((guiWorker ? argc < 7 : argc < 6) ||
        (mode != L"--convert" && !reverse && !guiWorker && mode != L"--self-test" && mode != L"--self-test-abort" && mode != L"--self-test-stop")) {
        std::puts("usage: docx-math-convert --convert <input.docx> <new-output.docx> <new-report.json> <paragraph|size=pt|task.eqp> [MathType-directory] [--scope=all|inline|display] [--numbers]"); return 2;
    }
    const bool selfTest = mode == L"--self-test" || mode == L"--self-test-abort" || mode == L"--self-test-reverse" || mode == L"--self-test-reverse-abort";
    const bool abortTest = mode == L"--self-test-abort" || mode == L"--self-test-reverse-abort";
    const bool stopTest = mode == L"--self-test-stop" || mode == L"--self-test-reverse-stop";
    options.selfTest=selfTest;options.abortTest=abortTest;options.stopTest=stopTest;
    const auto initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED); if (FAILED(initialized)) return 1;
    std::wstring candidate, stagedReport, output, sizePreferencePath;
    bool sizePreferenceOwned{};
    bool candidateOwned = false, reportOwned = false, outputOwned = false, previewOwned = false, publishedPreviewOwned = false;
    HANDLE sourceLock = INVALID_HANDLE_VALUE;
    int code = 1;
    ConversionTimings timings;
    try {
        std::puts("UI_PHASE PREPARING");
        std::wstring installProblem;
        std::wstring explicitInstall;ExperimentalOptions experimentalOptions;
        for(int i=guiWorker?7:6;i<argc;++i) {
            const std::wstring argument=argv[i];
            if(argument==L"--scope=all")options.scan.scope=EquationScope::All;
            else if(argument==L"--experimental-fast")options.scan.fast=true;
            else if(argument==L"--skip-omml")options.scan.convertOmml=false;
            else if(argument==L"--experimental-tables")experimentalOptions.tables=true;
            else if(argument==L"--experimental-textboxes")experimentalOptions.textboxes=true;
            else if(argument==L"--experimental-headers")experimentalOptions.headers=true;
            else if(argument==L"--experimental-notes")experimentalOptions.notes=true;
            else if(argument==L"--scope=inline")options.scan.scope=EquationScope::Inline;
            else if(argument==L"--scope=display")options.scan.scope=EquationScope::Display;
            else if(argument==L"--numbers")options.addRightNumbers=true;
            else if(argument==L"--number-custom")options.numbering.custom=true;
            else if(argument.rfind(L"--inline-spacing=",0)==0)options.inlineSpacing=parseSpacing(argument.substr(17));
            else if(argument.rfind(L"--display-spacing=",0)==0)options.displaySpacing=parseSpacing(argument.substr(18));
            else if(argument.rfind(L"--number-format=",0)==0)options.numbering.format=argument.substr(16);
            else if(argument.rfind(L"--number-start=",0)==0){const auto value=settingNumber(argument.substr(15),1,999999);require(value==std::floor(value),"Integer start required");options.numbering.start=static_cast<long>(value);}
            else if(argument.rfind(L"--number-chapter=",0)==0){const auto value=settingNumber(argument.substr(17),1,999999);require(value==std::floor(value),"Integer chapter required");options.numbering.chapter=static_cast<long>(value);}
            else if(argument.rfind(L"--number-section=",0)==0){const auto value=settingNumber(argument.substr(17),1,999999);require(value==std::floor(value),"Integer section required");options.numbering.section=static_cast<long>(value);}
            else {require(argument.rfind(L"--",0)!=0&&explicitInstall.empty(),"Unknown or duplicate conversion argument");explicitInstall=absolutePath(argv[i]);}
        }
        require(!reverse||!options.addRightNumbers,"Reverse cannot add equation numbers");
        require(!reverse||(!options.inlineSpacing.mode&&!options.displaySpacing.mode),"Reverse preserves paragraph spacing");
        if(options.addRightNumbers)numberParts(options.numbering,options.numbering.start);
        const auto install = explicitInstall.empty()?findMathTypeInstall(installProblem):explicitInstall;
        if (!install.empty()) installProblem=mathTypeInstallProblem(install);
        require(installProblem.empty(),utf8(installProblem).c_str());
        std::printf("MATHTYPE_DIRECTORY %s\n",wireHex(utf8(install)).c_str());
        const auto stopFile = guiWorker ? absolutePath(argv[6]) : L"";
        options.stopFile=stopFile;
        const auto input = absolutePath(argv[2]), report = absolutePath(argv[4]);
        const std::wstring styleArgument=argv[5];
        const bool followSize=!reverse&&styleArgument==L"paragraph", fixedSize=!reverse&&styleArgument.rfind(L"size=",0)==0;
        options.scan.capturePoints=followSize;
        double requestedPoints{};
        if(fixedSize){wchar_t* end{};requestedPoints=wcstod(styleArgument.c_str()+5,&end);require(end!=styleArgument.c_str()+5&&!*end&&std::isfinite(requestedPoints)&&requestedPoints>=1&&requestedPoints<=256,"Specified size must be 1..256 pt");}
        const auto eqp = reverse||followSize||fixedSize?styleArgument:absolutePath(argv[5]);
        options.requestedPoints=requestedPoints;
        require(!reverse || eqp == L"bracket" || eqp == L"dollar", "Reverse display delimiter must be bracket or dollar");
        output = absolutePath(argv[3]);
        require(input.size() > 5 && _wcsicmp(input.c_str() + input.size() - 5, L".docx") == 0 &&
                output.size() > 5 && _wcsicmp(output.c_str() + output.size() - 5, L".docx") == 0, "DOCX input/output required");
        require(_wcsicmp(input.c_str(), output.c_str()) != 0 && _wcsicmp(output.c_str(), report.c_str()) != 0, "Distinct input/output/report paths required");
        require(GetFileAttributesW(output.c_str()) == INVALID_FILE_ATTRIBUTES, "Output already exists");
        require(GetFileAttributesW(report.c_str()) == INVALID_FILE_ATTRIBUTES, "Report already exists");
        if (selfTest) require(GetFileAttributesW((output + L".pdf").c_str()) == INVALID_FILE_ATTRIBUTES, "Test preview already exists");
        requireDocxZip(input);
        if (!reverse&&!followSize&&!fixedSize) require(GetFileAttributesW(eqp.c_str()) != INVALID_FILE_ATTRIBUTES, "Task EQP required");
        sourceLock = CreateFileW(input.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if(sourceLock==INVALID_HANDLE_VALUE)throw std::runtime_error("Lock source against concurrent writes during copy and conversion (Win32="+std::to_string(GetLastError())+")");
        // 快速模式只读取排版所需的样式，不建立原生资产与关系核验基线。
        const auto originalPackage = inspectPackage(input,options.scan.fast?nullptr:"INPUT_PARTS",options.scan.fast,options.scan.progress);
        // 文件字节只读对比；输入同时保持只读共享锁，不能在扫描后被其他程序改写。
        auto readOriginal = [&](const std::vector<char>* expected=nullptr) {
            // 复用始终持有的原件句柄；末尾仍逐字节核对，只省去重新打开、重复扩容和第二份整文件缓冲。
            LARGE_INTEGER size{},zero{};require(GetFileSizeEx(sourceLock,&size)&&SetFilePointerEx(sourceLock,zero,nullptr,FILE_BEGIN),"Read locked original size and rewind");
            if(expected)require(static_cast<size_t>(size.QuadPart)==expected->size(),"Original DOCX size unchanged");
            std::vector<char> bytes(expected?0:static_cast<size_t>(size.QuadPart));char buffer[64*1024];
            for(size_t offset=0;offset<static_cast<size_t>(size.QuadPart);){
                const auto count=static_cast<DWORD>(std::min(static_cast<size_t>(size.QuadPart)-offset,sizeof(buffer)));DWORD read{};
                require(ReadFile(sourceLock,expected?buffer:bytes.data()+offset,count,&read,nullptr)&&read==count,"Complete original DOCX read");
                if(expected)require(std::equal(buffer,buffer+read,expected->begin()+offset),"Original DOCX bytes unchanged");offset+=read;
            }
            return bytes;
        };
        if(!options.scan.fast)std::puts("TIMING_STAGE SOURCE_SNAPSHOT");
        const auto originalBytes = options.scan.fast?std::vector<char>():readOriginal();
        std::puts("TIMING_STAGE COPY_AND_LINKS");
        const auto suffix = L".work-" + std::to_wstring(GetCurrentProcessId()) + L"-" + std::to_wstring(GetTickCount64());
        candidate = output + suffix + L".docx"; stagedReport = report + suffix + L".tmp";
        sizePreferencePath=candidate+L".eqp";
        taskCopy(input, candidate); candidateOwned = true;
        const auto linkUpdates=setDocxLinkUpdates(candidate);
        std::puts("UI_PHASE OPENING");
        WordSession word;
        auto documents = get(word.application.Get(), L"Documents");
        word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD, {Value(), false, Value(), false})); word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
        invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
        invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L}); word.document.Reset();
        taskTrackWordOwner(candidate);
        word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD,
            {candidate.c_str(), false, false, false, Value(), Value(), false, Value(), Value(), Value(), Value(), false}));
        require(!boolProperty(word.document.Get(), L"ReadOnly"), "Task copy must be writable");
        MathTypeApi api(install);
        auto body=convertBody(word,api,input,candidate,eqp,originalPackage,options,experimentalOptions,sizePreferencePath,sizePreferenceOwned,previewOwned,timings);
        auto& scan=body.scan;auto& outcomes=body.outcomes;auto& issues=body.issues;auto& experimental=body.experimental;
        const auto& experimentalSkip=body.experimentalSkip;const auto& afterPackage=body.package;
        const auto total=body.total,completed=body.completed,succeeded=body.succeeded;
        if(!options.scan.fast&&experimentalOptions.enabled()){
            if(experimentalSkip.empty()){
                // 正文阶段已核验的输出包直接作为实验基线，避免实验分支再整包读取一次已有公式。
                experimental=convertExperimental(word,api,candidate,experimentalOptions,reverse,eqp==L"dollar",scan.items.size()+1,afterPackage,selfTest,options.scan,timings,body.experimentalSourceLocations);
                // 表格前后的正文片段已核验不变；使用各表格的累计位移，不把多个累计值相加。
                for(auto& outcome:outcomes){
                    const auto after=std::upper_bound(experimental.bodyPositionShifts.begin(),experimental.bodyPositionShifts.end(),outcome.finalSpan.start,
                        [](long position,const BodyPositionShift& shift){return position<shift.boundary;});
                    if(after!=experimental.bodyPositionShifts.begin()){const long offset=(after-1)->offset;outcome.finalSpan.start+=offset;outcome.finalSpan.end+=offset;}
                }
                word.shutdown();if(taskShared)InterlockedExchange(&taskShared->mathReady,0);
                if(api.initialized){require(api.term()==0,"Terminate experimental MathType API");api.initialized=false;}
            }else experimental.skipped=experimentalSkip;
        }
        // 链接更新方式在正文与实验阶段共用同一次设置，实验保存后再统一恢复，避免重复读写整包。
        setDocxLinkUpdates(candidate,&linkUpdates);
        if(!options.scan.fast){std::puts("TIMING_STAGE SOURCE_COMPARE");readOriginal(&originalBytes);}
        require(CloseHandle(sourceLock) != 0, "Release original DOCX lock"); sourceLock = INVALID_HANDLE_VALUE;
        std::puts("UI_PHASE PUBLISHING");
        taskTestBoundary(L"BEFORE_PUBLISH");taskReservePrivatePath(stagedReport);
        writeNewText(stagedReport, outcomeReport(input, output, eqp, options,body)); reportOwned = true;
        taskMove(candidate, output);
        taskTestBoundary(L"AFTER_OUTPUT_MOVE");
        candidateOwned = false; outputOwned = true;
        if (selfTest&&!options.scan.fast) {
            taskMove(candidate + L".pdf", output + L".pdf");
            previewOwned = false; publishedPreviewOwned = true;
        }
        taskMove(stagedReport, report); reportOwned = false;
        taskCommit();
        taskTestBoundary(L"AFTER_COMMIT");
        outputOwned = false; publishedPreviewOwned = false;
        const auto failed = std::count_if(outcomes.begin(), outcomes.end(), failedOutcome);
        const auto rejected=std::count_if(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return !item.ready&&!item.existingDisplay&&!item.existingReference;});
        for(const auto& issue:issues)emitProcessingIssue(issue);
        for(const auto& issue:experimental.issues)emitProcessingIssue(issue);
        // 仅在最终文件发布后发送提醒，重建或撤回的中间结果不会留在界面中。
        std::printf("UI_LAYOUT_CHECKED %d\n",!options.scan.fast&&!options.reverse?1:0);
        const auto emitWarning=[](const LayoutWarning& warning){
            std::printf("UI_LAYOUT_WARNING %zu %s %s %s %s %s\n",warning.id,utf8(warning.code).c_str(),wireHex(utf8(warning.region)).c_str(),
                wireHex(utf8(warning.position.location)).c_str(),diagnosticExcerpt(warning.source).c_str(),wireHex(utf8(warningMessage(warning))).c_str());
        };
        for(const auto& warning:body.warnings)emitWarning(warning);
        for(const auto& warning:experimental.warnings)emitWarning(warning);
        std::printf("UI_CONVERT_SUMMARY %s %zu %zu %zu %zu %zu\n", failed || rejected || !issues.empty() || experimental.partial() ? "PARTIAL" : "COMPLETE",
                     static_cast<size_t>(total)+experimental.succeeded+experimental.failed+experimental.unattempted,static_cast<size_t>(succeeded)+experimental.succeeded,
                     static_cast<size_t>(failed)+experimental.failed,static_cast<size_t>(total)-completed+experimental.unattempted,static_cast<size_t>(rejected)+experimental.rejected);
        std::puts("BATCH_CONVERSION_PASSED"); code = 0;
    } catch (const std::exception& error) {
        std::printf("UI_ERROR %s\n", wireHex(error.what()).c_str());
        std::printf("BATCH_CONVERSION_FAILED: %s\n", error.what());
    }
    timings.report();
    if (sourceLock != INVALID_HANDLE_VALUE) CloseHandle(sourceLock);
    if(sizePreferenceOwned&&!DeleteFileW(sizePreferencePath.c_str()))std::puts("OWNED_SIZE_PREFERENCE_CLEANUP_FAILED");
    // 服务断连时先释放 COM 公寓，避免 OLE 存储的延迟释放阻止工作副本清理。
    const auto comReleaseStarted=GetTickCount64();
    CoUninitialize();
    std::printf("COM_UNINITIALIZE_MS=%llu\n",GetTickCount64()-comReleaseStarted);
    // 仅删除本次新建、尚未交付的工作副本/报告，不触碰任何已有用户文件。
    if (candidateOwned && !DeleteFileW(candidate.c_str())) std::printf("OWNED_BATCH_COPY_CLEANUP_FAILED WIN32=%lu\n", GetLastError());
    if (reportOwned && !DeleteFileW(stagedReport.c_str())) std::puts("OWNED_BATCH_REPORT_CLEANUP_FAILED");
    if (outputOwned && !DeleteFileW(output.c_str())) std::puts("OWNED_BATCH_OUTPUT_ROLLBACK_FAILED");
    if (previewOwned && !DeleteFileW((candidate + L".pdf").c_str())) std::puts("OWNED_BATCH_PREVIEW_CLEANUP_FAILED");
    if (publishedPreviewOwned && !DeleteFileW((output + L".pdf").c_str())) std::puts("OWNED_BATCH_PREVIEW_ROLLBACK_FAILED");
    return code;
}

// 有界恢复进程只消费未知的 NEXT_EQN，再确认自然默认和全局偏好保持，绝不覆盖用户默认。
int recoverPreferencesMain(const std::wstring& originalName) {
    const auto initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED); if (FAILED(initialized)) return 1;
    int code = 1;
    try {
        TaskMapping original(originalName, false); auto& originalState = *original.state;
        WordSession word; auto documents = get(word.application.Get(), L"Documents");
        word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD)); word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false}); invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
        if (originalState.linksChanged) {
            invoke(get(word.application.Get(), L"Options").Get(), L"UpdateLinksAtOpen", DISPATCH_PROPERTYPUT, {originalState.originalLinks});
            InterlockedExchange(&originalState.linksChanged, 0); std::puts("WORD_OPTIONS_RECOVERY_PASSED");
        }
        if (originalState.pendingPrefs) {
            const auto& record = originalState.recovery;
            require(record.nativeLength && record.nativeLength <= TaskNativeCapacity, "Recover only recorded preference baseline");
            MathTypeApi api(record.install); require(api.init(0, 10) > 0, "Initialize preference recovery API"); api.initialized = true;
            taskMathTypeConnected();
            require(readDefaultPrefs(api) == record.defaults, "User defaults changed; refuse preference recovery");
            const auto consumed = inspectStyle(emptyEquation(api, word));
            if (originalState.expectedPendingLength) {
                const auto expected = inspectStyle(std::vector<unsigned char>(originalState.expectedPendingNative, originalState.expectedPendingNative + originalState.expectedPendingLength));
                const auto natural = inspectStyle(std::vector<unsigned char>(record.native, record.native + record.nativeLength));
                require((consumed.inlineEquation == expected.inlineEquation && consumed.sizes == expected.sizes && consumed.styles == expected.styles) ||
                    (consumed.inlineEquation == natural.inlineEquation && consumed.sizes == natural.sizes && consumed.styles == natural.styles), "Recovered pending equation has task or natural style");
                std::puts("PENDING_STYLE_CONSUMPTION_VERIFIED");
            }
            requireStyle(inspectStyle(emptyEquation(api, word)), inspectStyle(std::vector<unsigned char>(record.native, record.native + record.nativeLength)));
            require(readDefaultPrefs(api) == record.defaults, "Global preferences unchanged after recovery");
            if (taskShared) InterlockedExchange(&taskShared->mathReady, 0);
            require(api.term() == 0, "Terminate preference recovery API"); api.initialized = false;
            InterlockedExchange(&originalState.pendingPrefs, 0); std::puts("PREFERENCE_RECOVERY_PASSED");
        }
        word.shutdown(); code = 0;
    } catch (const std::exception& error) { taskError(error.what()); std::printf("PREFERENCE_RECOVERY_FAILED: %s\n", error.what()); }
    CoUninitialize(); return code;
}

#ifndef DOCX_MATH_GUI
// 独立批处理入口保留；GUI 链接同一实现，工作模式不复制转换流程。
int wmain(int argc, wchar_t** argv) {
    if (argc > 3 && std::wstring(argv[1]) == L"--task-run") {
        try { TaskMapping mapping(argv[2], false); taskShared = mapping.state;
            const int code = std::wstring(argv[3]) == L"--worker-recover" && argc == 5 ? recoverPreferencesMain(argv[4]) : convertMain(argc - 2, argv + 2);
            taskShared = nullptr; return code;
        } catch (const std::exception& error) { taskError(error.what()); return 1; }
    }
    return superviseTask(argc, argv);
}
#endif
