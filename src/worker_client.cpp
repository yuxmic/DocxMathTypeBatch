// 保留原有独立进程和句柄白名单；本模块不创建 Word 或绘制资源。
#include "worker_client.h"
#include "word_automation.h"
#include "worker_protocol.h"
#include <climits>
#include <sstream>
// 保留诊断末尾，防止长批次日志无限占用内存；失败原因通常位于末尾。
void appendWorkerLog(std::string& log,const std::string& line) {
    log+=line;log+="\r\n";
    if(log.size()>MaximumLogBytes){
        const std::string notice="[Earlier log truncated; keeping the last 256 KiB]\r\n";
        size_t start=log.size()-(MaximumLogBytes-notice.size());const auto newline=log.find('\n',start);
        if(newline!=std::string::npos&&newline+1<log.size())start=newline+1;
        log.erase(0,start);log.insert(0,notice);
    }
}

std::wstring fromUtf8(const std::string& bytes) {
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
    if (reason == L"UNSUPPORTED_TEX_COMMENT") return L"公式含未转义的百分号注释，原式保留；百分号符号请写为 \\%";
    if (reason == L"REGION_SKIPPED") return L"所在区域处理失败，该公式未转换";
    if (reason == L"BODY_RANGE_SKIPPED") return L"该公式恢复未通过核验，已从原件重建副本并保留原式";
    if (reason == L"EXPERIMENTAL_CHECKPOINT_RESTORED") return L"实验区域未能完成核验，已恢复到完成正文转换时的副本";
    if (reason == L"Experimental region identities preserved") return L"实验区域保存前后的边界不一致";
    if (reason == L"Experimental recovery restores full original formula and layout") return L"无法确认实验公式及所在段落已经完整恢复";
    if (reason == L"Saved experimental text, layout, fields, OMML and ordered native contents verified") return L"实验区域保存后的内容或版式核验未通过";
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
    if (reason == L"Disable revision tracking in source before batch conversion") return L"文件开启修订，请先关闭修订并保存后重试。";
    if (reason == L"Protected document cannot be converted") return L"文件已限制编辑，请解除编辑限制后重试。";
    if (reason == L"Existing output directory required") return L"输出文件夹不存在或无法访问，请重新选择保存位置";
    if (reason == L"Create independent task copy" || reason == L"Reserve independent task copy" || reason == L"Fill reserved task copy") return L"无法在输出文件夹创建文档工作副本，请检查写入权限和磁盘剩余空间";
    if (reason == L"Publish verified file without overwrite") return L"无法交付输出文件，请检查目标文件是否被占用、文件夹写入权限及磁盘剩余空间";
    if (reason == L"Output already exists" || reason == L"No free output file name") return L"输出文件名已被占用，请修改文件名后缀或选择其他文件夹";
    if (reason == L"Write TeX into new object") return L"MathType 未能导入该公式的 LaTeX，请检查命令和环境是否受 MathType 支持";
    if (reason == L"Import OMML into new object") return L"MathType 未能导入该 Word 原生公式，原公式保留";
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
std::wstring failureText(const std::wstring& reason) {
    const auto chinese=reasonText(reason);
    return chinese==reason?chinese:chinese+L"\r\n原始错误："+reason;
}

std::wstring problemReason(const std::wstring& reason) {
    if(reason==L"OFFICE_OMML_TRANSFORM_MISSING")return L"Word 原生公式的转换组件缺失。";
    const auto translated=reasonText(reason);
    // 路径和接口原文留在技术日志，摘要只展示可读原因。
    if(translated.find(L":\\")!=std::wstring::npos||translated.find(L"\\\\")!=std::wstring::npos)return L"文件或安装组件无法访问，请查看原始日志中的具体位置。";
    if(translated==L"程序在当前阶段处理失败，请查看下方原始错误和后台日志")return L"此项处理失败，详细错误见原始日志。";
    return translated;
}
std::wstring problemAdvice(const std::wstring& reason) {
    if(reason==L"EMPTY")return L"补上公式内容，或删除空的公式定界符。";
    if(reason==L"UNBALANCED_BRACES")return L"检查并补齐成对的 { }，再重新转换。";
    if(reason==L"UNBALANCED_ENVIRONMENTS")return L"将每个 begin 与同名 end 配对，并检查嵌套顺序。";
    if(reason==L"UNMATCHED_OPEN"||reason==L"UNMATCHED_CLOSE"||reason==L"MIXED_DELIMITERS"||reason==L"UNSAFE_MISSING_DELIMITER"||reason==L"AMBIGUOUS_DOLLARS")return L"检查这一处公式的起止定界符，使用完整配对的 $...$、$$...$$、\\(...\\) 或 \\[...\\]。";
    if(reason==L"UNSUPPORTED_NON_ASCII_TEX"||reason==L"UNSUPPORTED_NON_ASCII_EXPORT")return L"将中文说明移到公式外，或改用当前支持的 LaTeX 命令。";
    if(reason==L"UNSUPPORTED_TEX_COMMENT")return L"删除公式内的百分号注释；需要百分号符号时写为 \\%。";
    if(reason==L"UNSUPPORTED_LINE_BREAK")return L"删除公式中的普通换行，或改为受支持的 array、cases 等多行结构。";
    if(reason==L"Write TeX into new object"||reason==L"Experimental equation import")return L"检查该式使用的命令；可先将原式单独粘贴到 MathType，确认能正常导入。";
    if(reason==L"OFFICE_OMML_TRANSFORM_MISSING")return L"修复桌面版 Word 安装中的公式转换组件，再重试；可暂时关闭 OMML 转换。";
    if(reason==L"Import OMML into new object")return L"在 Word 中检查这一处原生公式；可暂时关闭 OMML 转换以保留它。";
    if(reason==L"Disable revision tracking in source before batch conversion")return L"在 Word 中关闭修订，保存并关闭原文后重试。";
    if(reason==L"Protected document cannot be converted")return L"在 Word 中解除编辑限制，保存并关闭原文后重试。";
    if(reason==L"MATHTYPE_TASK_BUSY")return L"等待正在运行的转换任务结束后再试。";
    if(reason==L"TASK_TIMEOUT"||reason==L"MATHTYPE_SERVICE_DISCONNECTED"||reason==L"MathType initialization")return L"确认 Word 和 MathType 可正常启动，再重新处理原文。";
    if(reason.find(L"(Win32=32)")!=std::wstring::npos||reason.find(L"(Win32=33)")!=std::wstring::npos)return L"保存并关闭占用该文档的程序，再重试。";
    if(reason.find(L"(Win32=5)")!=std::wstring::npos)return L"将文档和输出位置放到有读写权限的文件夹，再重试。";
    return L"打开原文检查所列位置；若仍无法处理，请复制原始日志进一步排查。";
}
void WorkerState::addProblem(WorkerProblem problem,const std::wstring& detail) {
    const auto same=[&](const WorkerProblem& prior){return prior.id==problem.id&&prior.region==problem.region&&prior.location==problem.location&&prior.source==problem.source&&prior.reason==problem.reason&&prior.state==problem.state&&prior.stage==problem.stage;};
    if(std::any_of(problems.begin(),problems.end(),same))return;
    problems.push_back(std::move(problem));if(!detail.empty())failures.push_back(detail);
}

void WorkerState::reset(bool fastConversion) {
    *this={};fast=fastConversion;busy=true;lastStage=L"启动后台任务";
}

WorkerUpdate WorkerState::read(const std::string& value) {
    std::istringstream stream(value);std::string kind;stream>>kind;
    if(kind=="UI_CANCELLED"){outcome=L"CANCELLED";summary=false;error.clear();workerFailed=false;return {};}
    if(kind=="UI_COMMIT_READY")return {cancelling?WorkerEvent::None:WorkerEvent::CommitReady};
    if(cancelling&&kind!="UI_ERROR")return {};
    if(kind=="UI_LAYOUT_CHECKED"){
        int checked{};require(static_cast<bool>(stream>>checked)&&(checked==0||checked==1),"Layout check result");layoutChecked=checked!=0;return {};
    }
    if(kind=="UI_LAYOUT_WARNING"){
        size_t index{};std::string code,region,location,source,message;
        require(static_cast<bool>(stream>>index>>code>>region>>location>>source>>message)&&index>0,"Layout warning record");
        warnings.push_back({index,fromUtf8(code),unhex(region),unhex(location),unhex(source),unhex(message)});return {};
    }
    if(kind=="UI_BODY_RESET"){
        failures.clear();problems.clear();warnings.clear();layoutChecked=false;currentProblem={};interrupted={};currentFormula.clear();targets=inlines=completed=succeeded=failed=unattempted=rejected=0;scanned=false;experimentalStarted=experimentalConverting=false;experimentalDone=0;
        lastStage=L"已从原件重建副本，正在跳过失败公式并继续转换";return {WorkerEvent::Status,lastStage};
    }
    if(kind=="UI_EXPERIMENTAL_PHASE"){
        std::string phase;require(static_cast<bool>(stream>>phase),"Experimental phase record");
        if(!experimentalStarted){experimentalStarted=true;experimentalSucceeded=succeeded;experimentalFailed=failed;experimentalFailures=failures.size();experimentalProblems=problems.size();experimentalWarnings=warnings.size();experimentalLayoutChecked=layoutChecked;}
        if(phase=="PREPARING"){experimentalConverting=false;lastStage=L"正在准备实验区域";}
        else if(phase=="CONVERTING"){experimentalConverting=true;experimentalDone=0;lastStage=L"正在转换实验区域公式";}
        else if(phase=="VERIFYING"){experimentalConverting=false;lastStage=L"正在核验实验区域结果";}
        else{experimentalConverting=false;lastStage=L"正在处理实验区域";}
        return {WorkerEvent::BatchStatus,lastStage,0,0,true};
    }
    if(kind=="UI_EXPERIMENTAL_SKIPPED"){
        std::string reason;require(static_cast<bool>(stream>>reason),"Experimental skip reason");
        experimentalConverting=false;experimentalDone=0;currentFormula.clear();currentProblem={};lastStage=reason=="NO_TARGETS"?L"实验区域没有待转换公式":L"实验区域没有实际改动";
        return {WorkerEvent::BatchStatus,lastStage};
    }
    if(kind=="UI_EXPERIMENTAL_RESET"){
        require(experimentalStarted,"Experimental retry follows phase start");succeeded=experimentalSucceeded;failed=experimentalFailed;failures.resize(experimentalFailures);problems.resize(experimentalProblems);warnings.resize(experimentalWarnings);layoutChecked=experimentalLayoutChecked;currentProblem={};currentFormula.clear();experimentalConverting=false;experimentalDone=0;
        lastStage=L"已保留正文结果，正在跳过失败区域并继续转换";return {WorkerEvent::Status,lastStage};
    }
    if(kind=="UI_PROCESSING_ISSUE"){
        std::string region,stage,reason,location;require(static_cast<bool>(stream>>region>>stage>>reason>>location),"Processing issue record");const auto step=unhex(stage);
        std::string source;stream>>source;
        const auto detail=fast?L"快速模式下发生错误，当前输出未经核验。":step==L"NUMBERING"?L"本次编号及引用修改已撤回，公式转换结果保留。":step==L"FORMATTING"?L"本段排版未更新，原公式及段落格式保留。":L"该区域未转换，原内容保留。";
        addProblem({0,unhex(region),unhex(location),unhex(source),unhex(reason),fast?L"FAILED_UNVERIFIED":L"PROCESSING_ISSUE",step},unhex(location)+L"\r\n"+unhex(region)+L"\r\n"+failureText(unhex(reason))+L"\r\n"+detail);return {};
    }
    if(kind=="UI_EXPERIMENTAL_ITEM"||kind=="UI_EXPERIMENTAL_REJECTED"){
        size_t index{};std::string region,source,reason,location;require(static_cast<bool>(stream>>index>>region>>source),"Experimental diagnostic context");
        if(kind=="UI_EXPERIMENTAL_REJECTED")require(static_cast<bool>(stream>>reason),"Rejected experimental formula reason");
        require(static_cast<bool>(stream>>location),"Experimental source location");
        const auto context=L"实验性公式 "+std::to_wstring(index)+L"（"+unhex(region)+L"）\r\n原式片段："+unhex(source);
        WorkerProblem problem{index,unhex(region),unhex(location),unhex(source),unhex(reason),L"",L""};
        if(kind=="UI_EXPERIMENTAL_ITEM"){currentFormula=context;currentProblem=std::move(problem);}
        else{problem.state=L"REJECTED_PRESERVED";addProblem(std::move(problem),unhex(location)+L"\r\n"+context+L"\r\n"+failureText(unhex(reason))+L"\r\n该项未转换，原文保留。");}return {};
    }
    if(kind=="UI_ERROR"){
        std::string encoded;require(static_cast<bool>(stream>>encoded),"Worker error record");const auto message=unhex(encoded);
        if(cancelling&&message!=L"TASK_COPY_CLEANUP_FAILED")return {};
        if(!error.empty())error+=L"\r\n";if(!currentFormula.empty()){error+=L"中断时正在处理："+currentFormula+L"\r\n";interrupted=currentProblem;}
        addProblem({0,L"",L"",L"",message,L"TASK_FAILED",L""});
        error+=failureText(message);workerFailed=true;return {};
    }
    if(kind=="UI_ITEM_BEGIN"||kind=="UI_ITEM_REJECTED"){
        size_t index{};long paragraph{},start{},end{};std::string source,reason,location;
        require(static_cast<bool>(stream>>index>>paragraph>>start>>end>>source)&&index>0,"Formula diagnostic context");
        auto excerpt=unhex(source);for(auto& c:excerpt)if(c==L'\r'||c==L'\n'||c==L'\v'||c==L'\t')c=L' ';
        if(kind=="UI_ITEM_REJECTED")require(static_cast<bool>(stream>>reason),"Rejected formula reason");
        require(static_cast<bool>(stream>>location),"Formula source location");
        const auto context=L"公式 "+std::to_wstring(index)+L"（原文第 "+std::to_wstring(paragraph)+L" 段，Word 位置 "+std::to_wstring(start)+L"–"+std::to_wstring(end)+L"）\r\n原式片段："+excerpt;
        WorkerProblem problem{index,L"正文",unhex(location),excerpt,unhex(reason),L"",L""};
        if(kind=="UI_ITEM_BEGIN"){currentFormula=context;currentProblem=std::move(problem);}
        else{problem.state=L"REJECTED_PRESERVED";addProblem(std::move(problem),unhex(location)+L"\r\n"+context+L"\r\n"+failureText(unhex(reason))+L"\r\n该项未转换，原文保留。");}return {};
    }
    if(kind=="UI_ITEM_END"){currentFormula.clear();currentProblem={};return {};}
    if(kind=="UI_SCAN_COUNTS"){
        require(static_cast<bool>(stream>>existing>>targets>>inlines>>rejected)&&inlines<=targets,"Scan counters");scanned=true;return {WorkerEvent::ScanCounts};
    }
    if(kind=="UI_PHASE"){
        std::string phase;stream>>phase;currentFormula.clear();currentProblem={};
        const wchar_t* label=phase=="PREPARING"?L"正在检查输入文档…":phase=="OPENING"?L"正在加载文档…":phase=="SCANNING"?L"正在识别文档公式…":
            phase=="SCANNING_OBJECTS"?L"正在识别已有 MathType…":phase=="SOURCE_LOCATIONS"?L"正在记录原文公式位置…":phase=="INITIALIZING"?L"正在准备公式样式…":phase=="CONVERTING"?L"正在转换公式…":
            phase=="VERIFYING"?L"正在核对公式、残留原文和文档结构…":phase=="SAVING"?L"正在保存副本…":phase=="REOPENING"?L"正在重新打开并核验新公式…":
            phase=="RESCANNING"?L"正在检查保存后的正文…":phase=="RESCANNING_OBJECTS"?L"正在检查保存后的公式对象…":phase=="CLOSING"?L"正在关闭检查文档…":
            phase=="PACKAGE_VERIFY"?L"正在检查输出文件…":phase=="PACKAGE_COMPARE"?L"正在核对原文件及公式数据…":L"正在完成输出…";
        lastStage=fromUtf8(phase)+L" · "+label;
        return {WorkerEvent::Phase,label,0,targets,phase!="CONVERTING"||!targets,false,false,phase=="REOPENING"};
    }
    if(kind=="UI_STAGE_PROGRESS"){
        std::string stage;size_t current{},total{};require(static_cast<bool>(stream>>stage>>current>>total)&&total>0&&current<=total&&total<=INT_MAX,"Stage progress record");
        // 新进度阶段属于展示信息，不能把已经成功的转换判成失败。
        const wchar_t *label=L"正在处理文档",*unit=L"项";bool scanning=false,saved=false;
        if(stage=="INPUT_PARTS")label=L"正在检查输入文件";
        else if(stage=="SAVED_PARTS"){label=L"正在检查输出文件";saved=true;}
        else if(stage=="SCAN_TEXT"||stage=="RESCAN_TEXT"){saved=stage=="RESCAN_TEXT";scanning=!saved;label=saved?L"正在检查保存后的正文":L"正在识别公式";unit=L"段";}
        else if(stage=="SCAN_OBJECTS"||stage=="RESCAN_OBJECTS"){saved=stage=="RESCAN_OBJECTS";scanning=!saved;label=saved?L"正在检查保存后的公式对象":L"正在识别已有公式";unit=L"个对象";}
        else if(stage=="SCAN_RANGES"||stage=="RESCAN_RANGES"){saved=stage=="RESCAN_RANGES";scanning=!saved;label=saved?L"正在核对保留原文":L"正在核对识别范围";}
        else if(stage=="BIND_OBJECTS")label=L"正在核对公式位置";
        else if(stage=="VERIFY_POSITIONS"){label=L"正在核对保存后的公式位置";saved=true;}
        else if(stage=="VERIFY_ITEMS"){label=L"正在核验转换结果";saved=true;}
        else if(stage=="NUMBER_EQUATIONS")label=L"正在添加公式编号";
        const auto counts=L" "+std::to_wstring(current)+L" / "+std::to_wstring(total);lastStage=fromUtf8(stage)+L" · "+label+counts;
        return {WorkerEvent::Progress,std::wstring(label)+counts+L" "+unit,current,total,false,scanning,saved};
    }
    if(kind=="CONVERT_PROGRESS"){
        size_t total{};require(sscanf_s(value.c_str(),"CONVERT_PROGRESS %zu/%zu",&completed,&total)==2&&completed<=total&&total<=INT_MAX,"Conversion progress record");
        return {WorkerEvent::ConversionProgress,L"正在转换 "+std::to_wstring(completed)+L" / "+std::to_wstring(total)+L"  ·  成功 "+std::to_wstring(succeeded)+(fast?L"，异常 ":L"，保留 ")+std::to_wstring(failed),completed,total};
    }
    if(kind=="UI_ITEM_RESULT"){
        size_t index{};std::string state,reason;require(static_cast<bool>(stream>>index>>state>>reason)&&index>0,"Conversion item record");
        succeeded+=state=="SUCCEEDED";failed+=state=="FAILED_PRESERVED"||state=="FAILED_UNVERIFIED";
        if(state=="FAILED_PRESERVED"||state=="FAILED_UNVERIFIED"){
            auto problem=currentProblem;problem.id=index;problem.reason=unhex(reason);problem.state=fromUtf8(state);
            addProblem(std::move(problem),(currentFormula.empty()?L"公式 "+std::to_wstring(index):currentFormula)+L"\r\n"+failureText(unhex(reason))+(state=="FAILED_UNVERIFIED"?L"\r\n快速转换失败，当前内容未经核验。":L"\r\n该式转换失败，原文保留。"));
        }
        currentFormula.clear();currentProblem={};
        if(experimentalConverting){++experimentalDone;return {WorkerEvent::Status,L"正在转换实验区域公式（已完成 "+std::to_wstring(experimentalDone)+L"）"};}return {};
    }
    if(kind=="UI_CONVERT_SUMMARY"){
        std::string state;size_t total{};require(static_cast<bool>(stream>>state>>total>>succeeded>>failed>>unattempted>>rejected)&&total==succeeded+failed+unattempted,"Final conversion counters");
        require(state=="COMPLETE"||state=="PARTIAL","Known conversion outcome");targets=total;outcome=fromUtf8(state);summary=true;
    }
    return {};
}

bool WorkerState::finish(DWORD code,const std::wstring& output,const std::wstring& report) {
    busy=false;const bool cancelled=code==TaskCancelExit&&outcome==L"CANCELLED";
    workerFailed=!cancelled&&(workerFailed||code!=0||!summary||GetFileAttributesW(output.c_str())==INVALID_FILE_ATTRIBUTES||GetFileAttributesW(report.c_str())==INVALID_FILE_ATTRIBUTES);
    if(workerFailed&&error.empty()){
        const auto reason=code!=0?L"后台进程未正常完成，原文件保留。":!summary?L"未收到转换结果摘要，原文件保留。":GetFileAttributesW(output.c_str())==INVALID_FILE_ATTRIBUTES?L"未找到已核验的输出文档，原文件保留。":L"未找到转换核验报告，原文件保留。";
        error=(currentFormula.empty()?L"":L"中断时正在处理："+currentFormula+L"\r\n")+std::wstring(reason);interrupted=currentProblem;
        addProblem({0,L"",L"",L"",reason,L"TASK_FAILED",L""});
    }
    if(workerFailed&&std::none_of(problems.begin(),problems.end(),[](const WorkerProblem& problem){return problem.state==L"TASK_FAILED";}))addProblem({0,L"",L"",L"",L"文件处理未能完成，请查看原始日志。",L"TASK_FAILED",L""});
    // 保留逐式问题，任务收尾错误不能覆盖最早的失败原因。
    for(const auto& reason:failures){if(!error.empty())error+=L"\r\n\r\n";error+=reason;}
    if(rejected){if(!error.empty())error+=L"\r\n";error+=std::to_wstring(rejected)+L" 项不满足识别条件，原文保留。";}
    if(cancelled&&error.empty())error=L"用户已停止当前批次，当前文档未交付转换结果。";
    return cancelled;
}

WorkerClient::~WorkerClient() { join(); }
void WorkerClient::join() { if(thread.joinable())thread.join(); }
void WorkerClient::start(HWND target,const std::wstring& program,const std::vector<std::wstring>& arguments,bool trace) {
    join();
    thread=std::thread([target,program,arguments,trace]() {
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
                std::wstring command=taskQuote(program); for(const auto& argument:arguments) command+=L" "+taskQuote(argument);
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
