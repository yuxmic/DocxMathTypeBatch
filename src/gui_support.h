// 界面与验收共用的控件标识、队列记录和 Windows 辅助操作。
#pragma once
#include "worker_client.h"
#include "app_error.h"
#include <algorithm>
enum Control { Title=100, Forward, Reverse, FileHeading, FilePath, Choose, Existing,
    StageHeading, Detail, StyleHeading, ModeLabel, ModeBox, SizeLabel, SizeBox,
    SizeSpin, StyleNote, OutputHeading, DirectoryLabel, DirectoryBox, ChooseDirectory, NameLabel, NameBox,
    OutputNote, Status, Progress, Start, Stop, OpenOutput, OpenFolder, OpenReport, Environment,
    InlineDelimiterLabel, InlineDelimiterNote, DisplayDelimiterLabel, Bracket, Dollar, FollowSize, AddNumbers,
    QueueList, ClearQueue, AddFolder, OriginalDirectory, CustomDirectory, Advanced, Elapsed, NumberLabel, FastConversion, FastLabel, FastHelp };

enum class DocumentState { Pending, Running, Complete, Partial, Failed, Cancelled };
constexpr const wchar_t* DocumentStates[]={L"待处理",L"转换中",L"完成",L"部分完成",L"失败",L"已取消"};
struct QueuedDocument {
    std::wstring path, output, error;
    ULONGLONG bytes{};
    DocumentState state{DocumentState::Pending};
    std::wstring context, stage;
    std::string trace;
    DWORD exitCode{STILL_ACTIVE};
    bool fast{};
    std::vector<WorkerProblem> problems;
    WorkerProblem interrupted;
    std::vector<WorkerWarning> warnings;
    bool layoutChecked{};
};

inline std::wstring windowText(HWND window) {
    const int length = GetWindowTextLengthW(window);
    std::wstring result(static_cast<size_t>(length) + 1, L'\0');
    GetWindowTextW(window, result.data(), length + 1); result.resize(static_cast<size_t>(length)); return result;
}
inline void saveNew(const std::wstring& path, const std::string& bytes) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(file != INVALID_HANDLE_VALUE, "Create new GUI task file"); DWORD written{};
    const bool saved = bytes.size() <= MAXDWORD && WriteFile(file, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) && written == bytes.size();
    const bool flushed = FlushFileBuffers(file) != 0, closed = CloseHandle(file) != 0;
    if (!saved || !flushed || !closed) { DeleteFileW(path.c_str()); throw std::runtime_error("Save complete GUI task file"); }
}
inline std::wstring documentDetails(const QueuedDocument& item,const std::wstring& environment) {
        std::wstring text=L"文档："+item.path+L"\r\n结果："+DocumentStates[static_cast<int>(item.state)]+(item.fast?L"（快速，未核验）":L"")+L"\r\n退出码："+(item.exitCode==STILL_ACTIVE?std::wstring(L"尚未结束"):std::to_wstring(item.exitCode))+L"\r\n最后阶段："+item.stage+L"\r\n";
        if(!item.error.empty())text+=L"原因："+item.error+L"\r\n";
        else if(item.state==DocumentState::Pending){const auto& problem=environment;if(!problem.empty())text+=L"未开始原因："+problem+L"\r\n";}
        text+=L"\r\n"+item.context+L"\r\n本次后台日志\r\n";
        // 诊断输出允许替换无效 UTF-8 字节，不能因日志解码影响转换结果。
        const int length=MultiByteToWideChar(CP_UTF8,0,item.trace.data(),static_cast<int>(item.trace.size()),nullptr,0);
        if(length>0){std::wstring decoded(static_cast<size_t>(length),L'\0');MultiByteToWideChar(CP_UTF8,0,item.trace.data(),static_cast<int>(item.trace.size()),decoded.data(),length);text+=decoded;}
        else text+=L"尚无后台日志；任务完成或失败后可在这里查看。";
        return text;
}

inline std::wstring documentSummary(const QueuedDocument& item,const std::wstring& environment) {
    const auto name=item.path.substr(item.path.find_last_of(L"\\/")+1);
    std::wstring text=std::wstring(DocumentStates[static_cast<int>(item.state)])+(item.fast?L"（快速）":L"")+L" · "+std::to_wstring(item.problems.size())+L" 条问题";
    if(!item.warnings.empty())text+=L" · "+std::to_wstring(item.warnings.size())+L" 条版式提醒";
    text+=L"\r\n"+name+L"\r\n";
    if(item.state==DocumentState::Failed){
        text+=item.output.empty()?L"本文件未交付转换结果。\r\n":L"结果已保存，但任务收尾未完成。\r\n";
        if(!item.stage.empty()){const auto split=item.stage.find(L" · ");text+=L"处理阶段："+(split==std::wstring::npos?item.stage:item.stage.substr(split+3))+L"\r\n";}
    }else if(item.state==DocumentState::Cancelled)text+=L"用户已取消，当前文件未交付转换结果。\r\n";
    else if(item.state==DocumentState::Pending)text+=environment.empty()?L"尚未开始处理。\r\n":problemReason(environment)+L"\r\n";
    else if(item.state==DocumentState::Running)text+=L"正在处理，任务结束后可查看完整问题摘要。\r\n";
    else if(item.state==DocumentState::Complete&&item.problems.empty()&&item.warnings.empty())text+=item.fast?L"未报告转换错误，快速输出未经核验。\r\n":L"未发现转换问题。\r\n";
    if(item.fast&&!item.layoutChecked)text+=L"未进行版式检查。\r\n";
    const auto hasPage=[](const auto& record){return record.location.find(L"原文第")!=std::wstring::npos&&record.location.find(L"页")!=std::wstring::npos;};
    if(std::any_of(item.problems.begin(),item.problems.end(),hasPage)||hasPage(item.interrupted)||std::any_of(item.warnings.begin(),item.warnings.end(),hasPage)){
        text+=item.warnings.empty()?L"页码按转换前原文从首页连续计数（含封面）。\r\n":L"页码按转换前原文从首页连续计数（含封面）；版式提醒针对转换结果，原文页码不代表结果页码。\r\n";
    }else if(!item.warnings.empty())text+=L"版式提醒针对转换结果；所列位置对应转换前原文。\r\n";

    for(size_t i=0;i<item.problems.size();++i){
        const auto& problem=item.problems[i];
        const auto location=problem.state==L"TASK_FAILED"?L"文件处理问题":problem.location.empty()?(problem.region.empty()?L"原文页码未知":L"原文页码未知 · "+problem.region):problem.location;
        text+=L"\r\n"+std::to_wstring(i+1)+L". "+location+(problem.id?L" · 识别序号 "+std::to_wstring(problem.id):L"")+L"\r\n";
        if(!problem.source.empty()){
            auto source=problem.source;for(auto& c:source)if(c==L'\r'||c==L'\n'||c==L'\v'||c==L'\t')c=L' ';
            text+=L"原式片段："+source+L"\r\n";
        }
        const auto reason=problem.stage==L"RECOVERY"?L"该式转换失败，原式恢复未通过核验。":problemReason(problem.reason);
        const auto advice=problem.stage==L"RECOVERY"?L"打开原文检查上述公式；该式已保留，可先单独处理后再转换。":problemAdvice(problem.reason);
        text+=L"原因："+std::wstring(reason)+L"\r\n建议："+advice+L"\r\n";
        const wchar_t* retained=problem.state==L"TASK_FAILED"?L"本文件任务未正常完成。":problem.state==L"FAILED_UNVERIFIED"?L"快速输出未经核验，请检查该位置。":problem.state==L"REJECTED_PRESERVED"?L"未转换，原式保留。":problem.state==L"FAILED_PRESERVED"?L"转换失败，原式已保留。":problem.stage==L"NUMBERING"?L"编号及引用修改已撤回，公式转换结果保留。":problem.stage==L"FORMATTING"?L"本段排版保留原样。":problem.stage==L"RECOVERY"?L"已从原件重建副本并保留该式。":L"该区域未完成转换，原内容保留。";
        text+=L"处理结果："+std::wstring(retained)+L"\r\n";
    }
    if(item.interrupted.id||!item.interrupted.source.empty()){
        const auto& current=item.interrupted;text+=L"\r\n中断时正在处理："+(current.location.empty()?L"原文页码未知":current.location)+(current.id?L" · 识别序号 "+std::to_wstring(current.id):L"")+L"\r\n";
        if(!current.source.empty())text+=L"原式片段："+current.source+L"\r\n";
        text+=L"这仅表示中断位置，具体失败原因见上方文件处理问题。\r\n";
    }
    if(!item.warnings.empty())text+=L"\r\n版式提醒\r\n";
    for(size_t i=0;i<item.warnings.size();++i){
        const auto& warning=item.warnings[i];
        const auto location=warning.location.empty()?(warning.region.empty()?L"原文页码未知":L"原文页码未知 · "+warning.region):warning.location;
        text+=L"\r\n"+std::to_wstring(i+1)+L". "+location+(warning.id?L" · 识别序号 "+std::to_wstring(warning.id):L"")+L"\r\n";
        if(!warning.source.empty()){
            auto source=warning.source;for(auto& c:source)if(c==L'\r'||c==L'\n'||c==L'\v'||c==L'\t')c=L' ';
            text+=L"原式片段："+source+L"\r\n";
        }
        const auto advice=warning.code==L"FORMULA_TOO_WIDE"?L"在输出副本中检查此式；可适当缩小该式字号，或增加页面的可用宽度。":warning.code==L"EXACT_LINE_SPACING"?L"在输出副本中将所在段落的固定行距改为最小值或单倍行距，再检查公式是否完整显示。":warning.code==L"LAYOUT_CHECK_UNAVAILABLE"?L"打开输出副本检查该位置是否完整显示；显示正常无需修改。":L"打开输出副本检查该位置的公式显示，按提醒调整版式。";
        text+=L"提醒："+warning.message+L"\r\n建议："+advice+L"\r\n处理结果：公式处理结果已保留，请检查结果中的显示。\r\n";
    }
    return text;
}
