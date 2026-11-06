// 界面只消费任务状态；工作进程、管道和文本协议在此集中管理。
#pragma once
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <string>
#include <thread>
#include <vector>
#include "conversion_mode.h"

constexpr UINT WorkerLine=WM_APP+1,WorkerDone=WM_APP+2,WorkerLog=WM_APP+5;
constexpr size_t MaximumLogBytes=256*1024;
void appendWorkerLog(std::string& log,const std::string& line);
std::wstring fromUtf8(const std::string& bytes);
std::wstring failureText(const std::wstring& reason);
std::wstring problemReason(const std::wstring& reason);
std::wstring problemAdvice(const std::wstring& reason);

struct WorkerProblem {
    size_t id{};
    std::wstring region,location,source,reason,state,stage;
};
// 版式提醒只供人工检查，不参与转换问题或成功、失败计数。
struct WorkerWarning {
    size_t id{};
    std::wstring code,region,location,source,message;
};

enum class WorkerEvent { None, CommitReady, ScanCounts, Status, BatchStatus, Phase, Progress, ConversionProgress, ConversionCount };
// 正常单式步骤共用界面文案和计时；失败恢复继续单独展示。
inline bool formulaConversionStage(const std::string& stage) {
    return stage=="CONVERTING"||stage=="PROCESS_FORMULAS"||stage=="EXP_FORMULAS"||
        (stage.rfind("FORMULA_",0)==0&&stage!="FORMULA_RECOVERING");
}
struct WorkerUpdate {
    WorkerEvent event{WorkerEvent::None};
    std::wstring label;
    size_t current{},total{};
    bool marquee{},scanning{},saved{},reopening{};
};
struct WorkerState {
    bool busy{},workerFailed{},summary{},scanned{},cancelling{},acceptingResult{},layoutChecked{};
    ConversionMode mode{ConversionMode::Standard};
    long existing{};
    size_t targets{},inlines{},completed{},succeeded{},failed{},unattempted{},rejected{};
    std::string progressStage;
    std::wstring error,outcome,lastStage,currentFormula;
    std::vector<std::wstring> failures;
    std::vector<WorkerProblem> problems;
    std::vector<WorkerWarning> warnings;
    WorkerProblem currentProblem,interrupted;
    WorkerUpdate read(const std::string& value);
    void reset(ConversionMode conversionMode);
    bool finish(DWORD code,const std::wstring& output,const std::wstring& report);
private:
    bool processingFormulas{},experimentalStarted{},experimentalConverting{},experimentalLayoutChecked{};
    size_t formulaProgress{},formulaTotal{};
    static std::wstring formulaLabel(size_t current,size_t total) {
        return L"正在转换公式…  ·  已完成 "+std::to_wstring(current)+(total?L" / "+std::to_wstring(total):L"");
    }
    size_t experimentalSucceeded{},experimentalFailed{},experimentalFailures{},experimentalProblems{},experimentalWarnings{},experimentalDone{};
    void addProblem(WorkerProblem problem,const std::wstring& detail={});
};

struct WorkerClient {
    ~WorkerClient();
    void start(HWND target,const std::wstring& executable,const std::vector<std::wstring>& arguments,bool trace);
    void join();
private:
    std::thread thread;
};
