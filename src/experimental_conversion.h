// 实验项独立于两个基本方向，默认关闭，不保存到用户设置。
#pragma once
#include "docx_scan.h"
#include "layout_warnings.h"

struct MathTypeApi;
struct PackageState;
struct ConversionTimings;
struct ExperimentalOptions {
    bool tables{},textboxes{},headers{},notes{};
    bool enabled() const {return tables||textboxes||headers||notes;}
};
struct ExperimentalItem {
    size_t id{};
    std::wstring region,kind,source,status,reason;
    SourcePosition position;
    std::string json() const;
};
// 区域扫描和编号失败不一定有可计数的公式，单独记录，不能伪造逐式数量。
struct ProcessingIssue {
    std::wstring key,region,stage,reason;
    SourcePosition position;
    std::wstring source;
    std::string json(bool fast=false) const;
};
void emitProcessingIssue(const ProcessingIssue& issue);
// 按所属区域及原始次序冻结，不以可能重复的公式文本作为身份。
struct ExperimentalSourceItem {long relativeStart{};std::wstring kind;SourcePosition position;};
struct ExperimentalSourceRegion {
    std::wstring key;
    SourcePosition position;
    std::vector<ExperimentalSourceItem> items;
};
using ExperimentalSourceLocations=std::vector<ExperimentalSourceRegion>;
ExperimentalSourceLocations captureExperimentalSourceLocations(IDispatch* document,const std::wstring& path,
    const ExperimentalOptions& options,bool reverse,const ScanOptions& scanOptions,bool paginated);
// 实验分支只会通过正文表格移动主故事坐标；偏移从保存重开后的实际末端取得。
struct BodyPositionShift {long boundary{},offset{};};
struct ExperimentalResult {
    ExperimentalOptions options;
    std::vector<ExperimentalItem> items;
    std::vector<ProcessingIssue> issues;
    std::vector<LayoutWarning> warnings;
    std::vector<BodyPositionShift> bodyPositionShifts;
    size_t succeeded{},failed{},failedUnverified{},rejected{},unattempted{};
    bool enabled{};
    // NO_TARGETS/NO_CHANGES 表示实验阶段直接复用正文已核验结果，没有保存重开。
    std::wstring skipped;
    bool partial() const {return failed||rejected||unattempted||!issues.empty();}
    std::string json(bool fast=false) const;
};
ExperimentalResult convertExperimental(WordSession& word,MathTypeApi& api,const std::wstring& path,
    const ExperimentalOptions& options,bool reverse,bool dollar,size_t firstId,const PackageState& bodyPackage,bool selfTest,const ScanOptions& scanOptions,ConversionTimings& timings,const ExperimentalSourceLocations& sourceLocations);
// 只读文档上的低成本存在性判断：启用区域内没有任何可转换候选时返回 false。
bool experimentalTargetsPresent(IDispatch* document,const std::wstring& path,const ExperimentalOptions& options,bool reverse,const ScanOptions& scanOptions);
