// 实验项独立于两个基本方向，默认关闭，不保存到用户设置。
#pragma once
#include "mathtype_api.h"
#include "docx_scan.h"

struct PackageState;
struct ExperimentalOptions {
    bool tables{},textboxes{},headers{},notes{};
    bool enabled() const {return tables||textboxes||headers||notes;}
};
struct ExperimentalItem {
    size_t id{};
    std::wstring region,kind,source,status,reason;
    std::string json() const;
};
// 区域扫描和编号失败不一定有可计数的公式，单独记录，不能伪造逐式数量。
struct ProcessingIssue {
    std::wstring key,region,stage,reason;
    std::string json() const;
};
// 实验分支只会通过正文表格移动主故事坐标；偏移从保存重开后的实际末端取得。
struct BodyPositionShift {long boundary{},offset{};};
struct ExperimentalResult {
    ExperimentalOptions options;
    std::vector<ExperimentalItem> items;
    std::vector<ProcessingIssue> issues;
    std::vector<BodyPositionShift> bodyPositionShifts;
    size_t succeeded{},failed{},failedUnverified{},rejected{},unattempted{};
    bool enabled{};
    // NO_TARGETS/NO_CHANGES 表示实验阶段直接复用正文已核验结果，没有保存重开。
    std::wstring skipped;
    bool partial() const {return failed||rejected||unattempted||!issues.empty();}
    std::string json() const;
};
ExperimentalResult convertExperimental(WordSession& word,MathTypeApi& api,const std::wstring& path,
    const ExperimentalOptions& options,bool reverse,bool dollar,size_t firstId,const PackageState& bodyPackage,bool selfTest=false);
// 只读文档上的低成本存在性判断：启用区域内没有任何可转换候选时返回 false。
bool experimentalTargetsPresent(IDispatch* document,const std::wstring& path,const ExperimentalOptions& options,bool reverse);
// 正文与其他区域共用 OMML 适配；字号、布局和编号由所属流程决定。
std::wstring ommlMathML(WordSession& word,const std::wstring& sourcePackage,ComPtr<IXMLDOMDocument2>& stylesheet,bool display);
std::wstring ommlSource(IDispatch* range);
