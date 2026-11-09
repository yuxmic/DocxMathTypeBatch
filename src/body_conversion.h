// 正文阶段输入和结果；任务入口不持有公式坐标、编号跟踪或恢复状态。
#pragma once
#include "docx_scan.h"
#include "mathtype_api.h"
#include "docx_package.h"
#include "experimental_conversion.h"
#include "layout_options.h"
#include "layout_warnings.h"
#include <map>
struct ConversionTimings;
struct ItemOutcome {
    Span finalSpan;
    std::wstring status, reason;
    std::vector<unsigned char> native;
    long addedParagraphs{};
    long shapeOrdinal{};
    std::wstring outputText;
    std::vector<std::vector<unsigned char>> removedPreviews;
    long nativeOrdinal{};
    Span replacementSpan;
};
// 当前目标由调用方维护，只平移它后面的结果范围。
inline void shiftFollowingOutcomes(std::vector<ItemOutcome>& outcomes,size_t current,long delta){
    if(!delta)return;
    for(size_t later=current+1;later<outcomes.size();++later){auto& outcome=outcomes[later];
        outcome.finalSpan.start+=delta;outcome.finalSpan.end+=delta;
        outcome.replacementSpan.start+=delta;outcome.replacementSpan.end+=delta;
    }
}
inline bool failedOutcome(const ItemOutcome& item){return item.status==L"FAILED_PRESERVED"||item.status==L"FAILED_UNVERIFIED";}

inline bool displayOutcome(const ItemOutcome& outcome){return outcome.status==L"SUCCEEDED"||outcome.status==L"EXISTING_FORMATTED";}
inline bool nativeOutcome(const ItemOutcome& outcome){return displayOutcome(outcome)||outcome.status==L"EXISTING_PRESERVED";}
inline bool changedOutcome(const ItemOutcome& outcome){return displayOutcome(outcome)||outcome.status==L"REFERENCE_UPDATED";}

struct BodyTarget {
    long end{};
    std::string delimiterHex,sourceHex;
};
struct BodyOptions {
    ScanOptions scan;
    bool reverse{},addRightNumbers{},selfTest{},abortTest{},stopTest{},formulaTabs{};
    EquationSpacing inlineSpacing,displaySpacing;
    EquationNumbering numbering;
    double requestedPoints{};
    std::wstring stopFile;
    std::map<long,BodyTarget> targets;
};
inline bool selectedBodyItem(const BodyOptions& options,const ScanItem& item){
    return !options.scan.exactTargets||options.targets.count(item.span.start)!=0;
}
struct NumberingSummary {long inserted{},updated{},referencesUpdated{},addedFields{};};
struct BodyResult {
    ScanResult scan;
    std::vector<ItemOutcome> outcomes;
    std::vector<ProcessingIssue> issues;
    std::vector<LayoutWarning> warnings;
    PackageState package;
    ExperimentalResult experimental;
    ExperimentalSourceLocations experimentalSourceLocations;
    std::wstring experimentalSkip;
    NumberingSummary numbering;
    size_t total{},completed{},succeeded{};
    bool savedNative{};
};
ScanResult scanReverseDocument(IDispatch* document,const ScanOptions& options={});
BodyResult convertBody(WordSession& word,MathTypeApi& api,const std::wstring& input,const std::wstring& candidate,
    const std::wstring& style,const PackageState& originalPackage,const BodyOptions& options,
    const ExperimentalOptions& experimentalOptions,const std::wstring& sizePreferencePath,bool& sizePreferenceOwned,bool& previewOwned,ConversionTimings& timings);
