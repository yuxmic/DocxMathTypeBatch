// 正文阶段输入和结果；任务入口不持有公式坐标、编号跟踪或恢复状态。
#pragma once
#include "docx_scan.h"
#include "mathtype_api.h"
#include "docx_package.h"
#include "experimental_conversion.h"
#include "layout_options.h"
#include "layout_warnings.h"
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
inline bool failedOutcome(const ItemOutcome& item){return item.status==L"FAILED_PRESERVED"||item.status==L"FAILED_UNVERIFIED";}

inline bool displayOutcome(const ItemOutcome& outcome){return outcome.status==L"SUCCEEDED"||outcome.status==L"EXISTING_FORMATTED";}
inline bool nativeOutcome(const ItemOutcome& outcome){return displayOutcome(outcome)||outcome.status==L"EXISTING_PRESERVED";}
inline bool changedOutcome(const ItemOutcome& outcome){return displayOutcome(outcome)||outcome.status==L"REFERENCE_UPDATED";}

struct BodyOptions {
    ScanOptions scan;
    bool reverse{},addRightNumbers{},selfTest{},abortTest{},stopTest{};
    EquationSpacing inlineSpacing,displaySpacing;
    EquationNumbering numbering;
    double requestedPoints{};
    std::wstring stopFile;
};
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
