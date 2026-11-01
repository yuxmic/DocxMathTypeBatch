// 版面提醒纯回归：只验证正式判据、冻结定位和中文说明，不创建任何 COM 实例。
#pragma once
#include "layout_warnings.h"

namespace layout_warning_tests {
inline void run() {
    require(!exceedsLayoutLimit(300,300)&&!exceedsLayoutLimit(301,300)&&exceedsLayoutLimit(301.1,300),
            "Layout dimensions allow one point rounding tolerance");
    // 480 已是 Word 给出的净栏宽；左右缩进之外不再次扣装订线或镜像页边距。
    require(displayWidthLimit(480,18,12)==450&&displayWidthLimit(480,0,0)==480&&
            displayWidthLimit(240,12,6)==222,
            "Single and equal column widths preserve paragraph indents without subtracting gutter twice");
    require(exactLineSpacingRisk(4,30,20)&&!exactLineSpacingRisk(4,21,20)&&
            !exactLineSpacingRisk(3,30,20)&&!exactLineSpacingRisk(0,30,20)&&!exactLineSpacingRisk(5,30,20),
            "Only exact line spacing reports height above the tolerance");
    const LayoutWarning context{7,L"正文",L"",L"\\[x+1\\]",{42,3,L"原文第 3 页 · 正文"}};
    std::vector<LayoutWarning> warnings;
    appendDimensionWarning(L"FORMULA_TOO_WIDE",301,300,context,warnings);
    require(warnings.empty(),"A borderline equation does not produce a layout warning");
    appendDimensionWarning(L"FORMULA_TOO_WIDE",340.26,300,context,warnings);
    require(warnings.size()==1&&warnings[0].id==7&&warnings[0].source==context.source&&
            warnings[0].position.sourceStart==42&&warnings[0].position.page==3&&
            warnings[0].code==L"FORMULA_TOO_WIDE"&&warnings[0].actual==340.26&&warnings[0].limit==300,
            "Layout warning retains the original formula identity and frozen source location");
    const auto json=warnings[0].json();
    require(json.find("\"Code\":\"FORMULA_TOO_WIDE\"")!=std::string::npos&&
            json.find("\"Page\":3")!=std::string::npos&&json.find("\"SourceLocation\":")!=std::string::npos&&
            json.find("\"Actual\":340.260000")!=std::string::npos&&json.find("\"Limit\":300.000000")!=std::string::npos,
            "Layout warning JSON preserves numeric measurements and original page");
    const auto message=warningMessage(warnings[0]);
    require(message.find(L"340.3")!=std::wstring::npos&&message.find(L"300.0")!=std::wstring::npos&&
            message.find(L"可能")!=std::wstring::npos,"Width warning gives measured points without claiming a rendered overflow");
    appendDimensionWarning(L"EXACT_LINE_SPACING",30.5,20,context,warnings);
    require(warnings.size()==2&&warningMessage(warnings.back()).find(L"可能显示不全")!=std::wstring::npos,
            "Exact line spacing warning remains a suspected display issue");
    auto unknown=warnings.back();unknown.position={};
    require(unknown.json().find("\"Page\":null")!=std::string::npos,
            "Unknown source page is not replaced with a post-conversion page");
    auto unavailable=context;unavailable.code=L"LAYOUT_CHECK_UNAVAILABLE";
    require(warningMessage(unavailable).find(L"未能读取公式版式")!=std::wstring::npos&&
            unavailable.json().find("\"Code\":\"LAYOUT_CHECK_UNAVAILABLE\"")!=std::string::npos,
            "Unavailable layout measurement requests manual review without declaring conversion failure");
    std::puts("LAYOUT_WARNINGS_PASSED");
}
}
