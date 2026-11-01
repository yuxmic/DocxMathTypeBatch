// 版面提醒只描述测量结果，不改变公式内容，也不把疑似显示问题判成转换失败。
#pragma once
#include "equation_styles.h"

struct LayoutWarning {
    size_t id{};
    std::wstring region,code,source;
    SourcePosition position;
    double actual{},limit{};
    std::string json() const {
        return "{\"Id\":"+std::to_string(id)+",\"Region\":"+quote(region)+",\"Code\":"+quote(code)+
            ",\"Source\":"+quote(source)+sourcePositionJson(position)+",\"Actual\":"+std::to_string(actual)+
            ",\"Limit\":"+std::to_string(limit)+"}";
    }
};

// Word 尺寸有小量取整差异；差值超过 1 磅才提示，临界值不作肯定判断。
inline bool exceedsLayoutLimit(double actual,double limit) {return actual>limit+1.0;}
inline bool exactLineSpacingRisk(long rule,double height,double line) {
    return rule==4&&exceedsLayoutLimit(height,line);
}
inline double displayWidthLimit(double columnWidth,double leftIndent,double rightIndent) {
    return columnWidth-leftIndent-rightIndent;
}
inline void appendDimensionWarning(const wchar_t* code,double actual,double limit,const LayoutWarning& context,
                                   std::vector<LayoutWarning>& warnings) {
    if(!exceedsLayoutLimit(actual,limit))return;
    auto warning=context;warning.code=code;warning.actual=actual;warning.limit=limit;warnings.push_back(std::move(warning));
}
inline void appendFixedLineSpacingWarning(IDispatch* shapeRange,double height,const LayoutWarning& context,
                                          std::vector<LayoutWarning>& warnings) {
    auto format=get(shapeRange,L"ParagraphFormat");
    const long rule=integer(format.Get(),L"LineSpacingRule");
    if(rule!=4)return;
    const double line=number(format.Get(),L"LineSpacing");
    if(exactLineSpacingRisk(rule,height,line))appendDimensionWarning(L"EXACT_LINE_SPACING",height,line,context,warnings);
}
inline std::wstring layoutPoints(double value) {
    wchar_t text[64]{};swprintf_s(text,L"%.1f",value);return text;
}
inline std::wstring warningMessage(const LayoutWarning& warning) {
    if(warning.code==L"LAYOUT_CHECK_UNAVAILABLE")return L"未能读取公式版式，请人工检查显示范围和行距。";
    if(warning.code==L"FORMULA_TOO_WIDE")return L"公式宽度为 "+layoutPoints(warning.actual)+L" 磅，超过当前段落可用宽度 "+layoutPoints(warning.limit)+L" 磅，可能超出版心。";
    return L"公式高度为 "+layoutPoints(warning.actual)+L" 磅，超过固定行距 "+layoutPoints(warning.limit)+L" 磅，可能显示不全。";
}
