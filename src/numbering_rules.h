// 扫描、编号和核验共用厂商域识别规则；保持现有代码匹配方式，不依赖 Word。
#pragma once
#include <string_view>

struct NativeSequenceRule {std::wstring_view name,code;};
inline constexpr NativeSequenceRule NativeSequenceRules[]{
    {L"MTEqn",L"SEQ MTEqn "},{L"MTChap",L"SEQ MTChap "},{L"MTSec",L"SEQ MTSec "}
};
inline constexpr bool nativeNumberField(long kind,std::wstring_view code) {
    return kind==51&&code.find(L"MACROBUTTON MTPlaceRef")!=std::wstring_view::npos;
}
inline constexpr bool nativeSequenceField(long kind,std::wstring_view code,std::wstring_view name={}) {
    if(kind!=12)return false;
    for(const auto& rule:NativeSequenceRules)
        if((name.empty()||name==rule.name)&&code.find(rule.code)!=std::wstring_view::npos)return true;
    return false;
}
inline constexpr bool nativeResetSeed(long kind,std::wstring_view code,std::wstring_view name={}) {
    return nativeSequenceField(kind,code,name)&&code.find(L"\\r")!=std::wstring_view::npos;
}
inline constexpr bool nativeNumberWithoutSeparator(std::wstring_view code) {
    return code.find(L" MACROBUTTON MTPlaceRef ")==0||code.find(L" MACROBUTTON MTEditEquationSection")==0;
}
inline constexpr bool nativeCurrentCounter(std::wstring_view code) {
    return code.find(L"SEQ MTEqn \\c")!=std::wstring_view::npos;
}
inline constexpr bool nativeIncrementCounter(std::wstring_view code) {
    return nativeSequenceField(12,code,L"MTEqn")&&code.find(L"\\c")==std::wstring_view::npos;
}
