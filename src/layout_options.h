// 界面、命令行和编号预览共用同一组设置，默认不触发额外 Word 调用。
#pragma once
#include <cmath>
#include <string>

struct EquationSpacing {
    int mode{}; // 0 跟随；1/2/3 单倍/1.5/双倍；4/5/6 固定/最小/多倍。
    double value{20};
};
struct EquationNumbering {
    std::wstring format{L"(#E1)"};
    long start{1}, chapter{1}, section{1};
    bool custom{};
};
inline double settingNumber(const std::wstring& text,double minimum,double maximum) {
    wchar_t* end{};const double value=wcstod(text.c_str(),&end);
    require(!text.empty()&&end==text.c_str()+text.size()&&std::isfinite(value)&&value>=minimum&&value<=maximum,
            "Setting number outside supported range");return value;
}
inline EquationSpacing parseSpacing(const std::wstring& text) {
    EquationSpacing result;
    for(const auto* name:{L"follow",L"single",L"onehalf",L"double"}) {
        if(text==name)return result;++result.mode;
    }
    for(const auto* name:{L"exact:",L"atleast:",L"multiple:"}) {
        if(text.rfind(name,0)==0){result.value=settingNumber(text.substr(wcslen(name)),result.mode==6?0.5:1,result.mode==6?10:999);return result;}
        ++result.mode;
    }
    throw std::runtime_error("Unknown equation line spacing");
}
inline std::wstring spacingArgument(const EquationSpacing& spacing) {
    const wchar_t* names[]={L"follow",L"single",L"onehalf",L"double",L"exact:",L"atleast:",L"multiple:"};
    return std::wstring(names[spacing.mode])+(spacing.mode>=4?std::to_wstring(spacing.value):std::wstring());
}
inline std::wstring numberText(long value,wchar_t style) {
    require(value>=1&&value<=999999,"Number must be 1..999999");
    if(style==L'1')return std::to_wstring(value);
    // Word 原生 ALPHABETIC 在 Z 之后使用 AA、BB；预览与域结果保持一致。
    if(style==L'a'||style==L'A')return std::wstring(static_cast<size_t>((value-1)/26+1),static_cast<wchar_t>((style==L'a'?L'a':L'A')+(value-1)%26));
    require(style==L'i'||style==L'I',"Number style must be 1, a, A, i or I");
    const long values[]={1000,900,500,400,100,90,50,40,10,9,5,4,1};
    const wchar_t* symbols[]={L"M",L"CM",L"D",L"CD",L"C",L"XC",L"L",L"XL",L"X",L"IX",L"V",L"IV",L"I"};
    std::wstring result;for(size_t i=0;i<13;++i)while(value>=values[i]){result+=symbols[i];value-=values[i];}
    if(style==L'i')for(auto& letter:result)letter=static_cast<wchar_t>(towlower(letter));return result;
}
struct EquationNumberParts {std::wstring before,after,digits;wchar_t style{L'1'};};
inline EquationNumberParts numberParts(const EquationNumbering& settings,long ordinal) {
    require(!settings.format.empty()&&settings.format.size()<=128&&settings.format.find_first_of(L"\r\n\t") == std::wstring::npos,
            "Number format must be 1..128 characters on one line");
    EquationNumberParts result;bool equation{};
    for(size_t i=0;i<settings.format.size();++i){
        auto& output=equation?result.after:result.before;
        if(settings.format[i]==L'#'&&i+2<settings.format.size()&&
           (settings.format[i+1]==L'C'||settings.format[i+1]==L'S'||settings.format[i+1]==L'E')){
            const auto component=settings.format[i+1],style=settings.format[i+2];
            if(component==L'E'){require(!equation,"Use exactly one #E token");equation=true;result.style=style;result.digits=numberText(ordinal,style);}
            else output+=numberText(component==L'C'?settings.chapter:settings.section,style);
            i+=2;
        }else output+=settings.format[i];
    }
    require(equation,"Number format needs one #E1, #Ea, #EA, #Ei or #EI token");return result;
}
inline const wchar_t* numberSwitch(wchar_t style) {
    switch(style){case L'a':return L"alphabetic";case L'A':return L"ALPHABETIC";case L'i':return L"roman";case L'I':return L"ROMAN";default:return L"Arabic";}
}
