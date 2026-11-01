// TeX 文本规则与字符兼容改写；不读取文档或任务状态。
#pragma once
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <algorithm>
#include <cwctype>
#include <string>
#include <string_view>
#include <vector>
#include "app_error.h"

inline bool whitespaceOnly(const std::wstring& value) {
    return std::all_of(value.begin(), value.end(), [](wchar_t c) { return iswspace(c) || c == L'\a'; });
}
// Word 换行只是源码排版空白；TeX 的双反斜杠必须原样交给 MathType。
inline std::wstring normalizedTex(std::wstring value) {
    for (auto& c : value) if (c == L'\r' || c == L'\n' || c == L'\v') c = L' ';
    return value;
}
inline bool structuredMultiline(const std::wstring& value) {
    if (value.find(L"\\\\") != std::wstring::npos) return true;
    for (const auto* name : {L"align",L"aligned",L"alignat",L"alignedat",L"split",L"gather",L"gathered",L"subarray",L"array",L"cases",L"matrix",L"pmatrix",L"bmatrix",L"vmatrix",L"Vmatrix",L"smallmatrix"})
        if (value.find(std::wstring(L"\\begin{") + name + L"}") != std::wstring::npos ||
            value.find(std::wstring(L"\\begin{") + name + L"*}") != std::wstring::npos) return true;
    return false;
}
// 只检查基本语法与环境配对，支持度仍由 MathType 的导入结果判定。
inline std::wstring lexicalReason(const std::wstring& inner,bool fast=false) {
    if(fast)return {};
    if (whitespaceOnly(inner)) return L"EMPTY";
    if (inner.find_first_of(L"\f\a") != std::wstring::npos ||
        (inner.find_first_of(L"\r\n\v") != std::wstring::npos && !structuredMultiline(inner))) return L"UNSUPPORTED_LINE_BREAK";
    long depth{};
    std::vector<std::wstring> environments;
    for (size_t index = 0; index < inner.size(); ++index) {
        const auto character = inner[index];
        if (character == L'\\' && index + 1 < inner.size()) {
            const bool begin = inner.compare(index,7,L"\\begin{") == 0, end = inner.compare(index,5,L"\\end{") == 0;
            if (begin || end) {
                const auto first = index + (begin ? 7 : 5), last = inner.find(L'}',first);
                if (last == std::wstring::npos) return L"UNBALANCED_ENVIRONMENTS";
                const auto name = inner.substr(first,last-first);
                if (begin) environments.push_back(name);
                else { if (environments.empty() || environments.back()!=name) return L"UNBALANCED_ENVIRONMENTS"; environments.pop_back(); }
            }
            ++index; continue;
        }
        // MathType 会把注释直接排成公式内容，未转义的百分号必须保留原式。
        if (character == L'%') return L"UNSUPPORTED_TEX_COMMENT";
        if (character == L'{') ++depth;
        if (character == L'}' && --depth < 0) return L"UNBALANCED_BRACES";
    }
    if (depth != 0) return L"UNBALANCED_BRACES";
    return environments.empty() ? L"" : L"UNBALANCED_ENVIRONMENTS";
}

// 只识别括号内的数字及章-序号；单位、说明和数学表达式不作为手动编号。
inline bool manualEquationNumber(const std::wstring& value) {
    const auto first=value.find_first_not_of(L" \t\r\n\v\f\u00a0\u3000"),last=value.find_last_not_of(L" \t\r\n\v\f\u00a0\u3000");
    if(first==std::wstring::npos||last<=first+1)return false;
    if(!((value[first]==L'('&&value[last]==L')')||(value[first]==L'（'&&value[last]==L'）')))return false;
    bool digit=false;
    for(auto i=first+1;i<last;++i){
        const auto c=value[i];if(c>=L'0'&&c<=L'9'){digit=true;continue;}
        if((c==L'-'||c==L'.'||c==L'－')&&digit){digit=false;continue;}return false;
    }
    return digit;
}

struct MathAccent {wchar_t mark,symbol;const wchar_t* tex;};
inline const MathAccent* mathAccent(wchar_t mark) {
    // 组合码用于识别原文，symbol 采用 MathType 原生导入能识别的顶饰码。
    static const MathAccent accents[]={
        {0x0300,L'`',L"\\grave"},{0x0301,0x00b4,L"\\acute"},{0x0302,L'^',L"\\hat"},
        {0x0303,L'~',L"\\tilde"},{0x0304,L'-',L"\\bar"},{0x0306,0x2323,L"\\breve"},
        {0x0307,0x0307,L"\\dot"},{0x0308,0x0308,L"\\ddot"},{0x030a,0x2218,L"\\overset{\\circ}"},
        {0x030c,0x02c7,L"\\check"},{0x20db,0x20db,L"\\dddot"},{0x20dc,0x20dc,L"\\ddddot"},
        {0x20d6,0x2190,L"\\overleftarrow"},{0x20d7,0x2192,L"\\vec"},{0x20e1,0x2194,L"\\overleftrightarrow"}
    };
    for(const auto& accent:accents)if(accent.mark==mark)return &accent;
    return nullptr;
}
inline std::wstring decomposedMathCharacter(wchar_t character) {
    if(character<128||(character>=0xd800&&character<=0xdfff))return std::wstring(1,character);
    wchar_t buffer[8]{};const int count=FoldStringW(MAP_COMPOSITE,&character,1,buffer,8);
    require(count>0,"Decompose Unicode math character");
    // 只展开已支持的顶饰；其他文字、兼容字符及数学字体保持原样。
    if(count>1&&std::all_of(buffer+1,buffer+count,[](wchar_t c){return mathAccent(c)!=nullptr;}))return std::wstring(buffer,count);
    return std::wstring(1,character);
}
inline std::wstring compatibleTex(std::wstring value) {
    if (value.find(L'\\') == std::wstring::npos&&std::all_of(value.begin(),value.end(),[](wchar_t c){return c<128;})) return value;
    struct Alignment { size_t begin{}, length{}, columns{1}, maximum{1}; int environment{}, braces{}; bool centered{}; };
    struct Replacement { size_t start{}, length{}; std::wstring text; };
    std::vector<Alignment> aligned;
    std::vector<Replacement> replacements;
    int environment{}, braces{};
    // 按完整控制序列匹配，不能把 \bmfoo 或转义的 \\bm 当作 \bm。
    const std::pair<const wchar_t*,const wchar_t*> aliases[]={
        {L"bm",L"\\boldsymbol"},{L"textnormal",L"\\text"},{L"textup",L"\\text"},{L"textmd",L"\\text"},
        {L"mathup",L"\\mathrm"},{L"lvert",L"\\vert"},{L"rvert",L"\\vert"},
        {L"mleft",L"\\left"},{L"mright",L"\\right"},{L"dotsc",L"\\ldots"},{L"dotsm",L"\\cdots"},{L"dotso",L"\\ldots"},
        {L"coloneqq",L":="},{L"eqqcolon",L"=:"},{L"widetilde",L"\\tilde"},{L"mathring",L"\\overset{\\circ}"}
    };
    for (size_t i = 0; i < value.size(); ++i) {
        auto topLevel = [&] { return !aligned.empty() && environment == aligned.back().environment && braces == aligned.back().braces; };
        if (value[i] == L'\\' && i + 1 < value.size()) {
            if (value[i + 1] == L'\\') {
                if (topLevel()) { auto& a = aligned.back(); a.maximum = std::max(a.maximum, a.columns); a.columns = 1; }
                ++i; continue;
            }
            const bool begin = value.compare(i, 7, L"\\begin{") == 0, end = value.compare(i, 5, L"\\end{") == 0;
            if (begin || end) {
                const auto first = i + (begin ? 7 : 5), last = value.find(L'}', first);
                if(last==std::wstring::npos)break;
                const auto name = value.substr(first, last - first);
                const bool paired=name==L"aligned"||name==L"align*"||name==L"split"||name==L"alignedat"||name==L"alignat*";
                const bool centered=name==L"gathered"||name==L"gather*";
                if(name==L"subarray")replacements.push_back({i,last+1-i,begin?L"\\begin{array}":L"\\end{array}"});
                if (begin) {
                    ++environment;
                    if (paired||centered) {
                        size_t after = last + 1; while (after < value.size() && iswspace(value[after])) ++after;
                        const bool option = after + 2 < value.size() && value[after] == L'[' && value[after + 2] == L']' &&
                            (value[after + 1] == L't' || value[after + 1] == L'b' || value[after + 1] == L'c');
                        size_t finish=option?after+3:last+1;
                        // alignedat 的组数不是公式内容，array 按实际顶层列数生成列格式。
                        if(name==L"alignedat"||name==L"alignat*"){
                            after=finish;while(after<value.size()&&iswspace(value[after]))++after;
                            if(after>=value.size()||value[after]!=L'{') {i=last;continue;}
                            const auto close=value.find(L'}',after+1);
                            if(close==std::wstring::npos||close==after+1||
                               !std::all_of(value.begin()+after+1,value.begin()+close,[](wchar_t c){return c>=L'0'&&c<=L'9';})||
                               value.find_first_of(L"123456789",after+1)>=close){i=last;continue;}
                            finish=close+1;
                        }
                        aligned.push_back({i, finish-i, 1, 1, environment, braces,centered});
                        i=finish-1;continue;
                    }
                } else {
                    if ((paired||centered)&&!aligned.empty()&&environment==aligned.back().environment) {
                        const auto a = aligned.back(); aligned.pop_back();
                        std::wstring columns; for (size_t c = 0; c < std::max(a.centered?size_t{1}:size_t{2}, std::max(a.columns, a.maximum)); ++c) columns += a.centered?L'c':c % 2 ? L'l' : L'r';
                        replacements.push_back({a.begin, a.length, L"\\begin{array}{" + columns + L"}"});
                        replacements.push_back({i, last + 1 - i, L"\\end{array}"});
                    }
                    --environment;
                }
                i = last; continue;
            }
            size_t finish=i+1;while(finish<value.size()&&((value[finish]>=L'a'&&value[finish]<=L'z')||(value[finish]>=L'A'&&value[finish]<=L'Z')))++finish;
            if(finish==i+1){++i;continue;}
            const auto command=value.substr(i+1,finish-i-1);
            for(const auto& alias:aliases)if(command==alias.first){replacements.push_back({i,finish-i,alias.second});break;}
            // 文字参数保持原样，里面形似数学命令的文本不能再次改写。
            if(command==L"text"||command==L"mbox"||command==L"operatorname"||command==L"textnormal"||command==L"textup"||command==L"textmd"){
                size_t argument=finish;while(argument<value.size()&&iswspace(value[argument]))++argument;
                if(command==L"operatorname"&&argument<value.size()&&value[argument]==L'*'){++argument;while(argument<value.size()&&iswspace(value[argument]))++argument;}
                if(argument<value.size()&&value[argument]==L'{'){
                    int depth=1;for(size_t j=argument+1;j<value.size();++j){if(value[j]==L'\\'&&j+1<value.size()){++j;continue;}if(value[j]==L'{')++depth;else if(value[j]==L'}'&&--depth==0){finish=j+1;break;}}
                }
            }
            i=finish-1;continue;
        }
        if(iswalnum(value[i])) {
            auto atom=decomposedMathCharacter(value[i]);
            const auto base=atom.front();std::wstring marks=atom.substr(1);size_t finish=i+1;
            while(finish<value.size()&&mathAccent(value[finish]))marks+=value[finish++];
            if(!marks.empty()) {
                atom.assign(1,base);for(const auto mark:marks)atom=std::wstring(mathAccent(mark)->tex)+L"{"+atom+L"}";
                replacements.push_back({i,finish-i,std::move(atom)});i=finish-1;continue;
            }
        }
        if (value[i] == L'&' && topLevel()) ++aligned.back().columns;
        if (value[i] == L'{') ++braces;
        if (value[i] == L'}') --braces;
    }
    std::sort(replacements.begin(), replacements.end(), [](const Replacement& a, const Replacement& b) { return a.start > b.start; });
    for (const auto& replacement : replacements) value.replace(replacement.start, replacement.length, replacement.text);
    return value;
}

inline std::wstring exportedInner(const std::string& value, bool display,bool fast=false) {
    require(fast||std::all_of(value.begin(), value.end(), [](unsigned char c) {return c < 128;}), "UNSUPPORTED_NON_ASCII_EXPORT");
    std::wstring tex(value.begin(), value.end());
    const auto first = tex.find_first_not_of(L" \t\r\n"), last = tex.find_last_not_of(L" \t\r\n");
    require(first != std::wstring::npos, "EMPTY_NATIVE_EXPORT"); tex = tex.substr(first, last - first + 1);
    const std::wstring open = display ? L"\\[" : L"$", close = display ? L"\\]" : L"$";
    require(tex.size() >= open.size() + close.size() && tex.substr(0, open.size()) == open &&
            tex.substr(tex.size() - close.size()) == close, "Native export delimiter matches true equation type");
    auto inner = tex.substr(open.size(), tex.size() - open.size() - close.size());
    require(lexicalReason(inner,fast).empty(), "UNSUPPORTED_OR_INVALID_NATIVE_EXPORT");
    // 嵌套定界符会让保存后的扫描拆出多个公式，导出前拒绝以便只保留当前对象。
    for(size_t i=0;!fast&&i<inner.size();++i){
        if(inner[i]==L'\\'&&i+1<inner.size()){require(std::wstring_view(L"[]()").find(inner[++i])==std::wstring_view::npos,"UNSUPPORTED_OR_INVALID_NATIVE_EXPORT");continue;}
        require(inner[i]!=L'$',"UNSUPPORTED_OR_INVALID_NATIVE_EXPORT");
    }
    return normalizedTex(inner);
}
