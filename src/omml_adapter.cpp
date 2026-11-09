// Office OMML 到 MathType 可导入 MathML 的共同转换路径。
#include "omml_adapter.h"
#include "omml_tex.h"
#include "docx_package.h"
#include "latex_rules.h"
#include "conversion_timing.h"
#include <string_view>

namespace {
constexpr bool needsMathMLCompatibility(std::wstring_view decoded) {
    for(const auto character:decoded)if(character>0x7f)return true;
    return false;
}
// 判定对象是 DOM 解码后的文本；预组合和组合顶饰都必须进入兼容处理。
static_assert(!needsMathMLCompatibility(L"x+1=y"),"ASCII MathML needs no accent rewrite");
static_assert(needsMathMLCompatibility(L"\u0176"),"Precomposed accents need compatibility");
static_assert(needsMathMLCompatibility(L"Y\u0302"),"Combining accents need compatibility");

// 使用 Cambria Math 字宽和 Word 实际排版测得的字宽。
const wchar_t* ommlSpaceWidth(wchar_t character) {
    switch(character){
    case L'\u0085':return L"0.66em";
    case L'\u2000':case L'\u2002':return L"0.5em";
    case L'\u2004':return L"0.3335em";
    case L'\u2005':return L"0.25em";
    case L'\u2007':return L"0.5537em";
    case L'\u2008':return L"0.2051em";
    case L'\u200a':return L"0.0557em";
    case L'\u202f':return L"0.2002em";
    default:return nullptr;
    }
}
wchar_t equivalentMathSpace(wchar_t character) {
    switch(character){
    case L'\u2001':case L'\u2028':case L'\u2029':case L'\u3000':return L'\u2003';
    case L'\u2006':return L'\u2009';
    default:return character;
    }
}
bool layoutControl(wchar_t character) {
    // 单个公式对象内没有自动断行，这些控制符在未断行时没有字宽。
    return character==0x00ad||character==0x200b||character==0x2060||character==0xfeff;
}
bool joiningControl(wchar_t character) {
    return character==0x180e||character==0x200c||character==0x200d;
}
void prepareOmmlWhitespace(IXMLDOMNode* equation) {
    // XSL 会把普通空格和不换行空格合为一种。先按 Word 的原式布局区分两者。
    ComPtr<IXMLDOMNode> tab;
    check(equation->selectSingleNode(Value(L".//w:tab").v.bstrVal,&tab),"Find Word equation tab");
    require(!tab,"OMML_UNSUPPORTED_TAB_STOP");
    ComPtr<IXMLDOMNodeList> texts;check(equation->selectNodes(Value(L".//m:t[contains(.,' ') or contains(.,'\t') or contains(.,'\n') or contains(.,'\r')]").v.bstrVal,&texts),"Select OMML text with ordinary whitespace");
    long count{};check(texts->get_length(&count),"OMML text count");
    for(long i=0;i<count;++i){
        ComPtr<IXMLDOMNode> node,normal;check(texts->get_item(i,&node),"OMML text node");
        Value value;value.v.vt=VT_BSTR;check(node->get_text(&value.v.bstrVal),"OMML text value");
        const std::wstring original=value.v.bstrVal;
        require(original.find(L'\t')==std::wstring::npos,"OMML_UNSUPPORTED_TAB_STOP");
        check(node->selectSingleNode(Value(L"../m:rPr[last()]/m:nor[not(contains('|0|false|off|',concat('|',translate(@m:val,'FALSEO','falseo'),'|')))]").v.bstrVal,&normal),"Check normal OMML text");
        // Word 把公式中的换行字符导出为一个独立空格 run。只查看同一表达式中紧挨的 run。
        const auto adjoiningAlnum=[&](bool preceding){
            ComPtr<IXMLDOMNode> adjacent;
            check(node->selectSingleNode(Value(preceding?L"../preceding-sibling::*[1][self::m:r]/m:t[last()]":
                                                       L"../following-sibling::*[1][self::m:r]/m:t[1]").v.bstrVal,&adjacent),"Read neighboring OMML run");
            if(!adjacent){
                check(node->selectSingleNode(Value(preceding?L"../preceding-sibling::*[1][self::m:sSub or self::m:sSup or self::m:sSubSup]":
                                                           L"../following-sibling::*[1][self::m:sSub or self::m:sSup or self::m:sSubSup]").v.bstrVal,&adjacent),"Read adjacent OMML script");
                return static_cast<bool>(adjacent);
            }
            Value textValue;textValue.v.vt=VT_BSTR;check(adjacent->get_text(&textValue.v.bstrVal),"Read neighboring OMML text");
            const std::wstring_view text(textValue.v.bstrVal?textValue.v.bstrVal:L"");
            return !text.empty()&&IsCharAlphaNumericW(preceding?text.back():text.front());
        };
        std::wstring updated;
        const std::wstring_view white=L" \r\n";
        for(size_t start=0;start<original.size();){
            if(white.find(original[start])==std::wstring::npos){updated+=original[start++];continue;}
            const auto end=std::min(original.find_first_not_of(white,start),original.size());
            const auto plus=end<original.size()&&original[end]==L'+'?end:
                start&&original[start-1]==L'+'?start-1:std::wstring::npos;
            bool binaryPlus=false;
            if(plus!=std::wstring::npos&&plus>0){
                const auto left=original.find_last_not_of(white,plus-1),right=original.find_first_not_of(white,plus+1);
                binaryPlus=left!=std::wstring::npos&&right!=std::wstring::npos&&
                    IsCharAlphaNumericW(original[left])&&IsCharAlphaNumericW(original[right]);
            }
            // 普通文字的词间空格与数学相邻字母数字之间的空格有可见宽度。
            if(normal||binaryPlus||((start?IsCharAlphaNumericW(original[start-1])!=0:adjoiningAlnum(true))&&
                        (end<original.size()?IsCharAlphaNumericW(original[end])!=0:adjoiningAlnum(false))))
                updated.append(end-start,L'\u00a0');
            start=end;
        }
        if(updated!=original)check(node->put_text(Value(updated.c_str()).v.bstrVal),"Prepare ordinary OMML whitespace");
    }
}

std::wstring ommlNodeSource(IXMLDOMNode* equation) {
    Value xml;xml.v.vt=VT_BSTR;check(equation->get_xml(&xml.v.bstrVal),"OMML snapshot XML");return withoutWordRevisionIds(xml.v.bstrVal);
}

void compatibleMathML(IXMLDOMDocument2* dom,IXMLDOMNode* sourceEquation) {
    auto content=[](IXMLDOMNode* node){Value value;value.v.vt=VT_BSTR;check(node->get_text(&value.v.bstrVal),"MathML token text");return std::wstring(value.v.bstrVal?value.v.bstrVal:L"");};
    auto nodes=[&](const wchar_t* query){
        ComPtr<IXMLDOMNodeList> list;check(dom->selectNodes(Value(query).v.bstrVal,&list),"Select MathML nodes");
        long count{};check(list->get_length(&count),"MathML node count");std::vector<ComPtr<IXMLDOMNode>> result;
        for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> node;check(list->get_item(i,&node),"MathML node");result.push_back(node);}return result;
    };
    auto name=[](IXMLDOMNode* node){Value value;value.v.vt=VT_BSTR;check(node->get_baseName(&value.v.bstrVal),"MathML token name");return std::wstring(value.v.bstrVal?value.v.bstrVal:L"");};
    auto element=[&](const wchar_t* tag){ComPtr<IXMLDOMNode> node;check(dom->createNode(Value(1L).v,Value(tag).v.bstrVal,Value(L"http://www.w3.org/1998/Math/MathML").v.bstrVal,&node),"Create MathML accent element");return node;};
    auto append=[](IXMLDOMNode* parent,IXMLDOMNode* child){ComPtr<IXMLDOMNode> added;check(parent->appendChild(child,&added),"Append MathML accent element");};
    // 范数的 U+2016 括号会被厂商导入丢弃；属性不进入正文文本，须在 ASCII 快路径前归一化。
    for(const auto& delimiter:nodes(L"//*[namespace-uri()='http://www.w3.org/1998/Math/MathML' and local-name()='mfenced']/@*[(name()='open' or name()='close') and .='\u2016']"))
        check(delimiter->put_text(Value(L"||").v.bstrVal),"Normalize MathML norm delimiter");
    // 运算及关系分隔符必须显式写成 mo，否则 MathType 会漏掉减号、不等号等符号。
    const auto fences=nodes(L"//*[namespace-uri()='http://www.w3.org/1998/Math/MathML' and local-name()='mfenced' and count(*)>1][contains(@separators,'+') or contains(@separators,'-') or contains(@separators,'\u2212') or contains(@separators,'/') or contains(@separators,'<') or contains(@separators,'>') or contains(@separators,'\u2264') or contains(@separators,'\u2265')]");
    for(auto fence=fences.rbegin();fence!=fences.rend();++fence) {
        ComPtr<IXMLDOMElement> attributes;check(fence->As(&attributes),"MathML fence attributes");
        Value value;check(attributes->getAttribute(Value(L"separators").v.bstrVal,&value.v),"MathML fence separators");
        std::wstring decoded=value.v.bstrVal;
        decoded.erase(std::remove_if(decoded.begin(),decoded.end(),[](wchar_t c){return std::wstring_view(L" \t\r\n").find(c)!=std::wstring_view::npos;}),decoded.end());
        std::vector<std::wstring> separators;
        for(size_t i=0;i<decoded.size();++i) {
            std::wstring symbol(1,decoded[i]);
            if(decoded[i]>=0xd800&&decoded[i]<=0xdbff&&i+1<decoded.size()&&decoded[i+1]>=0xdc00&&decoded[i+1]<=0xdfff)symbol+=decoded[++i];
            separators.push_back(symbol==L"\u2212"?L"-":symbol);
        }
        ComPtr<IXMLDOMNodeList> children;check((*fence)->selectNodes(Value(L"*").v.bstrVal,&children),"Select fenced arguments");
        long count{};check(children->get_length(&count),"Fenced argument count");std::vector<ComPtr<IXMLDOMNode>> arguments;
        for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> child;check(children->get_item(i,&child),"Fenced argument");arguments.push_back(child);}
        auto row=element(L"mrow");
        for(size_t i=0;i<arguments.size();++i) {
            if(i){auto symbol=element(L"mo");const auto& separator=separators[std::min(i-1,separators.size()-1)];
                check(symbol->put_text(Value(separator.c_str()).v.bstrVal),"Set explicit MathML separator");append(row.Get(),symbol.Get());}
            append(row.Get(),arguments[i].Get());
        }
        append(fence->Get(),row.Get());check(attributes->setAttribute(Value(L"separators").v.bstrVal,Value(L"").v),"Clear implicit MathML separators");
    }
    const auto formulaText=content(dom);
    // MathType 会丢掉部分不可见运算符或写入问号，保留源公式才能保住运算含义。
    // Office 会为 m:func 生成函数应用标记。厂商可正常导入该结构，原文显式写入的算符仍保留。
    require(formulaText.find_first_of(L"\u2062\u2063\u2064")==std::wstring::npos&&
            (formulaText.find(L'\u2061')==std::wstring::npos||
             (sourceEquation&&content(sourceEquation).find(L'\u2061')==std::wstring::npos)),
            "OMML_UNSUPPORTED_INVISIBLE_OPERATOR");
    // Word 的回退字体会显示欧甘空格的一条短横线，普通间距会抹掉它。
    require(formulaText.find(L'\u1680')==std::wstring::npos,"OMML_UNSUPPORTED_OGHAM_SPACE");
    auto spacing=[](wchar_t c){return ommlSpaceWidth(c)||equivalentMathSpace(c)!=c||layoutControl(c)||joiningControl(c);};
    if(std::any_of(formulaText.begin(),formulaText.end(),joiningControl)){
        const auto simpleText=std::all_of(formulaText.begin(),formulaText.end(),[&](wchar_t c){
            return c<128||spacing(c)||std::wstring_view(L"\u00a0\u2003\u2009\u205f").find(c)!=std::wstring_view::npos;
        });
        require(simpleText,"OMML_UNSUPPORTED_JOINING_CONTROL");
    }
    if(formulaText.find_first_of(L"\u00a0\u2003\u2009\u205f")!=std::wstring::npos||
       std::any_of(formulaText.begin(),formulaText.end(),spacing)){
        const auto textTokens=nodes(L"//*[namespace-uri()='http://www.w3.org/1998/Math/MathML' and (local-name()='mi' or local-name()='mn' or local-name()='mo' or local-name()='mtext')]");
        const auto gap=[&](const std::wstring& width){
            auto node=element(L"mspace");ComPtr<IXMLDOMElement> attributes;check(node.As(&attributes),"MathML space attributes");
            check(attributes->setAttribute(Value(L"width").v.bstrVal,Value(width.c_str()).v),"Set MathML space width");return node;
        };
        const auto widthEm=[](wchar_t c){
            if(const auto* width=ommlSpaceWidth(c))return wcstod(width,nullptr);
            switch(equivalentMathSpace(c)){
            case L'\u00a0':return 0.22;
            case L'\u2003':return 1.0;
            case L'\u2009':return 0.1665;
            case L'\u205f':return 0.2222;
            default:return 0.0;
            }
        };
        const auto sibling=[](IXMLDOMNode* node,bool before){
            ComPtr<IXMLDOMNode> other;
            check(node->selectSingleNode(Value(before?L"preceding-sibling::*[1]":L"following-sibling::*[1]").v.bstrVal,&other),"Read adjacent MathML token");
            return other;
        };
        const auto atom=[&](IXMLDOMNode* node){
            if(!node)return false;
            const auto tag=name(node),value=content(node);
            return (tag==L"mi"||tag==L"mn"||tag==L"msub"||tag==L"msup"||tag==L"msubsup")&&
                !value.empty()&&(value.size()!=1||widthEm(value.front())==0);
        };
        // 厂商用显式间距替代逗号和二元加号的自动间距，补回在本机默认排版中测得的宽度。
        const auto operatorGap=[&](IXMLDOMNode* node){
            const auto before=sibling(node,true),after=sibling(node,false);
            if(before&&name(before.Get())==L"mo"&&content(before.Get())==L","&&atom(after.Get()))return 0.18;
            for(const bool previous:{true,false}){
                const auto& sign=previous?before:after;
                const auto& operand=previous?after:before;
                if(!sign||name(sign.Get())!=L"mo"||content(sign.Get())!=L"+"||!atom(operand.Get()))continue;
                auto other=sibling(sign.Get(),previous);
                if(other){
                    const auto value=content(other.Get());
                    if(name(other.Get())==L"mspace"||(value.size()==1&&widthEm(value.front())>0))
                        other=sibling(other.Get(),previous);
                }
                if(atom(other.Get()))return 0.26;
            }
            return 0.0;
        };
        for(const auto& token:textTokens){
            const auto original=content(token.Get());
            const bool normal=name(token.Get())==L"mtext";
            if(original.size()==1&&!normal){
                const auto width=widthEm(original.front());
                if(width>0){
                    const auto extra=operatorGap(token.Get());
                    if(extra>0||original.front()==L'\u00a0'){
                        auto space=gap(extra>0?std::to_wstring(width+extra)+L"em":L"0.18em");ComPtr<IXMLDOMNode> parent;
                        check(token->get_parentNode(&parent),"Explicit MathML space parent");
                        check(parent->replaceChild(space.Get(),token.Get(),nullptr),"Restore explicit math space width");
                        continue;
                    }
                }
            }
            if(std::none_of(original.begin(),original.end(),[&](wchar_t c){return spacing(c)||(normal&&widthEm(c)>0);}))continue;
            auto row=element(L"mrow");std::wstring plain;
            auto flush=[&]{
                if(plain.empty())return;
                ComPtr<IXMLDOMNode> copy;check(token->cloneNode(VARIANT_FALSE,&copy),"Clone MathML text token");
                check(copy->put_text(Value(plain.c_str()).v.bstrVal),"Set MathML text token");
                append(row.Get(),copy.Get());plain.clear();
            };
            for(const auto c:original){
                if(const auto* width=ommlSpaceWidth(c)){
                    flush();append(row.Get(),gap(width).Get());
                }else if(normal&&widthEm(c)>0){
                    // MathType 的普通文字预览可能把不换行空格画成问号。
                    flush();append(row.Get(),gap(c==L'\u00a0'?L"0.17em":std::to_wstring(widthEm(c))+L"em").Get());
                }else if(!layoutControl(c)&&!joiningControl(c))plain+=equivalentMathSpace(c);
            }
            flush();ComPtr<IXMLDOMNode> first,last,parent,replaced;
            check(row->get_firstChild(&first),"First rewritten MathML token");check(row->get_lastChild(&last),"Last rewritten MathML token");
            check(token->get_parentNode(&parent),"MathML space parent");
            check(parent->replaceChild(first&&first.Get()==last.Get()?first.Get():row.Get(),token.Get(),&replaced),"Replace MathML spacing token");
        }
    }
    // 原 XML 可能用字符实体表示顶饰，只能在 DOM 解码后判断 ASCII 快路径。
    if(!needsMathMLCompatibility(formulaText))return;
    auto setGlyph=[](IXMLDOMNode* node,const std::wstring& glyph){
        check(node->put_text(Value(glyph.c_str()).v.bstrVal),"Set MathML accent glyph");
    };
    auto wrap=[&](ComPtr<IXMLDOMNode> base,const std::wstring& glyph){
        auto mover=element(L"mover"),symbol=element(L"mo");ComPtr<IXMLDOMElement> attributes;check(mover.As(&attributes),"MathML accent attributes");
        check(attributes->setAttribute(Value(L"accent").v.bstrVal,Value(L"true").v),"Set MathML accent");
        setGlyph(symbol.Get(),glyph);
        append(mover.Get(),base.Get());append(mover.Get(),symbol.Get());return mover;
    };
    const auto tokens=nodes(L"//*[namespace-uri()='http://www.w3.org/1998/Math/MathML' and (local-name()='mi' or local-name()='mn' or local-name()='mo' or local-name()='mtext')][not(parent::*[local-name()='mover' or local-name()='munder' or local-name()='munderover'] and count(preceding-sibling::*)>=1)]");
    // Office 会把 Y 与后面的组合符拆成相邻 token；先合回，随后仅装饰最后一个字符。
    for(const auto& node:tokens) {
        const auto marks=content(node.Get());if(marks.empty()||!std::all_of(marks.begin(),marks.end(),[](wchar_t c){return mathAccent(c)!=nullptr;}))continue;
        if(name(node.Get())==L"mtext")continue;
        ComPtr<IXMLDOMNode> previous,parent;check(node->get_parentNode(&parent),"MathML token parent");if(!parent)continue;
        check(node->get_previousSibling(&previous),"Previous MathML token");
        while(previous&&name(previous.Get()).empty()&&whitespaceOnly(content(previous.Get()))){ComPtr<IXMLDOMNode> before;check(previous->get_previousSibling(&before),"Previous MathML whitespace");previous=before;}
        if(!previous)continue;const auto tag=name(previous.Get());if(tag!=L"mi"&&tag!=L"mn"&&tag!=L"mo")continue;
        const auto combined=content(previous.Get())+marks;check(previous->put_text(Value(combined.c_str()).v.bstrVal),"Join MathML combining accents");
        ComPtr<IXMLDOMNode> removed;check(parent->removeChild(node.Get(),&removed),"Remove joined MathML accent token");
    }
    for(const auto& node:tokens) {
        ComPtr<IXMLDOMNode> parent;check(node->get_parentNode(&parent),"MathML letter parent");if(!parent)continue;
        const auto original=content(node.Get());std::wstring decomposed;
        // 普通文字仅处理原有组合符，保留预组合字母和文字样式。
        if(name(node.Get())==L"mtext")decomposed=original;
        else for(const auto character:original)decomposed+=decomposedMathCharacter(character);
        if(decomposed.size()<2||std::none_of(decomposed.begin()+1,decomposed.end(),[](wchar_t c){return mathAccent(c)!=nullptr;}))continue;
        auto row=element(L"mrow");std::wstring plain;
        auto clone=[&](const std::wstring& text){ComPtr<IXMLDOMNode> token;check(node->cloneNode(VARIANT_FALSE,&token),"Clone styled MathML token");check(token->put_text(Value(text.c_str()).v.bstrVal),"Set MathML base text");return token;};
        auto flush=[&]{if(!plain.empty()){auto token=clone(plain);append(row.Get(),token.Get());plain.clear();}};
        for(size_t i=0;i<decomposed.size();) {
            size_t finish=i+1;if(decomposed[i]>=0xd800&&decomposed[i]<=0xdbff&&finish<decomposed.size()&&decomposed[finish]>=0xdc00&&decomposed[finish]<=0xdfff)++finish;
            const auto base=decomposed.substr(i,finish-i);const auto firstMark=finish;
            while(finish<decomposed.size()&&mathAccent(decomposed[finish]))++finish;
            if(finish==firstMark)plain+=base;
            else {flush();auto token=clone(base);for(size_t j=firstMark;j<finish;++j)token=wrap(token,std::wstring(1,decomposed[j]));append(row.Get(),token.Get());}
            i=finish;
        }
        flush();ComPtr<IXMLDOMNode> replaced;check(parent->replaceChild(row.Get(),node.Get(),&replaced),"Replace MathML accented token");
    }
    auto singleAccentAtom=[&](ComPtr<IXMLDOMNode> atom){
        while(atom){
            const auto tag=name(atom.Get());
            if(tag!=L"mrow")return (tag==L"mi"||tag==L"mn"||tag==L"mo"||tag==L"mtext")&&content(atom.Get()).size()==1;
            ComPtr<IXMLDOMNode> child;
            check(atom->selectSingleNode(Value(L"self::*[count(*)=1]/*").v.bstrVal,&child),"Single MathML accent base");
            atom=std::move(child);
        }
        return false;
    };
    const auto groupBrackets=std::wstring_view(L"\u23de\u23df\u23b4\u23b5\ufe37\ufe38");
    if(formulaText.find_first_of(groupBrackets)!=std::wstring::npos){
        for(const auto& node:nodes(L"//*[local-name()='mover' or local-name()='munder']/*[local-name()='mo' or local-name()='mi']")){
            const auto glyph=content(node.Get());
            if(glyph.size()!=1||groupBrackets.find(glyph.front())==std::wstring_view::npos)continue;
            ComPtr<IXMLDOMNode> base;
            if(glyph.front()==0x23b4)
                check(node->selectSingleNode(Value(L"preceding-sibling::*[1][parent::*[local-name()='mover' and count(*)=2]]").v.bstrVal,&base),"Single overbracket base");
            // 固定宽度方括号只适用于单字符；整组需要原生的伸展模板。
            require(singleAccentAtom(base),"OMML_UNSUPPORTED_STRETCHY_BRACKET");
        }
    }
    // 先构造组合符的 mover，再按基底结构选用顶饰，避免嵌套点饰落入厂商的未知组合码路径。
    for(const auto& node:nodes(L"//*[local-name()='mover']/*[2][local-name()='mo' or local-name()='mi']")){
        const auto text=content(node.Get());if(text.size()!=1)continue;auto mark=text.front();
        if(mark==0x02c6)mark=0x0302;else if(mark==0x02dc)mark=0x0303;else if(mark==0x02d8)mark=0x0306;
        else if(mark==0x00af||mark==0x203e||mark==0x0305)mark=0x0304;
        else if(mark==0x02d9)mark=0x0307;else if(mark==0x00a8)mark=0x0308;
        else if(mark==0x27f5)mark=0x20d6;else if(mark==0x27f6)mark=0x20d7;else if(mark==0x27f7)mark=0x20e1;
        if(mark==0x0307){setGlyph(node.Get(),L".");continue;}
        if(mark==0x23dc){setGlyph(node.Get(),L"\u2322");continue;}
        const auto* accent=mathAccent(mark);
        if(accent&&accent->layers>1){
            ComPtr<IXMLDOMNode> mover,base;
            check(node->get_parentNode(&mover),"MathML accent parent");
            check(node->selectSingleNode(Value(L"preceding-sibling::*[1]").v.bstrVal,&base),"MathML accent base");
            const std::wstring glyph(1,accent->symbol);
            for(unsigned layer=1;layer<accent->layers;++layer)base=wrap(base,glyph);
            check(mover->insertBefore(base.Get(),Value(node.Get()).v,nullptr),"Nest repeated MathML accents");
            setGlyph(node.Get(),glyph);continue;
        }
        const bool leftHarpoon=mark==0x20d0||mark==0x21bc;
        if(mark==0x0308||mark==0x20db||mark==0x20dc||leftHarpoon){
            ComPtr<IXMLDOMNode> base;
            check(node->selectSingleNode(Value(L"preceding-sibling::*[1]").v.bstrVal,&base),"MathML accent base");
            if(!singleAccentAtom(base)){
                // 短鱼叉只会落在基底中央；Word 的长鱼叉横跨整组。
                require(!leftHarpoon,"OMML_UNSUPPORTED_STRETCHY_HARPOON");
                setGlyph(node.Get(),std::wstring(mark==0x0308?2:mark==0x20db?3:4,L'.'));continue;
            }
        }
        if(accent)setGlyph(node.Get(),std::wstring(1,accent->symbol));
    }
}
}
std::wstring ommlSource(IDispatch* range) {
    auto dom=xmlDocument(text(range,L"WordOpenXML"));ComPtr<IXMLDOMNode> equation;
    check(dom->selectSingleNode(Value(L"//*[namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/math' and local-name()='oMath']").v.bstrVal,&equation),"Select OMML content snapshot");require(equation!=nullptr,"OMML snapshot exists");
    return ommlNodeSource(equation.Get());
}
bool ommlSourcesFromBodyXml(const std::wstring& package,size_t wordCount,std::vector<std::wstring>& sources,bool paragraphPackages) {
    auto dom=xmlDocument(package);
    check(dom->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(
        L"xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage' "
        L"xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' "
        L"xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' "
        L"xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'").v),"Set body OMML namespaces");
    ComPtr<IXMLDOMNodeList> bodies;check(dom->selectNodes(Value(L"/pkg:package/pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body").v.bstrVal,&bodies),"Select main story OMML body");
    long bodyCount{};check(bodies->get_length(&bodyCount),"Main story body count");if(bodyCount!=1)return false;
    ComPtr<IXMLDOMNode> body;check(bodies->get_item(0,&body),"Main story body");
    // 兼容分支和修订可能保留多种物理表示，不能把 XML 的物理顺序冒充 Word 的逻辑清单。
    ComPtr<IXMLDOMNode> ambiguous;check(body->selectSingleNode(Value(
        L".//mc:AlternateContent[.//m:oMath[not(ancestor::w:txbxContent)]] | "
        L".//*[self::w:ins or self::w:del or self::w:moveFrom or self::w:moveTo]"
        L"[ancestor::m:oMath or descendant::m:oMath][not(ancestor::w:txbxContent)]").v.bstrVal,&ambiguous),"Check unambiguous body OMML representation");
    if(ambiguous)return false;
    ComPtr<IXMLDOMNodeList> equations;check(body->selectNodes(Value(L".//m:oMath[not(ancestor::w:txbxContent)]").v.bstrVal,&equations),"Select main story equations without textboxes");
    long count{};check(equations->get_length(&count),"Body OMML count");if(static_cast<size_t>(count)!=wordCount)return false;
    std::vector<std::wstring> result;result.reserve(wordCount);
    std::wstring documentStart;
    if(paragraphPackages){
        ComPtr<IXMLDOMNode> document;check(body->get_parentNode(&document),"Body document element");
        Value xml;xml.v.vt=VT_BSTR;check(document->get_xml(&xml.v.bstrVal),"Body document namespaces");
        const std::wstring value=xml.v.bstrVal;documentStart=value.substr(0,value.find(L'>')+1);
    }
    for(long index=0;index<count;++index){
        ComPtr<IXMLDOMNode> equation;check(equations->get_item(index,&equation),"Body OMML equation");
        if(!paragraphPackages){result.push_back(ommlNodeSource(equation.Get()));continue;}
        // 快速输入只缓存单式普通段落，连同原文和 pPr 保留给局部恢复；其他结构继续逐范围读取。
        ComPtr<IXMLDOMNode> paragraph;check(equation->selectSingleNode(Value(L"ancestor::w:p[1][count(.//m:oMath)=1][not(.//w:fldChar or .//w:fldSimple or .//w:drawing or .//w:pict or .//w:object or .//w:bookmarkStart or .//w:bookmarkEnd or .//w:sdt or .//w:hyperlink)]").v.bstrVal,&paragraph),"Standalone OMML input paragraph");
        if(!paragraph){result.emplace_back();continue;}
        Value xml;xml.v.vt=VT_BSTR;check(paragraph->get_xml(&xml.v.bstrVal),"OMML original paragraph XML");
        result.push_back(L"<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData>"+
            documentStart+L"<w:body>"+xml.v.bstrVal+L"</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>");
    }
    sources=std::move(result);return true;
}
std::wstring ommlRecoveryPackage(const std::wstring& originalPackage,const std::wstring& paragraphPackage) {
    auto original=xmlDocument(originalPackage),source=xmlDocument(paragraphPackage);
    const auto bodyPath=L"/*[local-name()='package']/*[local-name()='part' and @*[local-name()='name']='/word/document.xml']/*[local-name()='xmlData']/*[local-name()='document']/*[local-name()='body']";
    ComPtr<IXMLDOMNode> target,body,child;check(original->selectSingleNode(Value(bodyPath).v.bstrVal,&target),"Original recovery body");
    check(source->selectSingleNode(Value(bodyPath).v.bstrVal,&body),"Cached recovery paragraph body");require(target&&body,"Recovery package bodies");
    // 失败时才组装完整恢复包；关系、样式和主题共用批量原包，不逐式复制到输入缓存。
    while(target->get_firstChild(&child)==S_OK&&child){check(target->removeChild(child.Get(),nullptr),"Clear bulk recovery body");child.Reset();}
    ComPtr<IXMLDOMDocument3> owner;check(original.As(&owner),"Recovery DOM import interface");
    check(body->get_firstChild(&child),"Cached recovery paragraph");require(child!=nullptr,"Nonempty cached recovery paragraph");
    ComPtr<IXMLDOMNode> imported;check(owner->importNode(child.Get(),VARIANT_TRUE,&imported),"Import original recovery paragraph");
    check(target->appendChild(imported.Get(),nullptr),"Restore single paragraph body");
    Value xml;xml.v.vt=VT_BSTR;check(original->get_xml(&xml.v.bstrVal),"Serialize full recovery package");return xml.v.bstrVal;
}
ComPtr<IXSLTemplate> compileOmmlStylesheet(const std::wstring& path,ConversionTimings& timings) {
    ScopedConversionTiming timing(timings,ConversionMetric::XslCompile);
    // IXSLTemplate 要求自由线程的样式 DOM；只返回完全加载、编译成功的模板。
    ComPtr<IXMLDOMDocument2> style;check(CoCreateInstance(__uuidof(FreeThreadedDOMDocument60),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&style)),"Create Office OMML stylesheet reader");
    check(style->put_async(VARIANT_FALSE),"Synchronous Office OMML stylesheet");check(style->put_preserveWhiteSpace(VARIANT_TRUE),"Keep Office stylesheet whitespace");
    check(style->put_resolveExternals(VARIANT_FALSE),"Disable Office stylesheet external resolution");check(style->setProperty(Value(L"ProhibitDTD").v.bstrVal,Value(true).v),"Disable Office stylesheet DTD");
    VARIANT_BOOL loaded{};check(style->load(Value(path.c_str()).v,&loaded),"Load Office OMML transform");require(loaded!=VARIANT_FALSE,"OFFICE_OMML_TRANSFORM_MISSING");
    ComPtr<IXSLTemplate> compiled;check(CoCreateInstance(__uuidof(XSLTemplate60),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&compiled)),"Create compiled Office OMML transform");
    check(compiled->putref_stylesheet(style.Get()),"Compile Office OMML transform");return compiled;
}
void restoreSavedOmmlTabs(const std::wstring& privateDocx) {
    // Word 保存时把公式内的文本 Tab 写成 w:tab，重开会丢掉它。只修复私有包中公式的这些节点。
    ComPtr<IOpcFactory> factory;
    check(CoCreateInstance(__uuidof(OpcFactory),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&factory)),"OMML tab package factory");
    ComPtr<IStream> input;
    check(factory->CreateStreamOnFile(privateDocx.c_str(),OPC_STREAM_IO_READ,nullptr,0,&input),"Read saved OMML tab package");
    ComPtr<IOpcPackage> package;
    check(factory->ReadPackageFromStream(input.Get(),OPC_CACHE_ON_ACCESS,&package),"Load saved OMML tab package");
    size_t repaired{};
    {
        ComPtr<IOpcPartSet> parts;check(package->GetPartSet(&parts),"OMML tab package parts");
        ComPtr<IOpcPartEnumerator> entries;check(parts->GetEnumerator(&entries),"Enumerate OMML tab package parts");
        for(BOOL more{};check(entries->MoveNext(&more),"Next OMML tab package part"),more;){
            taskPulse();
            ComPtr<IOpcPart> part;check(entries->GetCurrent(&part),"Read OMML tab package part");
            ComPtr<IOpcPartUri> uri;check(part->GetName(&uri),"OMML tab part URI");
            Value name;name.v.vt=VT_BSTR;check(uri->GetAbsoluteUri(&name.v.bstrVal),"OMML tab part name");
            const std::wstring_view path(name.v.bstrVal);
            if(path.rfind(L"/word/",0)!=0||path.size()<4||path.substr(path.size()-4)!=L".xml")continue;
            ComPtr<IStream> content;check(part->GetContentStream(&content),"Read saved Word XML part");
            auto dom=xmlDocument(L"<root/>");
            Value source;source.v.vt=VT_UNKNOWN;source.v.punkVal=content.Get();content->AddRef();
            VARIANT_BOOL loaded{};check(dom->load(source.v,&loaded),"Load saved Word XML part");require(loaded!=VARIANT_FALSE,"Valid saved Word XML part");
            check(dom->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(
                L"xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' "
                L"xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'").v),"Set OMML tab namespaces");
            ComPtr<IXMLDOMNodeList> tabs;check(dom->selectNodes(Value(L"//m:oMath//m:r/w:tab").v.bstrVal,&tabs),"Find saved formula tabs");
            long count{};check(tabs->get_length(&count),"Count saved formula tabs");
            if(!count)continue;
            for(long i=0;i<count;++i){
                ComPtr<IXMLDOMNode> tab,parent,token;check(tabs->get_item(i,&tab),"Read saved formula tab");
                check(tab->get_parentNode(&parent),"Read formula tab run");
                check(dom->createNode(Value(1L).v,Value(L"m:t").v.bstrVal,Value(L"http://schemas.openxmlformats.org/officeDocument/2006/math").v.bstrVal,&token),"Create formula text tab");
                ComPtr<IXMLDOMElement> element;check(token.As(&element),"Read formula text tab element");
                check(element->setAttribute(Value(L"xml:space").v.bstrVal,Value(L"preserve").v),"Preserve formula text whitespace");
                check(token->put_text(Value(L"\t").v.bstrVal),"Set literal formula tab");
                check(parent->replaceChild(token.Get(),tab.Get(),nullptr),"Restore formula text tab");
            }
            ULARGE_INTEGER empty{};LARGE_INTEGER zero{};
            check(content->SetSize(empty),"Replace saved Word XML part");check(content->Seek(zero,STREAM_SEEK_SET,nullptr),"Rewind saved Word XML part");
            check(dom->save(source.v),"Write restored formula tabs");
            repaired+=static_cast<size_t>(count);
        }
    }
    if(!repaired)return;
    ComPtr<IStream> output;check(CreateStreamOnHGlobal(nullptr,TRUE,&output),"Buffer restored OMML tab package");
    check(factory->WritePackageToStream(package.Get(),OPC_WRITE_DEFAULT,output.Get()),"Serialize restored OMML tab package");package.Reset();input.Reset();
    LARGE_INTEGER zero{};check(output->Seek(zero,STREAM_SEEK_SET,nullptr),"Rewind restored OMML tab package");
    ComPtr<IStream> destination;check(factory->CreateStreamOnFile(privateDocx.c_str(),OPC_STREAM_IO_WRITE,nullptr,0,&destination),"Replace private OMML tab package");
    ULARGE_INTEGER all{};all.QuadPart=~0ULL;check(output->CopyTo(destination.Get(),all,nullptr,nullptr),"Complete restored OMML tab package write");
    std::printf("OMML_TABS_RESTORED=%zu\n",repaired);
}
std::wstring ommlMathML(const std::wstring& sourcePackage,IXSLTemplate* stylesheet,bool display,ConversionTimings& timings,bool compactRows) {
    ComPtr<IXMLDOMDocument2> source;ComPtr<IXMLDOMNode> equation;
    {
    ScopedConversionTiming timing(timings,ConversionMetric::XmlPrepare);source=xmlDocument(sourcePackage);
    check(source->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(
        L"xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' "
        L"xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'").v),"Set source OMML namespaces");
    check(source->selectSingleNode(Value(L"//m:oMath").v.bstrVal,&equation),"Select source OMML");require(equation!=nullptr,"Source OMML exists");
    prepareOmmlWhitespace(equation.Get());
    }
    // 每式 processor 隔离转换状态，样式编译结果由本阶段的模板复用。
    Value transformed;
    {
    ScopedConversionTiming timing(timings,ConversionMetric::XslTransform);
    ComPtr<IXSLProcessor> processor;check(stylesheet->createProcessor(&processor),"Create OMML transform processor");
    check(processor->put_input(Value(equation.Get()).v),"Set source OMML transform node");
    VARIANT_BOOL complete{};check(processor->transform(&complete),"Transform OMML to MathML");require(complete!=VARIANT_FALSE,"Complete synchronous OMML transform");
    check(processor->get_output(&transformed.v),"Read transformed MathML");require(transformed.v.vt==VT_BSTR&&transformed.v.bstrVal,"MathML transform returns XML text");
    }
    return finishMathML(transformed.v.bstrVal,display,timings,compactRows,equation.Get());
}
std::wstring ommlGroupBraceTex(const std::wstring& sourcePackage) {
    // MathML 导入会把整组花括号缩成单个字符。仅在确有组括号时尝试完整的 TeX 导入。
    if(sourcePackage.find(L"groupChr")==std::wstring::npos)return {};
    auto source=xmlDocument(sourcePackage);
    ComPtr<IXMLDOMNode> equation;
    check(source->selectSingleNode(Value(L"//*[namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/math' and local-name()='oMath']").v.bstrVal,&equation),"Select source OMML group");
    std::wstring tex;
    if(!tryOmmlTex(equation.Get(),tex,true))return {};
    return tex;
}
std::wstring finishMathML(std::wstring fragment,bool display,ConversionTimings& timings,bool compactRows,IXMLDOMNode* sourceEquation) {
    // Office 的模板输出 MathML 片段，必须补上 math 根节点再交给厂商导入。
    ScopedConversionTiming timing(timings,ConversionMetric::MathmlAdapt);
    if(fragment.rfind(L"<?xml",0)==0){const auto end=fragment.find(L"?>");require(end!=std::wstring::npos,"Office MathML declaration");fragment.erase(0,end+2);}
    auto math=xmlDocument(std::wstring(L"<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"")+(display?L"block":L"inline")+L"\">"+fragment+L"</math>");
    compatibleMathML(math.Get(),sourceEquation);
    if(compactRows){
        // MathML 3 §3.3.1：无属性、单参数 mrow 与该参数等价；从内向外删除，保留分组和样式。
        ComPtr<IXMLDOMNodeList> rows;check(math->selectNodes(Value(L"//*[namespace-uri()='http://www.w3.org/1998/Math/MathML' and local-name()='mrow' and not(@*) and count(*)=1]").v.bstrVal,&rows),"Select redundant MathML rows");
        long count{};check(rows->get_length(&count),"Redundant MathML row count");
        for(long i=count;i>0;--i){ComPtr<IXMLDOMNode> row,child,parent;check(rows->get_item(i-1,&row),"Redundant MathML row");
            check(row->selectSingleNode(Value(L"*").v.bstrVal,&child),"Single MathML row argument");check(row->get_parentNode(&parent),"MathML row parent");
            check(parent->replaceChild(child.Get(),row.Get(),nullptr),"Remove redundant MathML row");}
    }
    Value xml;xml.v.vt=VT_BSTR;check(math->get_xml(&xml.v.bstrVal),"Compatible MathML XML");return xml.v.bstrVal;
}
