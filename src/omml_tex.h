// 仅处理有明确 TeX 映射的 OMML；未知结构交回现有 MathML 路径，不改动源 DOM。
#pragma once
#include "omml_adapter.h"
#include "latex_rules.h"
#include <map>

namespace omml_tex_detail {
constexpr const wchar_t* mathNamespace=L"http://schemas.openxmlformats.org/officeDocument/2006/math";
constexpr const wchar_t* wordNamespace=L"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
inline std::wstring name(IXMLDOMNode* node){Value value;value.v.vt=VT_BSTR;check(node->get_baseName(&value.v.bstrVal),"OMML TeX node name");return value.v.bstrVal?value.v.bstrVal:L"";}
inline std::wstring uri(IXMLDOMNode* node){Value value;value.v.vt=VT_BSTR;check(node->get_namespaceURI(&value.v.bstrVal),"OMML TeX namespace");return value.v.bstrVal?value.v.bstrVal:L"";}
inline std::wstring text(IXMLDOMNode* node){Value value;value.v.vt=VT_BSTR;check(node->get_text(&value.v.bstrVal),"OMML TeX text");return value.v.bstrVal?value.v.bstrVal:L"";}
inline std::vector<ComPtr<IXMLDOMNode>> children(IXMLDOMNode* node){
    std::vector<ComPtr<IXMLDOMNode>> result;if(!node)return result;
    ComPtr<IXMLDOMNodeList> list;check(node->selectNodes(Value(L"*").v.bstrVal,&list),"OMML TeX children");long count{};check(list->get_length(&count),"OMML TeX child count");
    for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> child;check(list->get_item(i,&child),"OMML TeX child");result.push_back(std::move(child));}return result;
}
inline ComPtr<IXMLDOMNode> child(IXMLDOMNode* node,const wchar_t* wanted){
    ComPtr<IXMLDOMNode> result;for(auto& item:children(node))if(name(item.Get())==wanted&&uri(item.Get())==mathNamespace)result=item;return result;
}
inline std::wstring value(IXMLDOMNode* node,const wchar_t* fallback=L""){
    if(!node)return fallback;ComPtr<IXMLDOMNode> attr;check(node->selectSingleNode(Value(L"@*[local-name()='val']").v.bstrVal,&attr),"OMML TeX property value");return attr?text(attr.Get()):fallback;
}
inline bool has(const std::wstring& item,std::initializer_list<const wchar_t*> names){return std::any_of(names.begin(),names.end(),[&](const wchar_t* candidate){return item==candidate;});}
inline std::wstring group(const std::wstring& item){return L"{"+item+L"}";}

class Compiler {
    bool valid=true;
    std::wstring reject(){valid=false;return {};}
    bool wordProperties(IXMLDOMNode* properties){
        for(auto& item:children(properties)){
            if(uri(item.Get())!=wordNamespace)return false;const auto tag=name(item.Get());
            if(tag==L"lang"||tag==L"sz"||tag==L"szCs")continue; // 调用方使用固定字号，语言标签不改变数学内容。
            if(tag!=L"rFonts")return false;
            ComPtr<IXMLDOMNamedNodeMap> attrs;check(item->get_attributes(&attrs),"OMML TeX font attributes");long count{};check(attrs->get_length(&count),"OMML TeX font attribute count");
            for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> attr;check(attrs->get_item(i,&attr),"OMML TeX font attribute");const auto key=name(attr.Get());
                if(has(key,{L"ascii",L"hAnsi",L"eastAsia",L"cs"})&&text(attr.Get())==L"Cambria Math")continue;
                if(key==L"hint")continue;return false;}
        }return true;
    }
    bool control(IXMLDOMNode* properties){for(auto& item:children(properties))if(uri(item.Get())!=wordNamespace||name(item.Get())!=L"rPr"||!wordProperties(item.Get()))return false;return true;}
    bool allowed(IXMLDOMNode* node,std::initializer_list<const wchar_t*> names){
        for(auto& item:children(node))if(uri(item.Get())!=mathNamespace||!has(name(item.Get()),names)){valid=false;return false;}return true;
    }
    ComPtr<IXMLDOMNode> properties(IXMLDOMNode* node,const wchar_t* kind,std::initializer_list<const wchar_t*> names){
        auto props=child(node,kind);for(auto& item:children(props.Get())){
            const auto tag=name(item.Get());if(uri(item.Get())!=mathNamespace||(tag==L"ctrlPr"?!control(item.Get()):!has(tag,names)))valid=false;
        }return props;
    }
    bool flag(IXMLDOMNode* properties,const wchar_t* key,bool fallback=false){
        auto item=child(properties,key);if(!item)return fallback;const auto val=value(item.Get(),L"1");
        if(has(val,{L"1",L"on",L"true"}))return true;if(has(val,{L"0",L"off",L"false"}))return false;valid=false;return false;
    }
    std::wstring character(IXMLDOMNode* properties,const wchar_t* key,const wchar_t* fallback){auto item=child(properties,key);return item?value(item.Get()):fallback;}
    std::wstring argument(IXMLDOMNode* node,const wchar_t* key){auto item=child(node,key);if(!item)return reject();return emit(item.Get());}
    std::wstring sequence(IXMLDOMNode* node){std::wstring out;for(auto& item:children(node)){if(!valid)break;const auto tag=name(item.Get());
        if(tag==L"argPr"){if(!children(item.Get()).empty())return reject();continue;}out+=emit(item.Get());}return out;}
    std::wstring symbol(wchar_t c){
        if((c>=L'a'&&c<=L'z')||(c>=L'A'&&c<=L'Z')||(c>=L'0'&&c<=L'9')||std::wstring_view(L"+-=.,:;!?/()[]|<>").find(c)!=std::wstring_view::npos)return std::wstring(1,c);
        if(std::wstring_view(L"{}_%&#$").find(c)!=std::wstring_view::npos)return std::wstring(L"\\")+c;
        if(c==L'\\')return L"\\backslash ";if(c==L'*')return L"\\ast ";if(c==0x2212)return L"-";
        static const std::map<int,const wchar_t*> symbols={
            {0x03b1,L"alpha"},{0x03b2,L"beta"},{0x03b3,L"gamma"},{0x03b4,L"delta"},{0x03b5,L"varepsilon"},{0x03b6,L"zeta"},{0x03b7,L"eta"},{0x03b8,L"theta"},{0x03b9,L"iota"},{0x03ba,L"kappa"},{0x03bb,L"lambda"},{0x03bc,L"mu"},{0x03bd,L"nu"},{0x03be,L"xi"},{0x03c0,L"pi"},{0x03c1,L"rho"},{0x03c2,L"varsigma"},{0x03c3,L"sigma"},{0x03c4,L"tau"},{0x03c5,L"upsilon"},{0x03c6,L"varphi"},{0x03c7,L"chi"},{0x03c8,L"psi"},{0x03c9,L"omega"},{0x03d1,L"vartheta"},{0x03d5,L"phi"},{0x03d6,L"varpi"},{0x03f1,L"varrho"},{0x03f5,L"epsilon"},
            {0x0393,L"Gamma"},{0x0394,L"Delta"},{0x0398,L"Theta"},{0x039b,L"Lambda"},{0x039e,L"Xi"},{0x03a0,L"Pi"},{0x03a3,L"Sigma"},{0x03a5,L"Upsilon"},{0x03a6,L"Phi"},{0x03a8,L"Psi"},{0x03a9,L"Omega"},
            {0x00b1,L"pm"},{0x2213,L"mp"},{0x00d7,L"times"},{0x00f7,L"div"},{0x22c5,L"cdot"},{0x2217,L"ast"},{0x22c6,L"star"},{0x2260,L"ne"},{0x2264,L"le"},{0x2265,L"ge"},{0x2248,L"approx"},{0x2261,L"equiv"},{0x221e,L"infty"},{0x2202,L"partial"},{0x2207,L"nabla"},{0x2208,L"in"},{0x2209,L"notin"},{0x2200,L"forall"},{0x2203,L"exists"},{0x2192,L"to"},{0x2190,L"leftarrow"},{0x2194,L"leftrightarrow"},{0x21d2,L"Rightarrow"},{0x21d4,L"Leftrightarrow"},{0x22a5,L"perp"},{0x2225,L"parallel"},{0x2229,L"cap"},{0x222a,L"cup"},{0x2282,L"subset"},{0x2286,L"subseteq"},{0x2026,L"ldots"},{0x22ef,L"cdots"},{0x22ee,L"vdots"},{0x22f1,L"ddots"},{0x2113,L"ell"}};
        const auto found=symbols.find(c);if(found==symbols.end())return reject();return L"\\"+std::wstring(found->second)+L" ";
    }
    std::wstring run(IXMLDOMNode* node){
        auto props=properties(node,L"rPr",{L"sty",L"scr",L"nor"});std::wstring input;
        for(auto& item:children(node)){const auto tag=name(item.Get());const auto space=uri(item.Get());
            if(space==wordNamespace&&tag==L"rPr"){if(!wordProperties(item.Get()))return reject();continue;}
            if(space!=mathNamespace)return reject();if(tag==L"rPr")continue;if(tag!=L"t")return reject();input+=text(item.Get());}
        input.erase(std::remove_if(input.begin(),input.end(),[](wchar_t c){return std::wstring_view(L" \t\r\n\u00a0").find(c)!=std::wstring_view::npos;}),input.end());
        const auto sty=value(child(props.Get(),L"sty").Get()),scr=value(child(props.Get(),L"scr").Get());const bool normal=flag(props.Get(),L"nor");
        if(normal){if(!std::all_of(input.begin(),input.end(),[](wchar_t c){return c<128;}))return reject();std::wstring literal;for(wchar_t c:input){if(c==L'\\'||c==L'^'||c==L'~')return reject();if(std::wstring_view(L"{}_%&#$").find(c)!=std::wstring_view::npos)literal+=L'\\';literal+=c;}return L"\\text"+group(literal);}
        if(!has(sty,{L"",L"p",L"i",L"b",L"bi"}))return reject();
        if(!has(scr,{L"",L"roman",L"script"}))return reject();if(scr==L"script"&&sty==L"b")return reject();
        // Office 按字符类别施加 mathvariant：数字和运算符不能随整个 run 一并变粗。
        std::wstring out;for(wchar_t c:input){auto atom=symbol(c);if(!valid)return {};
            const bool latin=(c>=L'a'&&c<=L'z')||(c>=L'A'&&c<=L'Z'),greek=c>=0x0391&&c<=0x03f5;
            if(latin||greek||c==0x2113){
                // 当前 MathType 的 TeX 字体命令不能同时保留粗体和斜体，整式交回 MathML。
                if(sty==L"bi")return reject();
                if(scr==L"script"){if(c<L'A'||c>L'Z')return reject();atom=L"\\mathcal"+group(atom);}
                else if(sty==L"p"||sty==L"b"){if(!latin)return reject();atom=std::wstring(sty==L"p"?L"\\mathrm":L"\\mathbf")+group(atom);}
                else if(sty==L"i"&&greek&&c<0x03b1)return reject();
            }out+=atom;
        }return out;
    }
    std::wstring delimiter(const std::wstring& value){if(value.empty())return L".";if(value.size()!=1)return reject();const auto c=value.front();
        if(std::wstring_view(L"()[]|").find(c)!=std::wstring_view::npos)return value;if(c==L'{'||c==L'}')return L"\\"+value;
        if(c==0x2016||c==0x2225)return L"\\Vert ";if(c==0x27e8||c==0x2329)return L"\\langle ";if(c==0x27e9||c==0x232a)return L"\\rangle ";return reject();}
    std::wstring matrix(IXMLDOMNode* node,bool array){
        auto props=properties(node,array?L"eqArrPr":L"mPr",array?std::initializer_list<const wchar_t*>{}:std::initializer_list<const wchar_t*>{L"mcs",L"plcHide"});
        if(child(props.Get(),L"plcHide")&&!flag(props.Get(),L"plcHide"))return reject();
        std::wstring columns;auto mcs=child(props.Get(),L"mcs");
        for(auto& mc:children(mcs.Get())){if(name(mc.Get())!=L"mc"||!allowed(mc.Get(),{L"mcPr"}))return reject();auto cp=properties(mc.Get(),L"mcPr",{L"count",L"mcJc"});
            const auto count=value(child(cp.Get(),L"count").Get(),L"1"),alignment=value(child(cp.Get(),L"mcJc").Get(),L"center");
            if(count.empty()||count.size()>3||!std::all_of(count.begin(),count.end(),[](wchar_t c){return c>=L'0'&&c<=L'9';}))return reject();
            const auto amount=std::stoul(count);if(!amount||amount>100||!has(alignment,{L"center",L"left",L"right"}))return reject();columns.append(amount,alignment==L"left"?L'l':alignment==L"right"?L'r':L'c');}
        size_t width{};std::wstring body;bool first=true;
        for(auto& row:children(node)){const auto tag=name(row.Get());if(tag==(array?L"eqArrPr":L"mPr"))continue;
            if(uri(row.Get())!=mathNamespace||tag!=(array?L"e":L"mr"))return reject();if(!first)body+=L"\\\\";first=false;
            if(array){if(text(row.Get()).find(L'&')!=std::wstring::npos)return reject();body+=emit(row.Get());width=1;continue;}
            const auto cells=children(row.Get());if(cells.empty()||(width&&width!=cells.size()))return reject();width=cells.size();
            for(size_t i=0;i<cells.size();++i){if(name(cells[i].Get())!=L"e")return reject();if(i)body+=L'&';body+=emit(cells[i].Get());}
        }
        if(!width)return reject();if(columns.empty())columns.assign(width,L'c');if(columns.size()!=width)return reject();
        return L"\\begin{array}"+group(columns)+body+L"\\end{array}";
    }
    std::wstring emit(IXMLDOMNode* node){
        if(!valid)return {};if(uri(node)!=mathNamespace)return reject();const auto tag=name(node);
        if(has(tag,{L"oMath",L"e",L"num",L"den",L"sub",L"sup",L"deg"}))return sequence(node);
        if(tag==L"r")return run(node);
        if(has(tag,{L"sSub",L"sSup",L"sSubSup"})){
            const auto kind=tag+L"Pr";if(!allowed(node,{L"e",L"sub",L"sup",L"sSubPr",L"sSupPr",L"sSubSupPr"}))return {};properties(node,kind.c_str(),{});
            auto out=group(argument(node,L"e"));if(tag!=L"sSup")out+=L"_"+group(argument(node,L"sub"));if(tag!=L"sSub")out+=L"^"+group(argument(node,L"sup"));return out;
        }
        if(tag==L"f"){
            if(!allowed(node,{L"fPr",L"num",L"den"}))return {};auto props=properties(node,L"fPr",{L"type"});const auto type=value(child(props.Get(),L"type").Get(),L"bar");
            const auto numerator=argument(node,L"num"),denominator=argument(node,L"den");if(type==L"bar")return L"\\frac"+group(numerator)+group(denominator);
            if(type==L"lin")return group(numerator)+L"/"+group(denominator);return reject();
        }
        if(tag==L"acc"||tag==L"bar"){
            if(!allowed(node,{L"accPr",L"barPr",L"e"}))return {};const auto body=argument(node,L"e");
            if(tag==L"bar"){auto props=properties(node,L"barPr",{L"pos"});const auto pos=value(child(props.Get(),L"pos").Get(),L"bot");if(!has(pos,{L"top",L"bot"}))return reject();return std::wstring(pos==L"top"?L"\\bar":L"\\underline")+group(body);}
            auto props=properties(node,L"accPr",{L"chr"});auto accent=character(props.Get(),L"chr",L"\u0302");if(accent.size()!=1)return reject();const auto* mark=mathAccent(accent.front());if(!mark)return reject();return std::wstring(mark->tex)+group(body);
        }
        if(tag==L"rad"){
            if(!allowed(node,{L"radPr",L"deg",L"e"}))return {};auto props=properties(node,L"radPr",{L"degHide"});const auto body=argument(node,L"e");
            if(flag(props.Get(),L"degHide"))return L"\\sqrt"+group(body);return L"\\sqrt["+group(argument(node,L"deg"))+L"]"+group(body);
        }
        if(tag==L"d"){
            if(!allowed(node,{L"dPr",L"e"}))return {};auto props=properties(node,L"dPr",{L"begChr",L"endChr",L"sepChr",L"grow"});
            if(child(props.Get(),L"grow")&&!flag(props.Get(),L"grow"))return reject();
            const auto left=delimiter(character(props.Get(),L"begChr",L"(")),right=delimiter(character(props.Get(),L"endChr",L")"));
            const auto sep=character(props.Get(),L"sepChr",L"|");if(sep.size()>1)return reject();const auto separator=sep.empty()?L"":symbol(sep.front());
            std::wstring out=L"\\left"+left+L" ";bool first=true;for(auto& item:children(node))if(name(item.Get())==L"e"){if(!first)out+=separator;first=false;out+=emit(item.Get());}return out+L"\\right"+right+L" ";
        }
        if(tag==L"m"||tag==L"eqArr")return matrix(node,tag==L"eqArr");
        if(tag==L"nary"){
            if(!allowed(node,{L"naryPr",L"e",L"sub",L"sup"}))return {};auto props=properties(node,L"naryPr",{L"chr",L"limLoc",L"grow",L"subHide",L"supHide"});
            const auto chr=character(props.Get(),L"chr",L"\u222b"),location=value(child(props.Get(),L"limLoc").Get(),L"subSup");
            if(!has(location,{L"subSup",L"undOvr"}))return reject();const bool grow=flag(props.Get(),L"grow");
            const wchar_t* command=chr==L"\u2211"?L"\\sum":chr==L"\u220f"?L"\\prod":chr==L"\u222b"?L"\\int":chr==L"\u222c"?L"\\iint":chr==L"\u222d"?L"\\iiint":chr==L"\u222e"?L"\\oint":nullptr;
            if(!command||!grow)return reject();std::wstring out=command;out+=location==L"subSup"?L"\\nolimits":L"\\limits";
            if(!flag(props.Get(),L"subHide"))out+=L"_"+group(argument(node,L"sub"));if(!flag(props.Get(),L"supHide"))out+=L"^"+group(argument(node,L"sup"));return out+group(argument(node,L"e"));
        }
        return reject();
    }
public:
    bool convert(IXMLDOMNode* equation,std::wstring& result){if(!equation||name(equation)!=L"oMath")return false;auto out=emit(equation);if(!valid||out.empty())return false;result=std::move(out);return true;}
};
}

// 调用方必须指定固定字号；不支持的公式不产生部分 TeX，继续走原有 MathML 导入。
inline bool tryOmmlTex(IXMLDOMNode* equation,std::wstring& tex){return omml_tex_detail::Compiler().convert(equation,tex);}
