// Office OMML 到 MathType 可导入 MathML 的共同转换路径。
#include "omml_adapter.h"
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

std::wstring ommlNodeSource(IXMLDOMNode* equation) {
    Value xml;xml.v.vt=VT_BSTR;check(equation->get_xml(&xml.v.bstrVal),"OMML snapshot XML");return withoutWordRevisionIds(xml.v.bstrVal);
}

void compatibleMathML(IXMLDOMDocument2* dom) {
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
    // 原 XML 可能用字符实体表示顶饰，只能在 DOM 解码后判断 ASCII 快路径。
    if(!needsMathMLCompatibility(content(dom)))return;
    auto wrap=[&](ComPtr<IXMLDOMNode> base,wchar_t mark){
        auto mover=element(L"mover"),symbol=element(L"mo");ComPtr<IXMLDOMElement> attributes;check(mover.As(&attributes),"MathML accent attributes");
        check(attributes->setAttribute(Value(L"accent").v.bstrVal,Value(L"true").v),"Set MathML accent");
        const std::wstring glyph(1,mathAccent(mark)->symbol);check(symbol->put_text(Value(glyph.c_str()).v.bstrVal),"Set MathML accent glyph");
        append(mover.Get(),base.Get());append(mover.Get(),symbol.Get());return mover;
    };
    // 只改顶饰位置的符号码，不能把正文中的减号、波浪号等全局替换。
    for(const auto& node:nodes(L"//*[local-name()='mover']/*[2][local-name()='mo' or local-name()='mi']")) {
        const auto text=content(node.Get());if(text.size()!=1)continue;auto mark=text.front();
        if(mark==0x02c6)mark=0x0302;else if(mark==0x02dc)mark=0x0303;else if(mark==0x02d8)mark=0x0306;
        else if(mark==0x00af)mark=0x0304;else if(mark==0x02d9)mark=0x0307;else if(mark==0x00a8)mark=0x0308;
        if(const auto* accent=mathAccent(mark)){const std::wstring glyph(1,accent->symbol);check(node->put_text(Value(glyph.c_str()).v.bstrVal),"Normalize MathML accent glyph");}
    }
    const auto tokens=nodes(L"//*[namespace-uri()='http://www.w3.org/1998/Math/MathML' and (local-name()='mi' or local-name()='mn' or local-name()='mo')][not(parent::*[local-name()='mover' or local-name()='munder' or local-name()='munderover'] and count(preceding-sibling::*)>=1)]");
    // Office 会把 Y 与后面的组合符拆成相邻 token；先合回，随后仅装饰最后一个字符。
    for(const auto& node:tokens) {
        const auto marks=content(node.Get());if(marks.empty()||!std::all_of(marks.begin(),marks.end(),[](wchar_t c){return mathAccent(c)!=nullptr;}))continue;
        ComPtr<IXMLDOMNode> previous,parent;check(node->get_parentNode(&parent),"MathML token parent");if(!parent)continue;
        check(node->get_previousSibling(&previous),"Previous MathML token");
        while(previous&&name(previous.Get()).empty()&&whitespaceOnly(content(previous.Get()))){ComPtr<IXMLDOMNode> before;check(previous->get_previousSibling(&before),"Previous MathML whitespace");previous=before;}
        if(!previous)continue;const auto tag=name(previous.Get());if(tag!=L"mi"&&tag!=L"mn"&&tag!=L"mo")continue;
        const auto combined=content(previous.Get())+marks;check(previous->put_text(Value(combined.c_str()).v.bstrVal),"Join MathML combining accents");
        ComPtr<IXMLDOMNode> removed;check(parent->removeChild(node.Get(),&removed),"Remove joined MathML accent token");
    }
    for(const auto& node:tokens) {
        ComPtr<IXMLDOMNode> parent;check(node->get_parentNode(&parent),"MathML letter parent");if(!parent)continue;
        const auto original=content(node.Get());std::wstring decomposed;for(const auto character:original)decomposed+=decomposedMathCharacter(character);
        if(decomposed.size()<2||std::none_of(decomposed.begin()+1,decomposed.end(),[](wchar_t c){return mathAccent(c)!=nullptr;}))continue;
        auto row=element(L"mrow");std::wstring plain;
        auto clone=[&](const std::wstring& text){ComPtr<IXMLDOMNode> token;check(node->cloneNode(VARIANT_FALSE,&token),"Clone styled MathML token");check(token->put_text(Value(text.c_str()).v.bstrVal),"Set MathML base text");return token;};
        auto flush=[&]{if(!plain.empty()){auto token=clone(plain);append(row.Get(),token.Get());plain.clear();}};
        for(size_t i=0;i<decomposed.size();) {
            size_t finish=i+1;if(decomposed[i]>=0xd800&&decomposed[i]<=0xdbff&&finish<decomposed.size()&&decomposed[finish]>=0xdc00&&decomposed[finish]<=0xdfff)++finish;
            const auto base=decomposed.substr(i,finish-i);const auto firstMark=finish;
            while(finish<decomposed.size()&&mathAccent(decomposed[finish]))++finish;
            if(finish==firstMark)plain+=base;
            else {flush();auto token=clone(base);for(size_t j=firstMark;j<finish;++j)token=wrap(token,decomposed[j]);append(row.Get(),token.Get());}
            i=finish;
        }
        flush();ComPtr<IXMLDOMNode> replaced;check(parent->replaceChild(row.Get(),node.Get(),&replaced),"Replace MathML accented token");
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
std::wstring ommlMathML(const std::wstring& sourcePackage,IXSLTemplate* stylesheet,bool display,ConversionTimings& timings,bool compactRows) {
    ComPtr<IXMLDOMDocument2> source;ComPtr<IXMLDOMNode> equation;
    {
    ScopedConversionTiming timing(timings,ConversionMetric::XmlPrepare);source=xmlDocument(sourcePackage);
    check(source->selectSingleNode(Value(L"//*[namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/math' and local-name()='oMath']").v.bstrVal,&equation),"Select source OMML");require(equation!=nullptr,"Source OMML exists");
    // 显式空白会在 MathType 中产生占位符；按产品约定统一删除（含普通文本），使用自动间距。
    ComPtr<IXMLDOMNodeList> texts;
    check(equation->selectNodes(Value(L".//*[namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/math' and local-name()='t']").v.bstrVal,&texts),"Select OMML text");
    long count{};check(texts->get_length(&count),"OMML text count");
    for(long i=0;i<count;++i){
        ComPtr<IXMLDOMNode> node;check(texts->get_item(i,&node),"OMML text node");
        Value value;value.v.vt=VT_BSTR;check(node->get_text(&value.v.bstrVal),"OMML text value");
        std::wstring content=value.v.bstrVal;
        content.erase(std::remove_if(content.begin(),content.end(),[](wchar_t c){return std::wstring_view(L" \u00a0\t\n\r").find(c)!=std::wstring_view::npos;}),content.end());
        if(content!=value.v.bstrVal)check(node->put_text(Value(content.c_str()).v.bstrVal),"Remove OMML whitespace");
    }
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
    return finishMathML(transformed.v.bstrVal,display,timings,compactRows);
}
std::wstring finishMathML(std::wstring fragment,bool display,ConversionTimings& timings,bool compactRows) {
    // Office 的模板输出 MathML 片段，必须补上 math 根节点再交给厂商导入。
    ScopedConversionTiming timing(timings,ConversionMetric::MathmlAdapt);
    if(fragment.rfind(L"<?xml",0)==0){const auto end=fragment.find(L"?>");require(end!=std::wstring::npos,"Office MathML declaration");fragment.erase(0,end+2);}
    auto math=xmlDocument(std::wstring(L"<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"")+(display?L"block":L"inline")+L"\">"+fragment+L"</math>");
    compatibleMathML(math.Get());
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
