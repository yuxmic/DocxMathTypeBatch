// 离线 MSXML 回归：调用正式适配器，不启动 Word 或 MathType，不写文档。
#pragma once
#include "omml_adapter.h"
#include "docx_package.h"
#include "conversion_timing.h"

namespace omml_adapter_tests {
inline std::wstring source(const std::wstring& math) {
    // 外围文本和第二式用于发现 processor 意外从整个 ownerDocument 开始转换。
    return L"<root xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><outside>OUTSIDE_BEFORE</outside><m:oMath>"+
        math+L"</m:oMath><outside>OUTSIDE_AFTER</outside><m:oMath><m:r><m:t>UNSELECTED_FORMULA</m:t></m:r></m:oMath></root>";
}
inline long count(const std::wstring& xml,const wchar_t* query) {
    auto dom=xmlDocument(xml);ComPtr<IXMLDOMNodeList> nodes;check(dom->selectNodes(Value(query).v.bstrVal,&nodes),"Test MathML nodes");
    long result{};check(nodes->get_length(&result),"Test MathML node count");return result;
}
inline std::wstring contents(const std::wstring& xml) {
    auto dom=xmlDocument(xml);Value value;value.v.vt=VT_BSTR;check(dom->get_text(&value.v.bstrVal),"Test XML text");return value.v.bstrVal?value.v.bstrVal:L"";
}
inline std::wstring legacy(IXMLDOMDocument2* stylesheet,const std::wstring& input,bool display,ConversionTimings& timings) {
    auto dom=xmlDocument(input);ComPtr<IXMLDOMNode> equation;
    check(dom->selectSingleNode(Value(L"//*[local-name()='oMath' and namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/math']").v.bstrVal,&equation),"Test legacy OMML node");
    require(equation!=nullptr,"Test legacy equation exists");Value result;result.v.vt=VT_BSTR;
    check(equation->transformNode(stylesheet,&result.v.bstrVal),"Test legacy transformNode");
    // 复用生产片段包装与兼容处理，只对比两种实际 XSLT 执行路径。
    return finishMathML(result.v.bstrVal,display,timings);
}
inline void checkBodySources() {
    const std::wstring start=L"<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><pkg:part pkg:name=\"/word/document.xml\"><pkg:xmlData><w:document><w:body>";
    const std::wstring end=L"</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/header1.xml\"><pkg:xmlData><w:hdr><m:oMath><m:r><m:t>HEADER</m:t></m:r></m:oMath></w:hdr></pkg:xmlData></pkg:part></pkg:package>";
    const auto equation=[](const wchar_t* text){return L"<m:oMath w:rsidR=\"1234\"><m:r><m:t>"+std::wstring(text)+L"</m:t></m:r></m:oMath>";};
    const std::wstring body=L"<w:p>"+equation(L"x")+L"</w:p><w:tbl><w:tr><w:tc><w:p>"+equation(L"y")+
        L"</w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:drawing><w:txbxContent><w:p>"+equation(L"BOX")+
        L"</w:p></w:txbxContent></w:drawing></w:r></w:p><w:p>"+equation(L"z")+L"</w:p>";
    std::vector<std::wstring> values;
    require(ommlSourcesFromBodyXml(start+body+end,3,values)&&values.size()==3,"Body OMML retains document order and excludes other regions");
    for(size_t i=0;i<values.size();++i)require(contents(values[i])==std::wstring(1,L"xyz"[i])&&values[i].find(L"rsid")==std::wstring::npos,"Body OMML source and stable metadata");
    require(!ommlSourcesFromBodyXml(start+body+end,2,values),"Body OMML count mismatch uses established fallback");
    const std::wstring alternate=L"<mc:AlternateContent><mc:Choice Requires=\"w\"><w:p>"+equation(L"x")+L"</w:p></mc:Choice></mc:AlternateContent>";
    require(!ommlSourcesFromBodyXml(start+alternate+end,1,values),"Ambiguous OMML compatibility branch uses established fallback");
    std::puts("OMML_BODY_SNAPSHOT_PASSED");
}
inline void checkWordSources(const std::wstring& input) {
    // 只读指定合成输入，逐式对比生产批量解析与旧 XML 快照；不执行转换或保存。
    const auto original=sampleFileBytes(input);requireDocxZip(input);
    WordSession word;auto documents=get(word.application.Get(),L"Documents");
    word.document=asObject(invoke(documents.Get(),L"Add",DISPATCH_METHOD,{Value(),false,Value(),false}));word.verifyOwnership();
    invoke(word.application.Get(),L"Visible",DISPATCH_PROPERTYPUT,{false});invoke(word.application.Get(),L"DisplayAlerts",DISPATCH_PROPERTYPUT,{0L});
    invoke(word.application.Get(),L"AutomationSecurity",DISPATCH_PROPERTYPUT,{3L});invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();
    word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{input.c_str(),false,true,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));
    const auto readOnly=invoke(word.document.Get(),L"ReadOnly",DISPATCH_PROPERTYGET);require(readOnly.v.vt==VT_BOOL&&readOnly.v.boolVal,"Snapshot comparison opens source read-only");
    auto content=get(word.document.Get(),L"Content"),maths=get(content.Get(),L"OMaths");const long total=integer(maths.Get(),L"Count");require(total>3,"Snapshot comparison needs at least four equations");
    require(integer(get(content.Get(),L"InlineShapes").Get(),L"Count")==0&&integer(get(word.document.Get(),L"Shapes").Get(),L"Count")==0,"Snapshot comparison exercises supported production bulk path");
    std::vector<std::wstring> batch;require(ommlSourcesFromBodyXml(text(content.Get(),L"WordOpenXML"),static_cast<size_t>(total),batch),"Actual Word package supports body OMML snapshot extraction");
    for(long index=1;index<=total;++index){auto range=get(collectionItem(maths.Get(),index).Get(),L"Range");const auto actual=ommlSource(range.Get());
        std::printf("OMML_SNAPSHOT_COMPARE ID=%ld WORD_TEXT=%s XML_TEXT=%s EQUAL=%d\n",index,quote(text(range.Get(),L"Text")).c_str(),quote(contents(batch[static_cast<size_t>(index-1)])).c_str(),actual==batch[static_cast<size_t>(index-1)]);
        require(actual==batch[static_cast<size_t>(index-1)],"Bulk snapshot exactly matches every actual Word OMath source in order");}
    maths.Reset();content.Reset();documents.Reset();word.shutdown();require(sampleFileBytes(input)==original,"Read-only snapshot comparison preserves source bytes");
    taskCommit();std::printf("OMML_WORD_SNAPSHOT_PASSED COUNT=%ld\n",total);
}
inline void run(const std::wstring& stylesheetPath) {
    ConversionTimings timings;auto compiled=compileOmmlStylesheet(stylesheetPath,timings);
    auto oldStyle=xmlDocument(L"<root/>");VARIANT_BOOL loaded{};check(oldStyle->load(Value(stylesheetPath.c_str()).v,&loaded),"Load reference Office stylesheet");require(loaded!=VARIANT_FALSE,"Reference Office stylesheet exists");
    struct Case {const char* name;const wchar_t* math;const wchar_t* expectedNode;};
    const Case cases[]{
        {"ascii",L"<m:r><m:t>x+1=y</m:t></m:r>",L"//*[local-name()='mi']"},
        {"fraction",L"<m:f><m:num><m:r><m:t>x+1</m:t></m:r></m:num><m:den><m:r><m:t>y</m:t></m:r></m:den></m:f>",L"//*[local-name()='mfrac']"},
        {"matrix",L"<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m>",L"//*[local-name()='mtable']"},
        {"accent",L"<m:acc><m:accPr><m:chr m:val=\"&#x0302;\"/></m:accPr><m:e><m:r><m:t>Y</m:t></m:r></m:e></m:acc>",L"//*[local-name()='mover']"},
        {"entity",L"<m:r><m:t>&#x0059;&#x0302;</m:t></m:r>",L"//*[local-name()='mover']"}
    };
    for(const auto& item:cases)for(const bool display:{false,true}){
        std::printf("OMML_ADAPTER_CASE %s DISPLAY=%d\n",item.name,display);const auto input=source(item.math),original=input;
        const auto actual=ommlMathML(input,compiled.Get(),display,timings),expected=legacy(oldStyle.Get(),input,display,timings);
        require(actual==expected,"Compiled processor equals legacy transformNode at selected OMML node");
        require(input==original,"Adapter leaves raw recovery XML unchanged");
        require(actual.find(L"OUTSIDE_")==std::wstring::npos&&actual.find(L"UNSELECTED_FORMULA")==std::wstring::npos,"Adapter transforms selected equation only");
        require(count(actual,item.expectedNode)>0,"Adapter preserves expected mathematical structure");
        require(count(actual,display?L"/*[local-name()='math' and @display='block']":L"/*[local-name()='math' and @display='inline']")==1,"Adapter preserves display mode");
    }
    const auto clean=source(L"<m:r><m:t>x+1</m:t></m:r>"),dirty=source(L"<m:r><m:t xml:space=\"preserve\"> \tx\u00a0+1\n </m:t></m:r>"),original=dirty;
    require(ommlMathML(clean,compiled.Get(),false,timings)==ommlMathML(dirty,compiled.Get(),false,timings)&&dirty==original,"Adapter removes import whitespace without changing recovery XML");
    // 普通文本同样清理显式空白，避免厂商导入产生空白占位符；此为产品兼容约定。
    const auto normalClean=source(L"<m:r><m:rPr><m:nor/></m:rPr><m:t>forallx</m:t></m:r>"),normalSpaced=source(L"<m:r><m:rPr><m:nor/></m:rPr><m:t xml:space=\"preserve\">for all x</m:t></m:r>");
    require(ommlMathML(normalClean,compiled.Get(),false,timings)==ommlMathML(normalSpaced,compiled.Get(),false,timings),"Normal OMML text retains the intentional whitespace cleanup policy");
    const auto ascii=finishMathML(L"<mi>x</mi><mo>+</mo><mn>1</mn>",false,timings);
    require(contents(ascii)==L"x+1"&&count(ascii,L"//*[local-name()='mover']")==0,"ASCII path retains exact mathematical text");
    for(const auto* fragment:{L"<mi mathvariant=\"bold\">Y&#x0302;</mi>",L"<mi mathvariant=\"bold\">\u0176</mi>",L"<mi mathvariant=\"bold\">Y</mi><mo>&#x0302;</mo>"}){
        const auto adapted=finishMathML(fragment,false,timings);
        require(count(adapted,L"//*[local-name()='mover']")==1&&count(adapted,L"//*[local-name()='mi' and @mathvariant='bold' and text()='Y']")==1,"Decoded entity and Unicode accents retain token styling");
        require(contents(adapted)==L"Y^","Accent compatibility emits the established MathType hat glyph");
    }
    checkBodySources();
    require(timings.samples[static_cast<size_t>(ConversionMetric::XslCompile)].count==1,"Repeated formulas reuse one compiled stylesheet");
    std::puts("OMML_ADAPTER_SELF_TEST_PASSED");
}
}
