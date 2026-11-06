// 离线 MSXML 回归：调用正式适配器，不启动 Word 或 MathType，不写文档。
#pragma once
#include "omml_adapter.h"
#include "omml_tex.h"
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
inline void checkDirectTex() {
    // 成功路径核对结构和文字边界；不支持的公式必须整式回退，原 DOM 与输出参数保持。
    struct Case {const wchar_t* math;const wchar_t* expected;};
    const Case cases[]{
        {L"<m:sSubSup><m:e><m:f><m:num><m:r><m:t>\u03b1x</m:t></m:r></m:num><m:den><m:r><m:t>b+1</m:t></m:r></m:den></m:f></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>",L"{\\frac{\\alpha x}{b+1}}_{i}^{2}"},
        {L"<m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>j</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:rPr><m:nor/></m:rPr><m:t>a_b&amp;c#1%{z}</m:t></m:r>",L"\\mathrm{j}+\\text{a\\_b\\&c\\#1\\%\\{z\\}}"},
        {L"<m:rad><m:deg><m:r><m:t>i]j</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>",L"\\sqrt[{i]j}]{x}"},
        {L"<m:nary><m:naryPr><m:chr m:val='\u222b'/><m:grow m:val='0'/><m:subHide m:val='1'/><m:supHide m:val='1'/></m:naryPr><m:sub/><m:sup/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>",nullptr},
        {L"<m:r><m:t>x+</m:t></m:r><m:phant><m:e><m:r><m:t>z</m:t></m:r></m:e></m:phant>",nullptr},
        {L"<m:r><m:t>x=</m:t></m:r><m:r><m:rPr><m:sty m:val='bi'/></m:rPr><m:t>U</m:t></m:r>",nullptr},
        {L"<m:r><m:t>x=</m:t></m:r><m:r><m:rPr><m:scr m:val='script'/></m:rPr><m:t>l</m:t></m:r>",nullptr},
        {L"<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:argPr><m:argSz m:val='-1'/></m:argPr><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>",nullptr}
    };
    for(const auto& item:cases){
        auto dom=xmlDocument(source(item.math));ComPtr<IXMLDOMNode> equation;
        check(dom->selectSingleNode(Value(L"/*/*[local-name()='oMath'][1]").v.bstrVal,&equation),"Direct TeX test equation");
        Value before;before.v.vt=VT_BSTR;check(dom->get_xml(&before.v.bstrVal),"Direct TeX original DOM");
        std::wstring actual=L"UNCHANGED";const bool converted=tryOmmlTex(equation.Get(),actual);
        require(converted==(item.expected!=nullptr)&&actual==(item.expected?item.expected:L"UNCHANGED"),"Direct TeX structure, escaping, and whole-formula fallback");
        Value after;after.v.vt=VT_BSTR;check(dom->get_xml(&after.v.bstrVal),"Direct TeX unchanged DOM");
        require(std::wstring(before.v.bstrVal)==after.v.bstrVal,"Direct TeX compilation leaves the source DOM untouched");
    }
    std::puts("OMML_DIRECT_TEX_SELF_TEST_PASSED");
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
    // 输入缓存必须保留段落格式和手动编号，不能把纯公式快照拿来恢复整段。
    const auto paragraph=L"<w:p><w:pPr><w:spacing w:after=\"120\"/></w:pPr>"+equation(L"x")+L"<w:r><w:t> (27)</w:t></w:r></w:p>";
    require(ommlSourcesFromBodyXml(start+paragraph+end,1,values,true)&&values.size()==1&&!values[0].empty(),"Cache standalone OMML input package");
    require(contents(values[0])==L"x (27)"&&values[0].find(L"w:after=\"120\"")!=std::wstring::npos,"Cached recovery includes paragraph properties and manual number");
    const auto recovery=ommlRecoveryPackage(start+body+end,values[0]);
    require(contents(recovery)==L"x (27)HEADER"&&recovery.find(L"/word/header1.xml")!=std::wstring::npos,"Recovery replaces only main body and retains package dependencies");
    require(ommlSourcesFromBodyXml(start+L"<w:p>"+equation(L"x")+equation(L"y")+L"</w:p>"+end,2,values,true)&&values[0].empty()&&values[1].empty(),"Shared paragraph keeps range-specific input and recovery");
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
    // ASCII 公式同样展开运算分隔符；嵌套、空括号和末个分隔符重复不能改变含义。
    const auto asciiFences=finishMathML(L"<mfenced open=\"\" close=\"\" separators=\" + &#x09; - \"><mi>a</mi><mfenced separators=\"+\"><mi>b</mi><mi>c</mi></mfenced><mi>d</mi><mi>e</mi></mfenced>",false,timings);
    require(contents(asciiFences)==L"a+b+c-d-e"&&count(asciiFences,L"//*[local-name()='mfenced' and @separators='' and count(*)=1 and *[local-name()='mrow']]")==2,"ASCII and nested arithmetic separators become explicit operators");
    require(count(asciiFences,L"/*[local-name()='math']/*[local-name()='mfenced' and @open='' and @close='']/*[local-name()='mrow' and count(*)=7 and *[2][local-name()='mo' and text()='+'] and *[4][local-name()='mo' and text()='-'] and *[6][local-name()='mo' and text()='-']]")==1,"Empty fences and separator sequence whitespace and repetition are preserved");
    // 范数只在属性中出现时仍需归一化，正文中的同一字符不能被全局改写。
    const auto norm=finishMathML(L"<mfenced open='&#x2016;' close='&#x2016;'><mi mathvariant='bold-italic'>x</mi></mfenced><mo>&#x2016;</mo>",false,timings);
    require(count(norm,L"//*[local-name()='mfenced' and @open='||' and @close='||']/*[local-name()='mi' and @mathvariant='bold-italic' and text()='x']")==1&&contents(norm)==L"x\u2016","Norm delimiters normalize without changing mathematical text or bold italic style");
    const auto normAscii=finishMathML(L"<mfenced open='&#x2016;' close='&#x2016;'><mi>x</mi></mfenced>",false,timings);
    require(count(normAscii,L"//*[local-name()='mfenced' and @open='||' and @close='||']")==1,"ASCII-only content still normalizes norm delimiter attributes");
    for(const auto* separator:{L"&#x2264;",L"&#x2265;",L"&lt;",L"&gt;",L"/"}){
        const auto relation=finishMathML(L"<mfenced open='&#x2016;' close='&#x2016;' separators='"+std::wstring(separator)+L"'><mi mathvariant='bold-italic'>x</mi><mi>y</mi></mfenced>",false,timings);
        const auto decoded=contents(L"<root>"+std::wstring(separator)+L"</root>");
        require(contents(relation)==L"x"+decoded+L"y"&&count(relation,L"//*[local-name()='mfenced' and @open='||' and @close='||' and @separators='']/*[local-name()='mrow' and count(*)=3 and *[2][local-name()='mo']]")==1&&
            count(relation,L"//*[local-name()='mi' and @mathvariant='bold-italic' and text()='x']")==1,"Relational and slash separators become explicit operators while preserving norm fences and styling");
    }
    // 指数函数文本及上下标作为完整子树移动，不能拆散基底、下标和上标。
    const auto scriptedFence=finishMathML(L"<mtext>exp</mtext><mfenced open=\"(\" close=\")\" separators=\"&#x2212;\"><msub><mi>x</mi><mi>i</mi></msub><msup><mi>y</mi><mn>2</mn></msup><msubsup><mi mathvariant=\"bold\">z</mi><mi>j</mi><mo>&#x22c6;</mo></msubsup></mfenced>",false,timings);
    require(contents(scriptedFence)==L"expxi-y2-zj\u22c6"&&count(scriptedFence,L"/*[local-name()='math']/*[local-name()='mtext' and text()='exp']")==1&&count(scriptedFence,L"//*[local-name()='mfenced' and @open='(' and @close=')' and @separators='']/*[local-name()='mrow']/*[local-name()='mo' and text()='-']")==2,"Unicode separator minus becomes explicit ASCII minus without changing exp text");
    require(count(scriptedFence,L"//*[local-name()='mrow']/*[local-name()='msub' and count(*)=2 and *[1][local-name()='mi' and text()='x'] and *[2][local-name()='mi' and text()='i']]")==1&&
        count(scriptedFence,L"//*[local-name()='mrow']/*[local-name()='msup' and count(*)=2 and *[1][local-name()='mi' and text()='y'] and *[2][local-name()='mn' and text()='2']]")==1&&
        count(scriptedFence,L"//*[local-name()='mrow']/*[local-name()='msubsup' and count(*)=3 and *[1][local-name()='mi' and text()='z' and @mathvariant='bold'] and *[2][local-name()='mi' and text()='j'] and *[3][local-name()='mo' and text()='\u22c6']]")==1,"Subscripts superscripts and conjugate star remain attached to their original bases");
    // 普通分隔符继续交给原路径，避免影响默认逗号、竖线或显式无分隔符。
    const auto ordinaryFences=finishMathML(L"<mfenced><mi>a</mi><mi>b</mi></mfenced><mfenced separators=\",\"><mi>a</mi><mi>b</mi></mfenced><mfenced separators=\"|\"><mi>a</mi><mi>b</mi></mfenced><mfenced separators=\"\"><mi>a</mi><mi>b</mi></mfenced>",false,timings);
    require(contents(ordinaryFences)==L"abababab"&&count(ordinaryFences,L"//*[local-name()='mfenced' and count(*)=2 and count(*[local-name()='mi'])=2]")==4&&
        count(ordinaryFences,L"//*[local-name()='mfenced' and not(@separators)]")==1&&count(ordinaryFences,L"//*[local-name()='mfenced' and @separators=',']")==1&&
        count(ordinaryFences,L"//*[local-name()='mfenced' and @separators='|']")==1&&count(ordinaryFences,L"//*[local-name()='mfenced' and @separators='']")==1,"Default comma vertical bar and empty separators remain unchanged");
    for(const auto* fragment:{L"<mi mathvariant=\"bold\">Y&#x0302;</mi>",L"<mi mathvariant=\"bold\">\u0176</mi>",L"<mi mathvariant=\"bold\">Y</mi><mo>&#x0302;</mo>"}){
        const auto adapted=finishMathML(fragment,false,timings);
        require(count(adapted,L"//*[local-name()='mover']")==1&&count(adapted,L"//*[local-name()='mi' and @mathvariant='bold' and text()='Y']")==1,"Decoded entity and Unicode accents retain token styling");
        require(contents(adapted)==L"Y^","Accent compatibility emits the established MathType hat glyph");
    }
    // 极速简化只删除无属性的单参数分组，保留上下标参数、复合分组及样式。
    const auto grouped=L"<mrow><mrow><msubsup><mrow><mi>x</mi></mrow><mrow><mi>i</mi><mo>+</mo><mn>1</mn></mrow><mrow mathvariant=\"bold\"><mn>2</mn></mrow></msubsup></mrow></mrow><mrow><mover><mrow><mi>Y</mi></mrow><mo>^</mo></mover></mrow>";
    const auto unchanged=finishMathML(grouped,false,timings),compact=finishMathML(grouped,false,timings,true);
    require(count(unchanged,L"//*[local-name()='mrow']")==7&&count(compact,L"//*[local-name()='mrow']")==2&&contents(compact)==contents(unchanged),"Only explicit ultra compaction removes redundant rows without changing tokens");
    require(count(compact,L"//*[local-name()='msubsup' and count(*)=3 and *[1][local-name()='mi' and text()='x'] and *[2][local-name()='mrow' and count(*)=3] and *[3][local-name()='mrow' and @mathvariant='bold']]")==1&&
        count(compact,L"//*[local-name()='mover' and count(*)=2 and *[1][local-name()='mi' and text()='Y'] and *[2][local-name()='mo' and text()='^']]")==1,"Compaction preserves script arity, styled rows, and accent operators");
    checkDirectTex();checkBodySources();
    require(timings.samples[static_cast<size_t>(ConversionMetric::XslCompile)].count==1,"Repeated formulas reuse one compiled stylesheet");
    std::puts("OMML_ADAPTER_SELF_TEST_PASSED");
}
}
