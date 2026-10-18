// 开发回归：调用已安装 MathType 的真实更新/格式化过程，不进入成品 GUI。
#pragma once
#include "native_numbering.h"
inline Value nativeMacroReference(Value& original){Value value;value.v.vt=VT_VARIANT|VT_BYREF;value.v.pvarVal=&original.v;return value;}
inline void nativeMacro(IDispatch* application,const wchar_t* name,Value first,Value second=Value()){
    // Word 的全局模板带参数过程使用模块名；附项目名前缀在本机返回 438。
    auto a=nativeMacroReference(first),b=nativeMacroReference(second);
    invoke(application,L"Run",DISPATCH_METHOD,{name,a,b,Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value(),Value()});
}
inline std::vector<ComPtr<IDispatch>> nativeNumbers(IDispatch* document){
    std::vector<ComPtr<IDispatch>> result;auto fields=get(get(document,L"Content").Get(),L"Fields");
    for(long i=1;i<=integer(fields.Get(),L"Count");++i){auto field=collectionItem(fields.Get(),i);if(integer(field.Get(),L"Type")==51&&text(get(field.Get(),L"Code").Get(),L"Text").find(L"MACROBUTTON MTPlaceRef")!=std::wstring::npos)result.push_back(field);}
    return result;
}
inline std::vector<std::wstring> numberingAssets(IDispatch* document){
    const auto xml=text(get(document,L"Content").Get(),L"WordOpenXML");std::vector<std::wstring> result;size_t position{};
    const std::wstring begin=L"<pkg:binaryData>",end=L"</pkg:binaryData>";
    while((position=xml.find(begin,position))!=std::wstring::npos){position+=begin.size();const auto last=xml.find(end,position);require(last!=std::wstring::npos,"Complete test binary part");result.push_back(xml.substr(position,last-position));position=last+end.size();}return result;
}
inline void checkNativeNumbering(const std::wstring& input,const std::wstring& directory,bool fixture,bool displayOnly=false){
    const auto original=sampleFileBytes(input);const auto copy=directory+(displayOnly?L"\\manual-number.docx":fixture?L"\\native-fixture.docx":L"\\native-checked.docx");
    const auto working=copy+L".work.docx";taskCopy(input,working);taskTrackWordOwner(working);
    std::wstring problem;const auto install=findMathTypeInstall(problem);require(problem.empty(),"Installed MathType required");
    WordSession word;auto docs=get(word.application.Get(),L"Documents");word.document=asObject(invoke(docs.Get(),L"Open",DISPATCH_METHOD,{working.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));word.verifyOwnership();auto doc=word.document.Get();
    invoke(word.application.Get(),L"AutomationSecurity",DISPATCH_PROPERTYPUT,{1L});
    invoke(get(word.application.Get(),L"AddIns").Get(),L"Add",DISPATCH_METHOD,{(install+L"\\Office Support\\64\\MathType Commands 2016.dotm").c_str(),true});
    invoke(word.application.Get(),L"Run",DISPATCH_METHOD,{L"MathTypeCommands.MTLibCommon.SetDLLPath"});
    invoke(word.application.Get(),L"AutomationSecurity",DISPATCH_PROPERTYPUT,{3L});invoke(doc,L"Activate",DISPATCH_METHOD);
    auto select=[&](long position){invoke(docRange(doc,position,position).Get(),L"Select",DISPATCH_METHOD);};
    if(displayOnly){
        // 在未编号输出上调用用户标签页的同一个按钮，核对确实按右制表位插入，不预改文档编号设置。
        const auto assets=numberingAssets(doc);require(nativeNumbers(doc).empty(),"Manual numbering starts without numbers");
        auto shapes=get(get(doc,L"Content").Get(),L"InlineShapes");bool inserted{};
        for(long i=1;i<=integer(shapes.Get(),L"Count");++i){auto range=get(collectionItem(shapes.Get(),i).Get(),L"Range");if(text(get(range.Get(),L"Style").Get(),L"NameLocal")!=L"MTDisplayEquation")continue;
            auto paragraph=collectionItem(get(range.Get(),L"Paragraphs").Get(),1);select(integer(get(paragraph.Get(),L"Range").Get(),L"End")-1);invoke(word.application.Get(),L"Run",DISPATCH_METHOD,{L"MathTypeCommands.MTEqnNum.InsertEqnNum"});inserted=true;break;}
        require(inserted&&nativeNumbers(doc).size()==1,"Vendor button inserts one number into native display paragraph");
        require(numberingAssets(doc)==assets,"Manual number leaves equation data unchanged");
        invoke(doc,L"ExportAsFixedFormat",DISPATCH_METHOD,{(copy+L".pdf").c_str(),17L});invoke(doc,L"Save",DISPATCH_METHOD);word.shutdown();
        require(sampleFileBytes(input)==original,"Manual button test preserves source");taskMove(working,copy);taskCommit();std::puts("NATIVE_DISPLAY_MANUAL_NUMBER_PASSED");return;
    }
    if(fixture){
        nativeProperty(doc,L"MTEquationNumber2",L"(#E1)");auto range=docRange(doc,0,0);invoke(range.Get(),L"InsertBefore",DISPATCH_METHOD,{L"已有原生编号：\r"});select(7);
        // 不预置已检查标记，让原生命令静默插入自己的起始章节断点。
        invoke(word.application.Get(),L"Run",DISPATCH_METHOD,{L"MathTypeCommands.MTEqnNum.InsertEqnNum"});
        require(nativeNumbers(doc).size()==1,"Vendor fixture contains one native number");
        // 已有章、节不是 1 时，默认转换也应接续原生断点，不能重置回 1。
        auto fields=get(get(doc,L"Content").Get(),L"Fields");
        for(long i=1;i<=integer(fields.Get(),L"Count");++i){auto field=collectionItem(fields.Get(),i);if(integer(field.Get(),L"Type")!=12)continue;auto code=get(field.Get(),L"Code");const auto instruction=text(code.Get(),L"Text");
            if(instruction.find(L"SEQ MTChap \\r")!=std::wstring::npos)invoke(code.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"SEQ MTChap \\r 5 \\h \\* MERGEFORMAT"});
            if(instruction.find(L"SEQ MTSec \\r")!=std::wstring::npos)invoke(code.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"SEQ MTSec \\r 7 \\h \\* MERGEFORMAT"});}
        nativeMacro(word.application.Get(),L"MTUpdateFields.UpdateFields",0L,true);
        invoke(doc,L"Save",DISPATCH_METHOD);word.shutdown();require(sampleFileBytes(input)==original,"Native fixture preserves source");taskMove(working,copy);taskCommit();std::puts("NATIVE_NUMBER_FIXTURE_PASSED");return;
    }
    const auto assets=numberingAssets(doc);require(!assets.empty(),"Native test includes equation assets");const auto initial=nativeNumbers(doc).size();require(initial>=5,"Native test needs at least five numbered equations");
    auto sectionCount=[&](){long count{};auto fields=get(get(doc,L"Content").Get(),L"Fields");for(long i=1;i<=integer(fields.Get(),L"Count");++i){auto field=collectionItem(fields.Get(),i);if(integer(field.Get(),L"Type")==51&&text(get(field.Get(),L"Code").Get(),L"Text").find(L"MACROBUTTON MTEditEquationSection")!=std::wstring::npos)++count;}return count;};
    const long sections=sectionCount();
    // 开启编辑标记后仍要求不产生新断点，不依赖隐藏文字的显示设置。
    invoke(get(get(word.application.Get(),L"ActiveWindow").Get(),L"View").Get(),L"ShowAll",DISPATCH_PROPERTYPUT,{true});
    auto exportPdf=[&](const wchar_t* name){require(sectionCount()==sections,"Native commands do not insert new equation section breaks");require(numberingAssets(doc)==assets,"Native commands preserve complete equation data and order");invoke(doc,L"ExportAsFixedFormat",DISPATCH_METHOD,{(directory+L"\\"+name+L".pdf").c_str(),17L});};
    exportPdf(L"original");
    const auto firstCode=text(get(nativeNumbers(doc).front().Get(),L"Code").Get(),L"Text");
    if(firstCode.find(L"\\r")==std::wstring::npos&&firstCode.find(L"SEQ MTChap")==std::wstring::npos&&firstCode.find(L"SEQ MTSec")==std::wstring::npos){
        // 默认起始值不能固定在旧首式上，否则用户在前面插入编号会产生重复的 1。
        select(0);invoke(word.application.Get(),L"Run",DISPATCH_METHOD,{L"MathTypeCommands.MTEqnNum.InsertEqnNum"});nativeMacro(word.application.Get(),L"MTUpdateFields.UpdateFields",0L,true);
        require(nativeNumbers(doc).size()==initial+1,"Vendor prepends a number and shifts existing sequence");exportPdf(L"prepended");
        auto inserted=nativeNumbers(doc).front();const auto span=nativeFieldSpan(inserted.Get());invoke(docRange(doc,span.start,span.end).Get(),L"Delete",DISPATCH_METHOD);nativeMacro(word.application.Get(),L"MTUpdateFields.UpdateFields",0L,true);
    }
    auto number=nativeNumbers(doc).at(1);const auto span=nativeFieldSpan(number.Get());invoke(docRange(doc,span.start,span.end).Get(),L"Delete",DISPATCH_METHOD);select(0);
    nativeMacro(word.application.Get(),L"MTUpdateFields.UpdateFields",0L,true);require(nativeNumbers(doc).size()==initial-1,"Native whole-document update after deletion");exportPdf(L"deleted-updated");
    nativeProperty(doc,L"MTEquationNumber2",L"[#E1]");nativeMacro(word.application.Get(),L"MTEqnNumFormat.UpdateEquationNumbers",0L,L"[#E1]");exportPdf(L"formatted");
    // 用标签页同一插入过程追加编号，再格式化全文，验证新旧编号共同参与序列。
    select(integer(get(doc,L"Content").Get(),L"End")-1);invoke(word.application.Get(),L"Run",DISPATCH_METHOD,{L"MathTypeCommands.MTEqnNum.InsertEqnNum"});
    require(nativeNumbers(doc).size()==initial,"Vendor appends next number without setup dialog");exportPdf(L"appended");
    nativeProperty(doc,L"MTEquationNumber2",L"(#E1)");nativeMacro(word.application.Get(),L"MTEqnNumFormat.UpdateEquationNumbers",0L,L"(#E1)");exportPdf(L"reformatted");
    invoke(doc,L"Save",DISPATCH_METHOD);invoke(doc,L"Close",DISPATCH_METHOD,{0L});word.document.Reset();
    word.document=asObject(invoke(docs.Get(),L"Open",DISPATCH_METHOD,{working.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),Value(),false}));doc=word.document.Get();invoke(doc,L"Activate",DISPATCH_METHOD);select(0);
    nativeMacro(word.application.Get(),L"MTUpdateFields.UpdateFields",0L,true);require(nativeNumbers(doc).size()==initial,"Native numbering survives save/reopen");exportPdf(L"reopened");
    invoke(doc,L"Save",DISPATCH_METHOD);word.shutdown();require(sampleFileBytes(input)==original,"Native interoperability test preserves source bytes");taskMove(working,copy);taskCommit();std::printf("NATIVE_NUMBER_INTEROPERABILITY_PASSED initial=%zu\n",initial);
}
