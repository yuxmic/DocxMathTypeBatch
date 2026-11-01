// MathType 编号协议：整个编号属于 MTPlaceRef，计数与章/节值均使用厂商序列。
#pragma once
#include "layout_options.h"
#include "docx_scan.h"
#include "numbering_rules.h"
inline Span nativeFieldSpan(IDispatch* field) {
    auto code=get(field,L"Code");return {integer(code.Get(),L"Start")-1,integer(code.Get(),L"End")+1};
}
inline ComPtr<IDispatch> nativeField(IDispatch* document,long position,const std::wstring& code) {
    auto range=docRange(document,position,position);
    return asObject(invoke(get(range.Get(),L"Fields").Get(),L"Add",DISPATCH_METHOD,{range.Get(),-1L,code.c_str(),false}));
}
inline void nativeLiteral(IDispatch* document,IDispatch* outer,const std::wstring& value) {
    if(value.empty())return;const long position=integer(get(outer,L"Code").Get(),L"End");
    invoke(docRange(document,position,position).Get(),L"InsertBefore",DISPATCH_METHOD,{value.c_str()});
}
inline ComPtr<IDispatch> nativeChild(IDispatch* document,IDispatch* outer,const std::wstring& code) {
    return nativeField(document,integer(get(outer,L"Code").Get(),L"End"),code);
}
inline void nativeProperty(IDispatch* document,const wchar_t* name,const std::wstring& value) {
    auto properties=get(document,L"CustomDocumentProperties");
    for(long i=1;i<=integer(properties.Get(),L"Count");++i){auto property=asObject(invoke(properties.Get(),L"Item",DISPATCH_PROPERTYGET,{i}));
        if(text(property.Get(),L"Name")==name){invoke(property.Get(),L"Value",DISPATCH_PROPERTYPUT,{value.c_str()});return;}}
    invoke(properties.Get(),L"Add",DISPATCH_METHOD,{name,false,4L,value.c_str()});
}
inline ComPtr<IDispatch> nativeNumber(IDispatch* document,long position,const EquationNumbering& settings) {
    auto field=nativeField(document,position,L"MACROBUTTON MTPlaceRef \\* MERGEFORMAT");
    nativeChild(document,field.Get(),L"SEQ MTEqn \\h \\* MERGEFORMAT");
    std::wstring literal;
    for(size_t i=0;i<settings.format.size();++i){
        if(settings.format[i]==L'#'&&i+2<settings.format.size()&&std::wstring(L"CSE").find(settings.format[i+1])!=std::wstring::npos){
            nativeLiteral(document,field.Get(),literal);literal.clear();const wchar_t component=settings.format[i+1],style=settings.format[i+2];
            const auto sequence=component==L'C'?L"MTChap":component==L'S'?L"MTSec":L"MTEqn";
            nativeChild(document,field.Get(),L"SEQ "+std::wstring(sequence)+L" \\c \\* "+numberSwitch(style)+L" \\* MERGEFORMAT");i+=2;
        }else literal+=settings.format[i];
    }
    nativeLiteral(document,field.Get(),literal);return field;
}
inline bool updateNativeNumber(IDispatch* field) {
    auto children=get(get(field,L"Code").Get(),L"Fields");bool visible{};
    for(long i=1;i<=integer(children.Get(),L"Count");++i){auto child=collectionItem(children.Get(),i);const auto updated=invoke(child.Get(),L"Update",DISPATCH_METHOD);
        require(updated.v.vt==VT_BOOL&&updated.v.boolVal!=VARIANT_FALSE,"Word updated the new native equation sequence without error");
        if(nativeCurrentCounter(text(get(child.Get(),L"Code").Get(),L"Text")))visible=!text(get(child.Get(),L"Result").Get(),L"Text").empty();}
    return visible;
}
inline void normalizeNativeNumberCounter(IDispatch* field) {
    auto children=get(get(field,L"Code").Get(),L"Fields");const long count=integer(children.Get(),L"Count");bool found{};
    for(long i=1;i<=count;++i){auto code=get(collectionItem(children.Get(),i).Get(),L"Code");const auto value=text(code.Get(),L"Text");
        if(nativeIncrementCounter(value)){
            const std::wstring next=L"SEQ MTEqn \\h \\* MERGEFORMAT";
            if(value!=next)invoke(code.Get(),L"Text",DISPATCH_PROPERTYPUT,{next.c_str()});found=true;
        }}
    require(found,"Native number contains its sequence counter");
}
