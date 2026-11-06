// MathType 编号协议：整个编号属于 MTPlaceRef，计数与章/节值均使用厂商序列。
#pragma once
#include "layout_options.h"
#include "docx_scan.h"
#include "numbering_rules.h"
inline Span nativeFieldSpan(IDispatch* field) {
    static DISPID codeId=DISPID_UNKNOWN;const auto span=bounds(get(field,L"Code",&codeId).Get());return {span.start-1,span.end+1};
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
struct NativeNumberUpdate {
    std::vector<ComPtr<IDispatch>> children;
    long visible{-1};
};
inline ComPtr<IDispatch> nativeNumber(IDispatch* document,long position,const EquationNumbering& settings,NativeNumberUpdate* update=nullptr) {
    ScopedConversionTiming timing(automationTimings,ConversionMetric::NumberCreate);
    if(update)*update={};
    auto field=nativeField(document,position,L"MACROBUTTON MTPlaceRef \\* MERGEFORMAT");
    auto increment=nativeChild(document,field.Get(),L"SEQ MTEqn \\h \\* MERGEFORMAT");
    if(update)update->children.push_back(increment);
    std::wstring literal;
    for(size_t i=0;i<settings.format.size();++i){
        if(settings.format[i]==L'#'&&i+2<settings.format.size()&&std::wstring(L"CSE").find(settings.format[i+1])!=std::wstring::npos){
            nativeLiteral(document,field.Get(),literal);literal.clear();const wchar_t component=settings.format[i+1],style=settings.format[i+2];
            const auto sequence=component==L'C'?L"MTChap":component==L'S'?L"MTSec":L"MTEqn";
            auto child=nativeChild(document,field.Get(),L"SEQ "+std::wstring(sequence)+L" \\c \\* "+numberSwitch(style)+L" \\* MERGEFORMAT");
            if(update){update->children.push_back(child);if(component==L'E')update->visible=static_cast<long>(update->children.size()-1);}i+=2;
        }else literal+=settings.format[i];
    }
    nativeLiteral(document,field.Get(),literal);return field;
}
inline ComPtr<IDispatch> cloneNativeNumber(IDispatch* document,long position,IDispatch* source,const NativeNumberUpdate& model,NativeNumberUpdate& update) {
    ScopedConversionTiming timing(automationTimings,ConversionMetric::NumberCreate);
    auto range=docRange(document,position,position);auto font=get(get(range.Get(),L"Font").Get(),L"Duplicate");
    // 复用完整域结构；字符格式仍继承目标段落，不沿用首个编号所在段落的字体。
    invoke(range.Get(),L"FormattedText",DISPATCH_PROPERTYPUT,{source});invoke(range.Get(),L"Font",DISPATCH_PROPERTYPUT,{font.Get()});
    auto field=collectionItem(get(range.Get(),L"Fields").Get(),1);auto children=get(get(field.Get(),L"Code").Get(),L"Fields");
    update.children=enumerateCollection(children.Get(),static_cast<long>(model.children.size()));update.visible=model.visible;return field;
}
inline bool updateNativeNumber(IDispatch* field,const NativeNumberUpdate* prepared=nullptr) {
    // 创建或规范化时已记录子域，更新时复用句柄，保留成功检查及显示结果核验。
    auto children=prepared?ComPtr<IDispatch>():get(get(field,L"Code").Get(),L"Fields");const long count=prepared?static_cast<long>(prepared->children.size()):integer(children.Get(),L"Count");bool visible{};
    static DISPID updateId=DISPID_UNKNOWN,resultId=DISPID_UNKNOWN;
    for(long i=0;i<count;++i){auto child=prepared?prepared->children[static_cast<size_t>(i)]:collectionItem(children.Get(),i+1);Value updated;
        {ScopedConversionTiming timing(automationTimings,ConversionMetric::NumberUpdate);updated=invoke(child.Get(),L"Update",DISPATCH_METHOD,{},&updateId);}
        require(updated.v.vt==VT_BOOL&&updated.v.boolVal!=VARIANT_FALSE,"Word updated the new native equation sequence without error");
        if(prepared?i==prepared->visible:nativeCurrentCounter(text(get(child.Get(),L"Code").Get(),L"Text"))){ScopedConversionTiming timing(automationTimings,ConversionMetric::NumberResult);visible=!text(get(child.Get(),L"Result",&resultId).Get(),L"Text").empty();}}
    return visible;
}
inline void normalizeNativeNumberCounter(IDispatch* field,NativeNumberUpdate* update=nullptr) {
    if(update)*update={};
    auto children=get(get(field,L"Code").Get(),L"Fields");const long count=integer(children.Get(),L"Count");bool found{};
    for(long i=1;i<=count;++i){auto child=collectionItem(children.Get(),i);auto code=get(child.Get(),L"Code");const auto value=text(code.Get(),L"Text");
        if(update){update->children.push_back(child);if(nativeCurrentCounter(value))update->visible=i-1;}
        if(nativeIncrementCounter(value)){
            const std::wstring next=L"SEQ MTEqn \\h \\* MERGEFORMAT";
            if(value!=next)invoke(code.Get(),L"Text",DISPATCH_PROPERTYPUT,{next.c_str()});found=true;
        }}
    require(found,"Native number contains its sequence counter");
}
