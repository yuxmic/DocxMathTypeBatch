// 本次编号的域更新、范围跟踪和计数；关闭文档前按域顺序记录核验豁免项。
#pragma once
#include "body_conversion.h"
#include "native_numbering.h"
#include "equation_objects.h"
#include <map>

struct NativeNumberingSession {
    EquationNumbering settings;
    NumberingSummary summary;
    bool reportProgress{};
    std::vector<ComPtr<IDispatch>> tracked;
    std::vector<long> saved;
    std::vector<long> indexed;
    NativeNumberingSession(const EquationNumbering& value,bool progress):settings(value),reportProgress(progress){}
    void reset(){summary={};tracked.clear();saved.clear();indexed.clear();}
    void remember(IDispatch* document,Span span){tracked.push_back(docRange(document,span.start,span.end));}
    std::vector<Span> ranges() const {
        std::vector<Span> result;for(const auto& range:tracked)result.push_back(bounds(range.Get()));return result;
    }
    void freeze(){
        if(tracked.empty())return;
        // 调用前的最终域核验已收集序号；保存只规范化代码长度，不再重新遍历全文域。
        saved=indexed;
        tracked.clear();
    }
    void includeReferences(IDispatch* document,ScanResult& scan){
        if(!std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.existingDisplay;}))return;
        ComPtr<IDispatch> bookmarks;std::vector<ScanItem> references;
        for(const auto& field:scan.fieldSnapshots){if(field.kind!=3)continue;
            const auto first=field.code.find_first_not_of(L" \t\r\n");if(first==std::wstring::npos||field.code.compare(first,4,L"REF "))continue;
            const auto start=field.code.find_first_not_of(L" \t\r\n",first+4);if(start==std::wstring::npos)continue;const auto end=field.code.find_first_of(L" \t\r\n",start);
            auto name=field.code.substr(start,end==std::wstring::npos?end:end-start);if(name.size()>1&&name.front()==L'"'&&name.back()==L'"')name=name.substr(1,name.size()-2);
            if(!bookmarks)bookmarks=get(document,L"Bookmarks");const auto exists=invoke(bookmarks.Get(),L"Exists",DISPATCH_METHOD,{name.c_str()});if(exists.v.vt!=VT_BOOL||exists.v.boolVal==VARIANT_FALSE)continue;
            const auto mark=asObject(invoke(bookmarks.Get(),L"Item",DISPATCH_METHOD,{name.c_str()}));const auto span=bounds(get(mark.Get(),L"Range").Get());
            const bool target=std::any_of(scan.items.begin(),scan.items.end(),[&](const ScanItem& item){return item.existingDisplay&&span.start>=item.span.start&&span.end<=item.span.end;});
            const bool number=std::any_of(scan.fieldSnapshots.begin(),scan.fieldSnapshots.end(),[&](const ScannedField& value){return nativeNumberField(value.kind,value.code)&&span.start>=value.span.start&&span.end<=value.span.end;});
            if(target&&number){ScanItem item;item.span=field.span;item.source=text(docRange(document,field.span.start,field.span.end).Get(),L"Text");item.existingReference=true;item.delimiter=L"EXISTING_REFERENCE";references.push_back(std::move(item));}
        }
        scan.items.insert(scan.items.end(),references.begin(),references.end());std::sort(scan.items.begin(),scan.items.end(),[](const ScanItem& a,const ScanItem& b){return a.span.start<b.span.start;});
        for(size_t i=1;i<scan.items.size();++i)require(scan.items[i-1].span.end<=scan.items[i].span.start,"Number reference changes do not overlap conversion targets");
    }
    std::vector<std::wstring> fields(IDispatch* document,const std::vector<ScannedField>* cached=nullptr) {
        std::vector<std::wstring> snapshot;const auto owned=ranges();
        indexed.clear();
        auto fields=cached?ComPtr<IDispatch>():get(get(document,L"Content").Get(),L"Fields");
        const long count=cached?static_cast<long>(cached->size()):integer(fields.Get(),L"Count");
        for(long index=0;index<count;++index){
            ScannedField read;const auto& value=cached?cached->at(static_cast<size_t>(index)):read;ComPtr<IDispatch> field,code;
            if(!cached){field=collectionItem(fields.Get(),index+1);read.kind=integer(field.Get(),L"Type");}
            if(value.kind==58)continue;
            if(!cached){code=get(field.Get(),L"Code");if(!owned.empty())read.codeSpan=bounds(code.Get());}
            if(std::binary_search(saved.begin(),saved.end(),index)||std::any_of(owned.begin(),owned.end(),[&](const Span& range){return range.start<=value.codeSpan.start&&range.end>=value.codeSpan.end;})){indexed.push_back(index);continue;}
            if(!cached){read.code=text(code.Get(),L"Text");read.result=text(get(field.Get(),L"Result").Get(),L"Text");}
            snapshot.push_back(std::to_wstring(value.kind));snapshot.push_back(value.code);snapshot.push_back(value.result);
        }return snapshot;
    }
    // 创建兼容字段而不加载交互宏；所有外层括号和内层序列都归 MathType 管理。
    void apply(IDispatch* document,const ScanResult& scan,std::vector<ItemOutcome>& outcomes) {
        std::vector<ComPtr<IDispatch>> native(outcomes.size());size_t first=outcomes.size(),count{};
        for(size_t i=0;i<outcomes.size();++i)if(scan.items[i].display&&displayOutcome(outcomes[i])){first=std::min(first,i);++count;}
        if(!count)return;
        const long originalFields=integer(get(get(document,L"Content").Get(),L"Fields").Get(),L"Count");
        nativeProperty(document,L"MTEquationNumber2",settings.format);nativeProperty(document,L"MTCustomEquationNumber",settings.custom?L"1":L"0");nativeProperty(document,L"MTEquationSection",L"1");
        auto remember=[&](IDispatch* field){this->remember(document,nativeFieldSpan(field));};
        StageProgress progress("NUMBER_EQUATIONS",count,reportProgress);
        for(size_t i=outcomes.size();i-- > 0;){
            if(!scan.items[i].display||!displayOutcome(outcomes[i]))continue;
            auto& outcome=outcomes[i];auto range=docRange(document,outcome.finalSpan.start,outcome.finalSpan.end);
            auto shape=collectionItem(get(range.Get(),L"InlineShapes").Get(),1);
            const long beforeEnd=integer(get(document,L"Content").Get(),L"End");
            std::map<std::wstring,ComPtr<IDispatch>> seeds;
            if(scan.items[i].existingDisplay){
                auto paragraph=collectionItem(get(range.Get(),L"Paragraphs").Get(),1);auto fields=get(get(paragraph.Get(),L"Range").Get(),L"Fields");const long fieldCount=integer(fields.Get(),L"Count");
                for(long n=1;n<=fieldCount;++n){auto field=collectionItem(fields.Get(),n);const long kind=integer(field.Get(),L"Type");if(kind!=51&&kind!=12)continue;
                    const auto code=text(get(field.Get(),L"Code").Get(),L"Text");
                    if(nativeNumberField(kind,code)){require(!native[i],"One native number per display paragraph");native[i]=field;}
                    if(nativeResetSeed(kind,code)){const auto span=nativeFieldSpan(field.Get());if(native[i]){const auto outer=nativeFieldSpan(native[i].Get());if(span.start>=outer.start&&span.end<=outer.end)continue;}for(const auto& rule:NativeSequenceRules)if(nativeSequenceField(kind,code,rule.name))seeds[std::wstring(rule.name)]=field;}
                }
                // 旧起始值域不再重置中途序列；首个目标的隐藏种子域负责显式起始值。
                const auto seed=seeds.find(L"MTEqn");if(seed!=seeds.end()){auto code=get(seed->second.Get(),L"Code");invoke(code.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"SEQ MTEqn \\c \\h"});invoke(seed->second.Get(),L"Update",DISPATCH_METHOD);}
            }
            const long end=integer(get(shape.Get(),L"Range").Get(),L"End");
            if(!native[i]){native[i]=nativeNumber(document,end+1,settings);++summary.inserted;}
            else{auto gap=docRange(document,end,nativeFieldSpan(native[i].Get()).start);if(seeds.empty()&&text(gap.Get(),L"Text")!=L"\t")invoke(gap.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"\t"});++summary.updated;}
            remember(native[i].Get());
            if(i==first){
                // 只为显式高级设置添加空结果序列域，不创建 MathType 章节断点或可见说明文字。
                std::vector<std::pair<std::wstring,long>> counters;
                // 厂商格式化会重建编号内部域；显式起始值放在外部空结果域，下一编号自然递增。
                if(settings.start!=1)counters.emplace_back(L"MTEqn",settings.start-1);
                if(settings.format.find(L"#C")!=std::wstring::npos)counters.emplace_back(L"MTChap",settings.chapter);
                if(settings.format.find(L"#S")!=std::wstring::npos)counters.emplace_back(L"MTSec",settings.section);
                for(const auto& counter:counters){
                    auto seed=seeds.count(counter.first)?seeds.at(counter.first):nativeField(document,nativeFieldSpan(native[i].Get()).start,L"SEQ "+counter.first+L" \\r "+std::to_wstring(counter.second)+L" \\h");
                    const auto instruction=L"SEQ "+counter.first+L" \\r "+std::to_wstring(counter.second)+L" \\h";invoke(get(seed.Get(),L"Code").Get(),L"Text",DISPATCH_PROPERTYPUT,{instruction.c_str()});
                    const auto updated=invoke(seed.Get(),L"Update",DISPATCH_METHOD);require(updated.v.vt==VT_BOOL&&updated.v.boolVal!=VARIANT_FALSE,"New numbering value initialized without error");
                    require(text(get(seed.Get(),L"Result").Get(),L"Text").empty(),"Numbering seed has no displayed result");
                    remember(seed.Get());
                }
            }
            outcome.finalSpan=bounds(get(shape.Get(),L"Range").Get());
            const long delta=integer(get(document,L"Content").Get(),L"End")-beforeEnd;
            outcome.replacementSpan.end+=delta;
            for(size_t later=i+1;later<outcomes.size();++later){outcomes[later].finalSpan.start+=delta;outcomes[later].finalSpan.end+=delta;outcomes[later].replacementSpan.start+=delta;outcomes[later].replacementSpan.end+=delta;}
            require(!conversionFault(L"DOCX_MATH_BODY_FAULT",L"*",L"NUMBERING"),"Injected numbering update failure");
            taskTestBoundary(L"NUMBERING_WRITTEN");
            progress.advance();
        }
        // 按文档顺序只更新本次新增序列，保持原有普通域的代码与缓存结果。
        for(size_t i=0;i<native.size();++i)if(native[i]){
            const long before=integer(get(document,L"Content").Get(),L"End"),position=nativeFieldSpan(native[i].Get()).start;
            normalizeNativeNumberCounter(native[i].Get());
            require(updateNativeNumber(native[i].Get()),"New MathType native sequence has a visible value");
            const long delta=integer(get(document,L"Content").Get(),L"End")-before;
            if(delta)for(auto& outcome:outcomes){if(outcome.finalSpan.start>position){outcome.finalSpan.start+=delta;outcome.finalSpan.end+=delta;}if(outcome.replacementSpan.start>position){outcome.replacementSpan.start+=delta;outcome.replacementSpan.end+=delta;}else if(outcome.replacementSpan.end>=position)outcome.replacementSpan.end+=delta;}
        }
        for(size_t i=0;i<outcomes.size();++i)if(scan.items[i].existingReference){
            auto& outcome=outcomes[i];auto range=docRange(document,outcome.replacementSpan.start,outcome.replacementSpan.end);auto field=collectionItem(get(range.Get(),L"Fields").Get(),1);
            const long before=integer(get(document,L"Content").Get(),L"End");
            const auto updated=invoke(field.Get(),L"Update",DISPATCH_METHOD);require(updated.v.vt==VT_BOOL&&updated.v.boolVal!=VARIANT_FALSE,"MathType number reference updated without error");
            const long delta=integer(get(document,L"Content").Get(),L"End")-before;outcome.replacementSpan.end+=delta;outcome.status=L"REFERENCE_UPDATED";++summary.referencesUpdated;
            for(size_t later=i+1;later<outcomes.size();++later){outcomes[later].finalSpan.start+=delta;outcomes[later].finalSpan.end+=delta;outcomes[later].replacementSpan.start+=delta;outcomes[later].replacementSpan.end+=delta;}
        }
        for(size_t i=0;i<outcomes.size();++i)if(native[i]||scan.items[i].existingReference)outcomes[i].outputText=text(docRange(document,outcomes[i].replacementSpan.start,outcomes[i].replacementSpan.end).Get(),L"Text");
        summary.addedFields=integer(get(get(document,L"Content").Get(),L"Fields").Get(),L"Count")-originalFields;
    }
};
