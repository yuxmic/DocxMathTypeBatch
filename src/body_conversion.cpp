// 正文转换阶段拥有坐标、编号和恢复状态；文件发布由任务入口负责。
#include "body_conversion.h"
#include "body_numbering.h"
#include "equation_styles.h"
#include "omml_adapter.h"
#include "scan_file.h"
#include "conversion_timing.h"
#include <map>
#include <set>

namespace {
// 控制台信号属于进程边界；正文策略和结果均为任务内状态。
volatile LONG stopRequested{};
BOOL WINAPI requestStop(DWORD event) {
    if(event!=CTRL_C_EVENT&&event!=CTRL_BREAK_EVENT)return FALSE;
    InterlockedExchange(&stopRequested,1);return TRUE;
}
}
// 反向清单仍只覆盖主故事正文/表格；边界对象保留，不激活任何既有 OLE。
ScanResult scanReverseDocument(IDispatch* document,const ScanOptions& options) {
    // 反向按原生类型及导出的定界符核对，保留旧文档中独占段落的行内对象。
    auto result = scanDocumentText(document, false,false,false,false,options); result.reverse = true;
    if (result.protectedDocument) return result;
    if(options.progress)std::puts(options.saved?"UI_PHASE RESCANNING_OBJECTS":"UI_PHASE SCANNING_OBJECTS");
    auto content = get(document, L"Content"); auto shapes = get(content.Get(), L"InlineShapes");
    std::vector<Span> excluded;
    for (const auto name : {L"Revisions", L"ContentControls", L"OMaths"}) {
        auto collection = get(content.Get(), name);
        const long count=integer(collection.Get(),L"Count");
        for (long index = 1; index <= count; ++index)
            excluded.push_back(bounds(get(collectionItem(collection.Get(), index).Get(), L"Range").Get()));
    }
    auto fields = get(content.Get(), L"Fields");
    for (long index = 1; index <= integer(fields.Get(), L"Count"); ++index) {
        auto field = collectionItem(fields.Get(), index); if (integer(field.Get(), L"Type") == 58) continue;
        excluded.push_back(fieldBounds(field.Get()));
    }
    const long shapeCount=integer(shapes.Get(),L"Count");
    StageProgress objectProgress(options.saved?"RESCAN_OBJECTS":"SCAN_OBJECTS",static_cast<size_t>(shapeCount),options.progress);
    for (long index = 1; index <= shapeCount; ++index) {
        auto shape = collectionItem(shapes.Get(), index); const auto type = integer(shape.Get(), L"Type");
        if (type != 1 && type != 2) {objectProgress.advance();continue;}
        const auto cls = text(get(shape.Get(), L"OLEFormat").Get(), L"ClassType");
        if (cls.find(L"Equation.DSMT") != 0) {objectProgress.advance();continue;}
        auto range = get(shape.Get(), L"Range"); ScanItem item;
        item.span = bounds(range.Get());item.position.sourceStart=item.span.start;item.source = text(range.Get(), L"Text"); item.delimiter = L"MATH_TYPE";
        // 复用文本扫描已取得的段落起点，反向诊断也有真实段号，不再逐对象重新计数全文。
        item.paragraph=static_cast<long>(std::upper_bound(result.paragraphStarts.begin(),result.paragraphStarts.end(),item.span.start)-result.paragraphStarts.begin());
        if (information(range.Get(), 12)) {objectProgress.advance();continue;}
        if (type != 1) item.reason = L"LINKED_MATHTYPE_PRESERVED";
        else if (cls != L"Equation.DSMT4") item.reason = L"UNSUPPORTED_MATHTYPE_CLASS";
        else if (std::any_of(excluded.begin(), excluded.end(), [&](Span span) {return span.start < item.span.end && span.end > item.span.start;}))
            item.reason = L"MATH_TYPE_IN_EXCLUDED_BOUNDARY";
        else {
            try { item.display = !inspectStyle(nativePayload(rangeNativeAsset(range.Get()).equation)).inlineEquation; }
            catch (const std::exception&) { item.reason = L"UNREADABLE_MATHTYPE_NATIVE"; }
        }
        item.ready = item.reason.empty();
        if(item.ready&&!selectedEquation(item.display,options.scope)){objectProgress.advance();continue;}
        result.items.push_back(std::move(item));
        objectProgress.advance();
    }
    std::sort(result.items.begin(), result.items.end(), [](const ScanItem& a, const ScanItem& b) {return a.span.start < b.span.start;});
    long end = -1;
    for (const auto& item : result.items) {require(item.span.start >= end, "Non-overlapping reverse targets"); end = item.span.end;}
    return result;
}

namespace {
static std::vector<std::wstring> nativeMathSnapshot(IDispatch* document,ConversionTimings& timings,bool contentOnly=false) {
    ScopedConversionTiming timing(timings,ConversionMetric::OmmlSnapshot);
    std::vector<std::wstring> snapshot;auto content=get(document,L"Content");auto maths=get(content.Get(),L"OMaths");
    const long count=integer(maths.Get(),L"Count");
    // 无嵌入/浮动对象的正文可批量取 XML，避免为每式重复导出完整的样式包。
    // shortcut: 混合资产文档仍逐式读取以避免导出巨大图片/OLE，成为快照瓶颈时再优化。
    if(contentOnly&&count>3&&integer(get(content.Get(),L"InlineShapes").Get(),L"Count")==0&&integer(get(document,L"Shapes").Get(),L"Count")==0){
        if(ommlSourcesFromBodyXml(text(content.Get(),L"WordOpenXML"),static_cast<size_t>(count),snapshot)){
            // Word.Text 会线性化结构并映射数学字母；用完整 XML 首尾锚点核对，不猜其文本表示。
            const auto first=ommlSource(get(collectionItem(maths.Get(),1).Get(),L"Range").Get());
            const auto last=ommlSource(get(collectionItem(maths.Get(),count).Get(),L"Range").Get());
            if(snapshot.front()==first&&snapshot.back()==last){std::printf("OMML_SNAPSHOT_BATCH COUNT=%ld\n",count);return snapshot;}
        }
        snapshot.clear();
    }
    for (long index = 1; index <= count; ++index)
        snapshot.push_back(contentOnly?ommlSource(get(collectionItem(maths.Get(),index).Get(),L"Range").Get()):sourceSignature(get(collectionItem(maths.Get(), index).Get(), L"Range").Get()));
    return snapshot;
}
static std::vector<long> tableStructure(IDispatch* document) {
    std::vector<long> structure;
    auto tables = get(document, L"Tables"); const auto count = integer(tables.Get(), L"Count");
    structure.push_back(count);
    for (long index = 1; index <= count; ++index) {
        auto table = collectionItem(tables.Get(), index);
        auto cells = get(get(table.Get(), L"Range").Get(), L"Cells");
        const auto cellCount = integer(cells.Get(), L"Count"); structure.push_back(cellCount);
        for (long cell = 1; cell <= cellCount; ++cell) {
            auto item = collectionItem(cells.Get(), cell);
            structure.push_back(integer(item.Get(), L"RowIndex")); structure.push_back(integer(item.Get(), L"ColumnIndex"));
        }
    }
    return structure;
}
static void requireDisplayLayout(IDispatch* document, IDispatch* range) {
    auto paragraph = collectionItem(get(range, L"Paragraphs").Get(), 1);
    const auto para = bounds(get(paragraph.Get(), L"Range").Get()), shape = bounds(range);
    require(whitespaceOnly(text(docRange(document, para.start, shape.start).Get(), L"Text")) &&
            whitespaceOnly(text(docRange(document, shape.end, para.end).Get(), L"Text")), "Display equation is in its own paragraph");
    // 段距继承原文；旧归零断言保留，以便按需恢复。
    // auto format = get(range, L"ParagraphFormat");
    // require(number(format.Get(), L"SpaceBefore") == 0 && number(format.Get(), L"SpaceAfter") == 0, "Display paragraph spacing");
}
// 返回普通失败前必须完成恢复；无法确认恢复时由正文调度重建副本并跳过该式。
// 兼容改写只作用于导入副本，正文和实验区共用 compatibleTex。
// 原生显示布局按文档初始化一次；宽度只在实际节变化时重新读取，不加载交互宏。
struct DisplayLayout {
    IDispatch* document{};bool stylePresent{},addRightNumbers{};EquationNumbering numbering;bool singleSection{};ComPtr<IDispatch> style;std::map<long,double> widths;std::map<std::wstring,ComPtr<IDispatch>> formats;
    void apply(IDispatch* range,const std::wstring* sourceXml=nullptr){
        if(!style){
            auto styles=get(document,L"Styles");
            if(stylePresent)style=asObject(invoke(styles.Get(),L"Item",DISPATCH_METHOD,{L"MTDisplayEquation"}));
            // 新建样式前先把选区放到公式所在段落：Styles.Add 会复制当前选区的字符格式，不能沿用它打开文档时的开头选区。
            else{auto paragraph=collectionItem(get(range,L"Paragraphs").Get(),1);invoke(get(paragraph.Get(),L"Range").Get(),L"Select",DISPATCH_METHOD);
                style=asObject(invoke(styles.Get(),L"Add",DISPATCH_METHOD,{L"MTDisplayEquation",1L}));auto base=get(range,L"Style");invoke(style.Get(),L"BaseStyle",DISPATCH_PROPERTYPUT,{base.Get()});}
            if(!addRightNumbers){nativeProperty(document,L"MTEquationSection",L"1");if(!stylePresent){nativeProperty(document,L"MTEquationNumber2",numbering.format);nativeProperty(document,L"MTCustomEquationNumber",numbering.custom?L"1":L"0");}}
            singleSection=integer(get(document,L"Sections").Get(),L"Count")==1;
        }
        const long section=singleSection?1:information(range,2);
        if(widths.find(section)==widths.end()){
            auto page=get(collectionItem(get(document,L"Sections").Get(),section).Get(),L"PageSetup");
            // 实际栏宽已包含装订线和镜像页边距；不等宽多栏仍沿用原先的首栏排版约定。
            auto columns=get(page.Get(),L"TextColumns");const double width=number(collectionItem(columns.Get(),1).Get(),L"Width");widths[section]=width;
            if(widths.size()==1&&!stylePresent){auto format=get(style.Get(),L"ParagraphFormat");invoke(format.Get(),L"Alignment",DISPATCH_PROPERTYPUT,{0L});auto tabs=get(format.Get(),L"TabStops");invoke(tabs.Get(),L"ClearAll",DISPATCH_METHOD);invoke(tabs.Get(),L"Add",DISPATCH_METHOD,{width/2,1L,0L});invoke(tabs.Get(),L"Add",DISPATCH_METHOD,{width,2L,0L});}
        }
        // 直接复用原流程已有的段落 XML；缺少完整段落属性时不缓存，避免错误复用格式。
        std::wstring key;
        if(singleSection&&sourceXml){const auto first=sourceXml->find(L"<w:pPr"),paragraphEnd=sourceXml->find(L"</w:p>");if(first!=std::wstring::npos&&first<paragraphEnd){const auto closed=sourceXml->find(L"</w:pPr>",first),empty=sourceXml->find(L"/>",first);const auto end=closed!=std::wstring::npos&&closed<paragraphEnd?closed+8:empty+2;if(end>first&&end<=paragraphEnd)key=sourceXml->substr(first,end-first);}}
        if(!key.empty()){const auto found=formats.find(key);if(found!=formats.end()){invoke(range,L"ParagraphFormat",DISPATCH_PROPERTYPUT,{found->second.Get()});return;}}
        auto format=get(get(range,L"ParagraphFormat").Get(),L"Duplicate");
        const double left=number(format.Get(),L"LeftIndent"),right=widths.at(section)-number(format.Get(),L"RightIndent");require(right>left,"Positive display paragraph width");
        // Word 切换样式会重置段落属性，仅保存本工具承诺保留的段距、行距和左右缩进。
        const double before=number(format.Get(),L"SpaceBefore"),after=number(format.Get(),L"SpaceAfter"),line=number(format.Get(),L"LineSpacing");const long rule=integer(format.Get(),L"LineSpacingRule");
        invoke(format.Get(),L"Style",DISPATCH_PROPERTYPUT,{style.Get()});
        invoke(format.Get(),L"SpaceBefore",DISPATCH_PROPERTYPUT,{before});invoke(format.Get(),L"SpaceAfter",DISPATCH_PROPERTYPUT,{after});
        invoke(format.Get(),L"LineSpacingRule",DISPATCH_PROPERTYPUT,{rule});if(rule>=3)invoke(format.Get(),L"LineSpacing",DISPATCH_PROPERTYPUT,{line});
        invoke(format.Get(),L"LeftIndent",DISPATCH_PROPERTYPUT,{left});invoke(format.Get(),L"RightIndent",DISPATCH_PROPERTYPUT,{widths.at(section)-right});
        invoke(format.Get(),L"Alignment",DISPATCH_PROPERTYPUT,{0L});invoke(format.Get(),L"CharacterUnitFirstLineIndent",DISPATCH_PROPERTYPUT,{0L});invoke(format.Get(),L"FirstLineIndent",DISPATCH_PROPERTYPUT,{0L});
        auto tabs=get(format.Get(),L"TabStops");invoke(tabs.Get(),L"ClearAll",DISPATCH_METHOD);invoke(tabs.Get(),L"Add",DISPATCH_METHOD,{(left+right)/2,1L,0L});invoke(tabs.Get(),L"Add",DISPATCH_METHOD,{right,2L,0L});
        invoke(range,L"ParagraphFormat",DISPATCH_PROPERTYPUT,{format.Get()});
        if(!key.empty())formats.emplace(key,format);
    }
};

struct BodyRangeError : std::runtime_error {
    ScanItem item;
    BodyRangeError(const ScanItem& source,const std::exception& error):std::runtime_error(error.what()),item(source){}
};
static ItemOutcome failedPreserved(const ScanItem& item,const std::exception& error) {
    const std::string message=error.what();std::printf("ITEM_PREPARATION_PRESERVED %s\n",message.c_str());
    return {item.span,L"FAILED_PRESERVED",std::wstring(message.begin(),message.end())};
}
class BodyConversion {
    const BodyOptions& options;
    ConversionTimings& timings;
    ScanOptions scanOptions;
    NativeNumberingSession numbers;
    bool batchSavedNative{};
    // 新导出只接收已经验收的 ASCII 单段 TeX；其他情况保持原对象并报告原因。
    template<class ExpectedText>
    ItemOutcome reverseOne(MathTypeApi& api, WordSession& word, const ScanItem& item, bool dollar,
                                  bool selfTest, bool abortTest, const ExpectedText& expectedText) {
        auto document = word.document.Get();ComPtr<IDispatch> range,shapes,undo;long count{},end{};
        std::wstring body,sourceXml,prefix,suffix;
        const bool object = item.delimiter == L"MATH_TYPE";
        try{
        range = docRange(document, item.span.start, item.span.end);
        if(!options.scan.fast)require(text(range.Get(), L"Text") == item.source, "Reverse source rechecked before export");
        shapes = get(get(document, L"Content").Get(), L"InlineShapes");count = integer(shapes.Get(), L"Count");
        end = integer(get(document, L"Content").Get(), L"End");
        // 正常路径复用调度层的正文片段快照，失败恢复时才重新核对完整正文。
        body = selfTest&&!options.scan.fast ? text(get(document, L"Content").Get(), L"Text") : std::wstring();sourceXml = options.scan.fast?std::wstring():sourceSignature(range.Get(), object);
        prefix = selfTest&&!options.scan.fast ? text(docRange(document, 0, item.span.start).Get(), L"Text") : std::wstring();
        suffix = selfTest&&!options.scan.fast ? text(docRange(document, item.span.end, end).Get(), L"Text") : std::wstring();
        undo = get(word.application.Get(), L"UndoRecord");
        }catch(const std::exception& error){return failedPreserved(item,error);}
        NativeAsset asset;
        bool recordOpen = false;
        try {
            // 现有 LaTeX 只改定界符，保留原文中的 Word 换行及空白。
            // 缺失左反斜杠的 [..\] 起止长度不同，截取与替换必须分别计算。
            const size_t leftDelimiter=item.delimiter==L"$"||item.delimiter==L"[\\]"?1:2,rightDelimiter=item.delimiter==L"$"?1:2;
            std::wstring inner = object?item.inner:item.source.substr(leftDelimiter,item.source.size()-leftDelimiter-rightDelimiter);
            if (object) {
                asset = rangeNativeAsset(range.Get()); const auto payload = nativePayload(asset.equation);
                if(!options.scan.fast)require(inspectStyle(payload).inlineEquation == !item.display, "Reverse native type still matches scan");
                inner = exportedInner(exportNativeTex(api, payload), item.display,options.scan.fast);
            }
            const std::wstring open = item.display ? (dollar ? L"$$" : L"\\[") : L"$";
            const std::wstring close = item.display ? (dollar ? L"$$" : L"\\]") : L"$";
            const auto output = open + inner + close;
            invoke(document, L"Activate", DISPATCH_METHOD);
            invoke(undo.Get(), L"StartCustomRecord", DISPATCH_METHOD, {L"MathType 导出/定界符归一"}); recordOpen = true;
            // 已有 LaTeX 只改两端，保留内部原文和字符格式；MathType 成功导出后才替换对象。
            if (object) invoke(range.Get(), L"Text", DISPATCH_PROPERTYPUT, {output.c_str()});
            else if (output != item.source) {
                invoke(docRange(document, item.span.end - static_cast<long>(rightDelimiter), item.span.end).Get(), L"Text", DISPATCH_PROPERTYPUT, {close.c_str()});
                invoke(docRange(document, item.span.start, item.span.start + static_cast<long>(leftDelimiter)).Get(), L"Text", DISPATCH_PROPERTYPUT, {open.c_str()});
            }
            require(!(selfTest && object && item.display), "Injected reverse replacement failure");
            require(!conversionFault(L"DOCX_MATH_BODY_FAULT",std::to_wstring(item.paragraph),L"RECOVERY"),"Injected reverse body replacement failure");
            const long afterEnd = integer(get(document, L"Content").Get(), L"End"), replacementEnd = item.span.end + afterEnd - end;
            if(!options.scan.fast)require((!selfTest||(text(docRange(document, 0, item.span.start).Get(), L"Text") == prefix &&
                    text(docRange(document, replacementEnd, afterEnd).Get(), L"Text") == suffix)) &&
                    text(docRange(document, item.span.start, replacementEnd).Get(), L"Text") == output &&
                    integer(shapes.Get(), L"Count") == count - static_cast<long>(object), "Reverse replaces only the target range");
            if (!options.scan.fast&&item.row) {
                auto replaced = docRange(document, item.span.start, replacementEnd);
                require(information(replaced.Get(), 12) && information(replaced.Get(), 13) == item.row &&
                        information(replaced.Get(), 16) == item.column, "Reverse target remains in original cell");
            }
            invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD); recordOpen = false;
            ItemOutcome outcome{{item.span.start, replacementEnd}, L"SUCCEEDED", L"", std::move(asset.equation)};
            outcome.outputText = output; outcome.removedPreviews = std::move(asset.previews); return outcome;
        } catch (const std::exception& error) {
            if (recordOpen) invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD);
            const bool changed = integer(get(document, L"Content").Get(), L"End") != end || integer(shapes.Get(), L"Count") != count ||
                                 (!options.scan.fast&&sourceSignature(docRange(document, item.span.start, item.span.end).Get(), object) != sourceXml);
            if (changed) {
                const auto undone = invoke(document, L"Undo", DISPATCH_METHOD, {1L});
                require(undone.v.vt == VT_BOOL && undone.v.boolVal != VARIANT_FALSE, "Undo failed reverse replacement");
            }
            if(options.scan.fast){const std::string reason=error.what();return {item.span,L"FAILED_UNVERIFIED",std::wstring(reason.begin(),reason.end())};}
            const auto recoveredXml = sourceSignature(docRange(document, item.span.start, item.span.end).Get(), object);
            require(text(get(document, L"Content").Get(), L"Text") == (selfTest?body:expectedText()) &&
                    integer(get(document,L"Content").Get(),L"End")==end && integer(shapes.Get(), L"Count") == count &&
                    recoveredXml == sourceXml, "Reverse recovery restores source, formatting and OLE");
            if (object && !asset.equation.empty()) {
                const auto recovered = rangeNativeAsset(docRange(document, item.span.start, item.span.end).Get());
                require(recovered.equation == asset.equation && recovered.previews == asset.previews, "Reverse recovery restores complete native contents and previews");
            }
            require(!abortTest, "Injected reverse recovery verification failure");
            require(!conversionFault(L"DOCX_MATH_BODY_FAULT",std::to_wstring(item.paragraph),L"RECOVERY"),"Injected reverse body recovery verification failure");
            const std::string reason = error.what();
            std::printf("REVERSE_ITEM_RECOVERY_PASSED %s\n", reason.c_str());
            return {item.span, L"FAILED_PRESERVED", std::wstring(reason.begin(), reason.end())};
        }
    }
    void validateReverse(IDispatch* document, const ScanResult& before, std::vector<ItemOutcome>& outcomes,
                                const std::wstring& body, const std::vector<long>& tables,
                                const std::vector<std::wstring>& fields, const std::vector<std::wstring>& maths) {
        require(text(get(document, L"Content").Get(), L"Text") == body && tableStructure(document) == tables &&
                numbers.fields(document) == fields && nativeMathSnapshot(document,timings) == maths, "Reverse preserves body, table identities, fields and full OMML");
        const auto after = scanReverseDocument(document,scanOptions); long removed{};
        for (size_t index = 0; index < outcomes.size(); ++index)
            removed += before.items[index].delimiter == L"MATH_TYPE" && outcomes[index].status == L"SUCCEEDED";
        require(after.items.size() == before.items.size() && after.inlineShapes == before.inlineShapes - removed &&
                after.existingMathType == before.existingMathType - removed && after.nativeMath == before.nativeMath &&
                after.fields == before.fields - removed && after.revisions == before.revisions && after.controls == before.controls &&
                after.noteReferences == before.noteReferences, "Reverse target/OLE/EMBED and excluded counts");
        // 保存重开可能改变域坐标，用同一反向重扫的有序清单重新绑定真实位置。
        StageProgress itemProgress("VERIFY_ITEMS",outcomes.size(),scanOptions.progress);
        for (size_t index = 0; index < outcomes.size(); ++index) {
            auto& outcome = outcomes[index]; const auto& original = before.items[index]; const auto& actual = after.items[index];
            require(actual.row == original.row && actual.column == original.column &&
                    actual.source == (outcome.status == L"SUCCEEDED" ? outcome.outputText : original.source), "Every reverse output/residual remains exact");
            if (outcome.status == L"SUCCEEDED") require(actual.delimiter != L"MATH_TYPE" && actual.display == original.display, "Reverse output inline/display type");
            else require(actual.delimiter == original.delimiter && actual.reason == original.reason, "Reverse rejected/failed boundary preserved");
            outcome.finalSpan = actual.span;
            itemProgress.advance();
        }
    }
    ItemOutcome convertOne(MathTypeApi& api, WordSession& word, const ScanItem& item,
                                  const std::vector<unsigned char>& empty,DisplayLayout& displayLayout,ComPtr<IXSLTemplate>& stylesheet,bool selfTest, bool abortTest) {
        auto document = word.document.Get();
        ComPtr<IDispatch> range,shapes,undo;long count{},end{};std::wstring body,sourceXml,sourcePackage,prefix,suffix;
        try{
        range = docRange(document, item.span.start, item.span.end);
        if(!options.scan.fast)require(text(range.Get(), L"Text") == item.source, "Source rechecked immediately before conversion");
        shapes = get(get(document, L"Content").Get(), L"InlineShapes");
        count = integer(shapes.Get(), L"Count");end = integer(get(document, L"Content").Get(), L"End");
        body = selfTest&&!options.scan.fast ? text(get(document, L"Content").Get(), L"Text") : std::wstring();
        if(item.ommlOrdinal||!options.scan.fast){
            ScopedConversionTiming timing(timings,ConversionMetric::SourceXml);
            sourcePackage=item.ommlOrdinal?text(range.Get(),L"WordOpenXML"):std::wstring();
            sourceXml=options.scan.fast?std::wstring():item.ommlOrdinal?sourceSignature(sourcePackage):sourceSignature(range.Get());
        }
        prefix = selfTest&&!options.scan.fast ? text(docRange(document, 0, item.span.start).Get(), L"Text") : std::wstring();
        suffix = selfTest&&!options.scan.fast ? text(docRange(document, item.span.end, end).Get(), L"Text") : std::wstring();
        undo = get(word.application.Get(), L"UndoRecord");
        }catch(const std::exception& error){return failedPreserved(item,error);}
        bool active = false, recordOpen = false;
        ComPtr<IDispatch> object;
        try {
            invoke(document, L"Activate", DISPATCH_METHOD);
            invoke(undo.Get(), L"StartCustomRecord", DISPATCH_METHOD, {L"正文公式单条转换"});recordOpen=true;
            // 顶饰先改为 ASCII 命令；未适配的非 ASCII 内容仍保留原式。
            const bool omml=item.ommlOrdinal!=0;
            require(!(selfTest && item.source == L"$beta_2$"), "Injected preparation failure");
            const auto inner=omml?std::wstring():compatibleTex(item.inner);
            require(options.scan.fast||omml||std::all_of(inner.begin(),inner.end(),[](wchar_t c){return c<=127;}),"UNSUPPORTED_NON_ASCII_TEX");
            if(omml&&!stylesheet)stylesheet=compileOmmlStylesheet(text(word.application.Get(),L"Path")+L"\\OMML2MML.XSL",timings);
            const auto input=omml?ommlMathML(sourcePackage,stylesheet.Get(),item.display,timings):(item.display?L"\\["+inner+L"\\]":L"$"+inner+L"$");
            // 是否拆段按当前段落重算：同段后面的公式可能已经拆段，不能沿用扫描提示。
            bool splitBefore{},splitAfter{};
            if(item.display){
                auto paragraphs=get(range.Get(),L"Paragraphs");
                auto first=collectionItem(paragraphs.Get(),1),last=collectionItem(paragraphs.Get(),integer(paragraphs.Get(),L"Count"));
                const Span para{integer(get(first.Get(),L"Range").Get(),L"Start"),integer(get(last.Get(),L"Range").Get(),L"End")};
                splitBefore=!whitespaceOnly(text(docRange(document,para.start,item.span.start).Get(),L"Text"));
                splitAfter=!whitespaceOnly(text(docRange(document,item.span.end,para.end).Get(),L"Text"));
            }
            const bool leadingTab=item.display;
            // 两个 Tab 与源码一次替换；对象插在中间，段尾预留右对齐位置供原生按钮插入编号。
            // OMML 先替换为对象，离开原生公式区域后再补 Tab，避免破坏可撤销的公式结构。
            invoke(range.Get(), L"Text", DISPATCH_PROPERTYPUT, {item.display&&!omml?L"\t\t":L""});
            if(item.display&&!omml)invoke(range.Get(),L"SetRange",DISPATCH_METHOD,{item.span.start+1,item.span.start+1});else invoke(range.Get(),L"Collapse",DISPATCH_METHOD,{1L});
            require(!(selfTest && item.source == L"$split_1+split_2$"), "Injected creation failure");
            auto shape=insertBlankEquation(document,range.Get(),api.install,nullptr,options.scan.fast);
            auto ole=get(shape.Get(),L"OLEFormat");invoke(ole.Get(),L"DoVerb",DISPATCH_METHOD,{2L});
            object=get(ole.Get(),L"Object");active=true;
            require(api.writeBytes(object.Get(),3,empty.data(),static_cast<long>(empty.size()))==0,"Apply isolated equation preferences");
            taskTestBoundary(L"NATIVE_STYLE_WRITTEN");
            {ScopedConversionTiming timing(timings,omml?ConversionMetric::MathmlImport:ConversionMetric::TexImport);
                require(input.size()<=LONG_MAX&&api.write(object.Get(),omml?2:1,input.c_str(),static_cast<long>(input.size()))==0,omml?"Import OMML into new object":"Write TeX into new object");}
            auto native=options.scan.fast?std::vector<unsigned char>():readNative(api,object.Get());
            if(!options.scan.fast){require(native!=empty&&native.size()>empty.size(),"New equation contains native content");requireStyle(inspectStyle(native),inspectStyle(empty));}
            require(api.close(OLECLOSE_NOSAVE,object.Get())==0,"Close converted equation");active=false;
            if(omml&&item.display){auto inserted=get(shape.Get(),L"Range");const long finish=integer(inserted.Get(),L"End"),start=integer(inserted.Get(),L"Start");
                invoke(docRange(document,finish,finish).Get(),L"InsertBefore",DISPATCH_METHOD,{L"\t"});
                invoke(docRange(document,start,start).Get(),L"InsertBefore",DISPATCH_METHOD,{L"\t"});}
            auto shapeRange = get(shape.Get(), L"Range");
            if (splitAfter) {
                const long position = integer(shapeRange.Get(), L"End")+(item.display?1:0);
                invoke(docRange(document, position, position).Get(), L"InsertBefore", DISPATCH_METHOD, {L"\r"});
            }
            if (splitBefore) {
                const long position = integer(shapeRange.Get(), L"Start")-(leadingTab?1:0);
                invoke(docRange(document, position, position).Get(), L"InsertBefore", DISPATCH_METHOD, {L"\r"});
            }
            shapeRange = get(shape.Get(), L"Range");
            if (item.display) {
                displayLayout.apply(shapeRange.Get(),&sourceXml);
                // 段前段后继续继承，保留旧设置以便按需恢复。
                // invoke(format.Get(), L"SpaceBefore", DISPATCH_PROPERTYPUT, {0L});
                // invoke(format.Get(), L"SpaceAfter", DISPATCH_PROPERTYPUT, {0L});
                if(!options.scan.fast)requireDisplayLayout(document, shapeRange.Get());
            }
            const auto& spacing=item.display?options.displaySpacing:options.inlineSpacing;
            if(spacing.mode){
                auto format=get(shapeRange.Get(),L"ParagraphFormat");
                const long rules[]={0,0,1,2,4,3,5};
                invoke(format.Get(),L"LineSpacingRule",DISPATCH_PROPERTYPUT,{rules[spacing.mode]});
                if(spacing.mode>=4)invoke(format.Get(),L"LineSpacing",DISPATCH_PROPERTYPUT,{spacing.value*(spacing.mode==6?12:1)});
                // 显式行距不受文档网格吸附覆盖；默认跟随原文时不读写这些属性。
                invoke(format.Get(),L"DisableLineHeightGrid",DISPATCH_PROPERTYPUT,{true});
            }
            if(omml){taskTestBoundary(L"BODY_OMML_REPLACED");require(!(selfTest&&item.source.find(L"987654321")!=std::wstring::npos),"Injected body OMML replacement failure");}
            require(!conversionFault(L"DOCX_MATH_BODY_FAULT",std::to_wstring(item.paragraph),L"RECOVERY"),"Injected body replacement failure");
            shapeRange=get(shape.Get(),L"Range");
            require(!(selfTest && item.source == L"$$cell_2$$"), "Injected layout failure");
            const auto replaced = (splitBefore ? std::wstring(L"\r") : L"") + (leadingTab?L"\t":L"") + text(shapeRange.Get(), L"Text") + (item.display?L"\t":L"") + (splitAfter ? L"\r" : L"");
            const long afterEnd = integer(get(document, L"Content").Get(), L"End"), replacementEnd = item.span.end + afterEnd - end;
            // Word.Text 在空单元格处随 Range 边界出现不同末尾标记，不能拼接不同 Range.Text 比较全文。
            // 相同逻辑范围分别核对前后正文，再核对被替换区间，保持真实 Word 坐标而非字符串下标。
            if(!options.scan.fast)require((!selfTest || (text(docRange(document, 0, item.span.start).Get(), L"Text") == prefix &&
                    text(docRange(document, replacementEnd, afterEnd).Get(), L"Text") == suffix)) &&
                    text(docRange(document, item.span.start, replacementEnd).Get(), L"Text") == replaced &&
                    integer(shapes.Get(), L"Count") == count + 1, "Only source text replaced by one OLE and required paragraph breaks");
            if(!options.scan.fast)require(number(shape.Get(), L"Width") > 0 && number(shape.Get(), L"Height") > 0, "New equation dimensions");
            if (!options.scan.fast&&item.row) require(information(shapeRange.Get(), 12) && information(shapeRange.Get(), 13) == item.row &&
                                  information(shapeRange.Get(), 16) == item.column, "New equation remains in original cell");
            invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD); recordOpen = false;
            ItemOutcome outcome{bounds(shapeRange.Get()), L"SUCCEEDED", L"", std::move(native), static_cast<long>(splitBefore) + static_cast<long>(splitAfter)};
            outcome.outputText=replaced;outcome.replacementSpan={item.span.start,replacementEnd};return outcome;
        } catch (const std::exception& error) {
            const auto message = std::string(error.what());
            std::printf("ITEM_FAILURE START=%ld: %s\n", item.span.start, message.c_str());
            if (active) require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close failed task equation");
            if (recordOpen) invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD);
            const bool changed = integer(get(document,L"Content").Get(),L"End") != end || integer(shapes.Get(), L"Count") != count ||
                (!options.scan.fast&&sourceSignature(docRange(document, item.span.start, item.span.end).Get()) != sourceXml);
            // 空准备记录不能 Undo，否则会撤销前一条已经成功的公式。
            if (changed) {
                if(item.ommlOrdinal){
                    // Word 的结构化公式编辑会结束自定义 Undo，按原始原生 XML 恢复本次范围。
                    auto replacement=docRange(document,item.span.start,item.span.end+integer(get(document,L"Content").Get(),L"End")-end);
                    invoke(replacement.Get(),L"InsertXML",DISPATCH_METHOD,{sourcePackage.c_str()});
                    // 行内片段导入会携带尾段落符，仅清理导入新增且已核实的段落标记。
                    const long extra=integer(get(document,L"Content").Get(),L"End")-end;
                    if(!item.display&&extra>0){auto trailing=docRange(document,item.span.end,item.span.end+extra);
                        require(text(trailing.Get(),L"Text")==std::wstring(static_cast<size_t>(extra),L'\r'),"OMML restoration adds only trailing paragraph markers");
                        invoke(trailing.Get(),L"Text",DISPATCH_PROPERTYPUT,{L""});}
                }else{
                    auto result = invoke(document, L"Undo", DISPATCH_METHOD, {1L});
                    require(result.v.vt == VT_BOOL && result.v.boolVal != VARIANT_FALSE, "Undo failed conversion");
                }
            }
            if(options.scan.fast)return {item.span,L"FAILED_UNVERIFIED",std::wstring(message.begin(),message.end())};
            const auto recoveredXml = sourceSignature(docRange(document, item.span.start, item.span.end).Get());
            if(item.ommlOrdinal)std::printf("BODY_OMML_RECOVERY end=%d shapes=%d xml=%d\n",integer(get(document,L"Content").Get(),L"End")==end,integer(shapes.Get(),L"Count")==count,recoveredXml==sourceXml);
            require((!selfTest || text(get(document, L"Content").Get(), L"Text") == body) && integer(get(document,L"Content").Get(),L"End")==end && integer(shapes.Get(), L"Count") == count && recoveredXml == sourceXml,
                    "Recovery must restore exact source text, character/paragraph formatting and object count");
            require(!abortTest, "Injected recovery verification failure");
            require(!conversionFault(L"DOCX_MATH_BODY_FAULT",std::to_wstring(item.paragraph),L"RECOVERY"),"Injected body recovery verification failure");
            std::printf("ITEM_RECOVERY_PASSED: %s\n", message.c_str());
            return {item.span, L"FAILED_PRESERVED", std::wstring(message.begin(), message.end()), {}, 0};
        }
    }
    ItemOutcome formatExistingDisplay(WordSession& word,const ScanItem& item,DisplayLayout& layout,const std::vector<std::vector<unsigned char>>& assets){
        auto document=word.document.Get();auto previousLayout=layout;
        auto range=docRange(document,item.span.start,item.span.end);if(!options.scan.fast)require(text(range.Get(),L"Text")==item.source,"Existing display source remains exact before formatting");
        auto shape=collectionItem(get(range.Get(),L"InlineShapes").Get(),1);auto objectRange=get(shape.Get(),L"Range");const auto before=integer(get(document,L"Content").Get(),L"End");
        const auto sourceXml=options.scan.fast?std::wstring():sourceSignature(range.Get(),true);const auto native=options.scan.fast?std::vector<unsigned char>():assets.empty()?nativePayload(rangeNativeAsset(objectRange.Get()).equation):assets.at(static_cast<size_t>(item.existingNativeOrdinal-1));
        auto undo=get(word.application.Get(),L"UndoRecord");bool record=false;
        try{
        invoke(document,L"Activate",DISPATCH_METHOD);invoke(undo.Get(),L"StartCustomRecord",DISPATCH_METHOD,{L"已有公式单段排版"});record=true;
        auto prefix=docRange(document,item.span.start,integer(objectRange.Get(),L"Start"));
        if(text(prefix.Get(),L"Text")!=L"\t")invoke(prefix.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"\t"});
        objectRange=get(shape.Get(),L"Range");
        const long end=integer(objectRange.Get(),L"End");if(text(docRange(document,end,end+1).Get(),L"Text")!=L"\t")invoke(docRange(document,end,end).Get(),L"InsertBefore",DISPATCH_METHOD,{L"\t"});
        objectRange=get(shape.Get(),L"Range");
        layout.apply(objectRange.Get());
        require(!conversionFault(L"DOCX_MATH_BODY_FAULT",std::to_wstring(item.paragraph),L"FORMAT"),"Injected existing display formatting failure");
        const long delta=integer(get(document,L"Content").Get(),L"End")-before;
        ItemOutcome outcome{bounds(objectRange.Get()),L"EXISTING_FORMATTED",L"",native,0};
        outcome.replacementSpan={item.span.start,item.span.end+delta};outcome.outputText=text(docRange(document,outcome.replacementSpan.start,outcome.replacementSpan.end).Get(),L"Text");
        invoke(undo.Get(),L"EndCustomRecord",DISPATCH_METHOD);record=false;return outcome;
        }catch(const std::exception& error){
            if(record)invoke(undo.Get(),L"EndCustomRecord",DISPATCH_METHOD);
            if(integer(get(document,L"Content").Get(),L"End")!=before||(!options.scan.fast&&sourceSignature(docRange(document,item.span.start,item.span.end).Get(),true)!=sourceXml)){
                auto undone=invoke(document,L"Undo",DISPATCH_METHOD,{1L});require(undone.v.vt==VT_BOOL&&undone.v.boolVal,"Undo failed existing display formatting");}
            if(options.scan.fast){layout=std::move(previousLayout);const std::string message=error.what();ItemOutcome outcome{item.span,L"EXISTING_PRESERVED",std::wstring(message.begin(),message.end())};outcome.replacementSpan=item.span;return outcome;}
            require(integer(get(document,L"Content").Get(),L"End")==before&&sourceSignature(docRange(document,item.span.start,item.span.end).Get(),true)==sourceXml,"Existing display recovery restores original paragraph");
            auto original=collectionItem(get(docRange(document,item.span.start,item.span.end).Get(),L"InlineShapes").Get(),1);auto originalRange=get(original.Get(),L"Range");
            require(nativePayload(rangeNativeAsset(originalRange.Get()).equation)==native,"Existing display recovery preserves original native contents");
            layout=std::move(previousLayout);const std::string message=error.what();ItemOutcome outcome{bounds(originalRange.Get()),L"EXISTING_PRESERVED",std::wstring(message.begin(),message.end()),native};outcome.replacementSpan=item.span;return outcome;
        }
    }
    void bindShapeOrdinals(IDispatch* document,std::vector<ItemOutcome>& outcomes) {
        auto shapes=get(get(document,L"Content").Get(),L"InlineShapes");
        std::map<std::pair<long,long>,long> ordinals;
        std::map<std::pair<long,long>,long> nativeOrdinals;long nativeOrdinal{};
        const long count=integer(shapes.Get(),L"Count");
        StageProgress objectProgress("BIND_OBJECTS",static_cast<size_t>(count),scanOptions.progress);
        // 一次记录对象位置，替代每条新公式遍历整份对象列表。
        for (long index=1;index<=count;++index) {
            auto shape=collectionItem(shapes.Get(),index);const auto span=bounds(get(shape.Get(),L"Range").Get());
            ordinals.emplace(std::make_pair(span.start,span.end),index);
            const long type=integer(shape.Get(),L"Type");
            if((type==1||type==2)&&text(get(shape.Get(),L"OLEFormat").Get(),L"ClassType").find(L"Equation.DSMT")==0)nativeOrdinals.emplace(std::make_pair(span.start,span.end),++nativeOrdinal);
            objectProgress.advance();
        }
        for (auto& outcome:outcomes) if(nativeOutcome(outcome)) {
            const auto found=ordinals.find({outcome.finalSpan.start,outcome.finalSpan.end});
            require(found!=ordinals.end(),"New OLE ordinal recorded before save");outcome.shapeOrdinal=found->second;
            outcome.nativeOrdinal=nativeOrdinals.at({outcome.finalSpan.start,outcome.finalSpan.end});
        }
    }
    void validateBeforeSave(IDispatch* document,const ScanResult& before,const std::vector<ItemOutcome>& outcomes,
                                   const std::wstring& body,const std::vector<long>& tables,
                                   const std::vector<std::wstring>& fields,const std::vector<std::wstring>& maths) {
        require(text(get(document,L"Content").Get(),L"Text")==body && tableStructure(document)==tables && numbers.fields(document)==fields,
                "Conversion preserves body, table identities and ordinary fields");
        long delta{},removedMath{};
        for(size_t index=0;index<outcomes.size();++index) if(outcomes[index].status==L"SUCCEEDED")
            {delta+=before.reverse ? (before.items[index].delimiter==L"MATH_TYPE"?-1:0) : 1;removedMath+=before.items[index].ommlOrdinal!=0;}
        auto content=get(document,L"Content");
        require(integer(get(content.Get(),L"InlineShapes").Get(),L"Count")==before.inlineShapes+delta &&
                integer(get(content.Get(),L"Fields").Get(),L"Count")==before.fields+delta+numbers.summary.addedFields &&
                integer(get(content.Get(),L"OMaths").Get(),L"Count")==before.nativeMath-removedMath &&
                integer(get(content.Get(),L"Revisions").Get(),L"Count")==before.revisions &&
                integer(get(content.Get(),L"ContentControls").Get(),L"Count")==before.controls &&
                integer(get(document,L"Footnotes").Get(),L"Count")+integer(get(document,L"Endnotes").Get(),L"Count")==before.noteReferences,
                "Conversion object and excluded boundary counts");
        if(before.reverse)require(nativeMathSnapshot(document,timings)==maths,"Conversion preserves full OMML");
        if(before.bodyOmml)require(nativeMathSnapshot(document,timings,true)==maths,"Unconverted OMML contents and order preserved");
    }
    void validateDocument(MathTypeApi& api, IDispatch* document, const ScanResult& before,
                                 std::vector<ItemOutcome>& outcomes, const std::wstring& expectedBody,
                                 const std::vector<long>& tables, const std::vector<std::wstring>& fields,const std::vector<std::wstring>& maths, bool reopen, bool readback=true,
                                 std::vector<LayoutWarning>* layoutWarnings=nullptr) {
        std::puts("TIMING_STAGE SAVED_STRUCTURE");
        require(text(get(document, L"Content").Get(), L"Text") == expectedBody, "Saved and reopened body equals verified conversion body");
        require(tableStructure(document) == tables, "Table cells and row/column identities unchanged");
        std::puts("TIMING_STAGE SAVED_SCAN");
        const auto after = scanDocumentText(document,true,true,false,before.bodyOmml,scanOptions);
        require(numbers.fields(document,&after.fieldSnapshots) == fields, "Existing non-embedded field code and results unchanged");
        const auto succeeded = std::count_if(outcomes.begin(), outcomes.end(), [](const ItemOutcome& item) { return item.status == L"SUCCEEDED"; });
        long removedMath{};for(size_t i=0;i<outcomes.size();++i)removedMath+=before.items[i].ommlOrdinal&&outcomes[i].status==L"SUCCEEDED";
        require(after.inlineShapes == before.inlineShapes + succeeded && after.existingMathType == before.existingMathType + succeeded &&
                after.nativeMath == before.nativeMath-removedMath && after.fields == before.fields + succeeded + numbers.summary.addedFields && after.revisions == before.revisions &&
                after.controls == before.controls && after.noteReferences == before.noteReferences, "Object and excluded boundary counts unchanged");
        if(before.bodyOmml)require(nativeMathSnapshot(document,timings,true)==maths,"Saved unconverted OMML contents and order preserved");
        auto shapes = get(get(document, L"Content").Get(), L"InlineShapes");
        std::puts("TIMING_STAGE SAVED_NATIVE");
        const auto checked=std::count_if(outcomes.begin(),outcomes.end(),nativeOutcome);
        const auto nativeAssets=reopen&&!readback&&checked>0?bodyNativeAssets(document):std::vector<std::vector<unsigned char>>();
        const bool batchNative=reopen&&!readback&&nativeAssets.size()==static_cast<size_t>(after.existingMathType);
        if(reopen&&!readback)batchSavedNative=batchNative&&checked>0;
        if(!reopen)bindShapeOrdinals(document,outcomes);
        size_t remaining{};
        // 保存重开会规范化 EMBED 域代码，Word 坐标可变；先用重扫/对象顺序取得新快照。
        // 回读新对象也可能更新域代码长度，必须在全部范围核对之后再激活。
        std::puts("TIMING_STAGE SAVED_POSITIONS");
        std::puts("TIMING_STAGE SAVED_ITEMS");
        StageProgress itemProgress("VERIFY_ITEMS",outcomes.size(),scanOptions.progress);
        std::map<long,double> layoutWidths;
        for (size_t index = 0; index < outcomes.size(); ++index) {
            auto& outcome = outcomes[index]; const auto& item = before.items[index];
            if(item.existingReference){require(outcome.status==L"REFERENCE_UPDATED"||outcome.status==L"REFERENCE_PRESERVED","Planned number reference was updated or preserved");itemProgress.advance();continue;}
            if (!nativeOutcome(outcome)) {
                require(remaining < after.items.size() && after.items[remaining].source == item.source &&
                        after.items[remaining].reason == item.reason && after.items[remaining].row == item.row &&
                        after.items[remaining].column == item.column, "Ordered residuals match original failed/rejected list");
                if (reopen) outcome.finalSpan = after.items[remaining].span;
                ++remaining;
                auto range=docRange(document,outcome.finalSpan.start,outcome.finalSpan.end);
                require(text(range.Get(), L"Text") == item.source, "Every failed/rejected/unattempted source remains exact");
                itemProgress.advance();continue;
            }
            const auto& saved=after.shapes.at(static_cast<size_t>(outcome.shapeOrdinal-1));
            if(reopen)outcome.finalSpan=saved.span;
            // 扫描已记录顺序、范围与类名；逐式检查复用快照，仅取一次实际对象。
            auto shape=collectionItem(shapes.Get(),outcome.shapeOrdinal);auto shapeRange=get(shape.Get(),L"Range");
            require(integer(get(shapeRange.Get(), L"InlineShapes").Get(), L"Count") == 1, "Converted target has exactly one OLE");
            require(saved.classType == L"Equation.DSMT4", "Saved new object is editable MathType");
            // 旧行外段落允许保留原生编号；扫描已核对边界，编号仍由普通域快照逐项核验。
            if (item.display&&!options.addRightNumbers&&!item.existingDisplay) requireDisplayLayout(document, shapeRange.Get());
            if (item.row) require(information(shapeRange.Get(), 12) && information(shapeRange.Get(), 13) == item.row &&
                                  information(shapeRange.Get(), 16) == item.column, "Saved new OLE remains in original cell");
            const double width=number(shape.Get(),L"Width"),height=number(shape.Get(),L"Height");
            require(width>0&&height>0,"Saved OLE dimensions");
            if(layoutWarnings&&displayOutcome(outcome)){
                const LayoutWarning context{index+1,item.row?L"正文表格":L"正文",L"",item.source,item.position};
                try{
                appendFixedLineSpacingWarning(shapeRange.Get(),height,context,*layoutWarnings);
                // 成功行外式已经由本流程拆为独占段落；编号仍在同一段，不估算其占宽。
                if(item.display&&!item.row){
                    const long section=information(shapeRange.Get(),2);
                    if(layoutWidths.find(section)==layoutWidths.end()){
                        auto page=get(collectionItem(get(document,L"Sections").Get(),section).Get(),L"PageSetup");
                        auto columns=get(page.Get(),L"TextColumns");
                        // 无法可靠确定不等宽栏的实际栏号时，跳过宽度，避免用首栏误判。
                        layoutWidths[section]=integer(columns.Get(),L"Count")==1||integer(columns.Get(),L"EvenlySpaced")==-1?
                            number(collectionItem(columns.Get(),1).Get(),L"Width"):0;
                    }
                    if(layoutWidths.at(section)>0){
                        auto format=get(shapeRange.Get(),L"ParagraphFormat");
                        const double limit=displayWidthLimit(layoutWidths.at(section),number(format.Get(),L"LeftIndent"),number(format.Get(),L"RightIndent"));
                        appendDimensionWarning(L"FORMULA_TOO_WIDE",width,limit,context,*layoutWarnings);
                    }
                }
                }catch(const std::exception& error){
                    require(!taskCancelled()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");
                    std::printf("LAYOUT_CHECK_FAILED %s\n",error.what());auto warning=context;warning.code=L"LAYOUT_CHECK_UNAVAILABLE";layoutWarnings->push_back(std::move(warning));
                }
            }
            // 只读导出当前范围的原生资产，核对公式与位置的对应关系；不激活 MathType。
            if(reopen&&!readback)require((batchNative?nativeAssets.at(static_cast<size_t>(outcome.nativeOrdinal-1)):nativePayload(rangeNativeAsset(shapeRange.Get()).equation))==outcome.native,
                                        "Reopened target native content and style identical at its original position");
            itemProgress.advance();
        }
        require(remaining == after.items.size(), "No unexpected residual items");
        if (reopen && readback) {
            for (const auto& outcome : outcomes) {
                if (outcome.status != L"SUCCEEDED") continue;
                auto ole = get(collectionItem(shapes.Get(), outcome.shapeOrdinal).Get(), L"OLEFormat");
                // 只激活本任务的新对象；既有公式始终仅从包里读流。
                invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L}); auto object = get(ole.Get(), L"Object");
                try {
                    const auto native = readNative(api, object.Get());
                    require(native == outcome.native, "Reopened new native content and style identical");
                    require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close reopened new equation");
                } catch (...) {
                    require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close failed reopened equation"); throw;
                }
            }
            require(text(get(document, L"Content").Get(), L"Text") == expectedBody && numbers.fields(document) == fields,
                    "Native verification leaves ordinary text and fields unchanged");
        }
    }
public:
    BodyConversion(const BodyOptions& value,ConversionTimings& taskTimings):options(value),timings(taskTimings),scanOptions(value.scan),numbers(value.numbering,value.scan.progress){}
    BodyResult run(WordSession& word,MathTypeApi& api,const std::wstring& input,const std::wstring& candidate,
        const std::wstring& eqp,const PackageState& originalPackage,const ExperimentalOptions& experimentalOptions,
        const std::wstring& sizePreferencePath,bool& sizePreferenceOwned,bool& previewOwned) {
        BodyResult result;auto& scan=result.scan;auto& outcomes=result.outcomes;auto& issues=result.issues;
        auto& afterPackage=result.package;auto& experimental=result.experimental;
        size_t& total=result.total;size_t& completed=result.completed;size_t& succeeded=result.succeeded;
        const bool reverse=options.reverse,selfTest=options.selfTest,abortTest=options.abortTest,stopTest=options.stopTest,guiWorker=scanOptions.progress;
        const bool followSize=!reverse&&eqp==L"paragraph",fixedSize=!reverse&&eqp.rfind(L"size=",0)==0;
        const double requestedPoints=options.requestedPoints;
        auto documents=get(word.application.Get(),L"Documents");
        auto isStopping=[&](){
            if(!options.stopFile.empty()&&GetFileAttributesW(options.stopFile.c_str())!=INVALID_FILE_ATTRIBUTES)InterlockedExchange(&stopRequested,1);
            return taskCancelled()||InterlockedCompareExchange(&stopRequested,0,0)!=0;
        };
        std::vector<ScanItem> skippedBody;std::vector<ProcessingIssue> recoveryIssues;std::wstring experimentalSkip;
        std::vector<SourcePosition> originalPositions;bool positionsCaptured=false;
        experimental.options=experimentalOptions;experimental.enabled=experimentalOptions.enabled();
        if(options.scan.fast)std::puts("UI_VALIDATION_SKIPPED EXPERIMENTAL_FAST");
        for(;;){try{
        // 原件始终持有只读锁；只有单式恢复无法确认时才重建副本并跳过该式，正常路径不新增磁盘备份。
        issues=recoveryIssues;result.warnings.clear();numbers.reset();scanOptions.saved=false;
        std::puts("UI_PHASE SCANNING");
        scan = reverse ? scanReverseDocument(word.document.Get(),scanOptions) : scanDocumentText(word.document.Get(),true,false,!options.scan.fast||options.addRightNumbers,true,scanOptions);
        // 两种模式均复用扫描状态拒绝不可编辑输入，不再单独读取修订属性。
        require(!scan.tracksChanges,"Disable revision tracking in source before batch conversion");
        require(!scan.protectedDocument,"Protected document cannot be converted");
        if(options.addRightNumbers)numbers.includeReferences(word.document.Get(),scan);
        for(auto& item:scan.items)if(item.position.sourceStart<0)item.position.sourceStart=item.span.start;
        if(!positionsCaptured){
            ScopedConversionTiming timing(timings,ConversionMetric::SourceLocations);
            std::puts("UI_PHASE SOURCE_LOCATIONS");bool paginated=false;
            if(!scan.items.empty()||experimentalOptions.enabled()){
                require(!isStopping(),"TASK_CANCELLED");
                try{invoke(word.document.Get(),L"Repaginate",DISPATCH_METHOD);paginated=true;}
                catch(const std::exception& error){require(!isStopping()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");std::printf("SOURCE_PAGINATION_UNAVAILABLE %s\n",error.what());}
            }
            for(auto& item:scan.items){require(!isStopping(),"TASK_CANCELLED");item.position=originalSourcePosition(word.document.Get(),item.position.sourceStart,L"正文",paginated);originalPositions.push_back(item.position);}
            result.experimentalSourceLocations=captureExperimentalSourceLocations(word.document.Get(),input,experimentalOptions,reverse,scanOptions,paginated);
            positionsCaptured=true;
        }else{
            // 重建来自同一只读原件；复用首次页码，不再分页或读取当前排版。
            for(size_t i=0;i<scan.items.size();++i){auto& item=scan.items[i];
                if(scan.items.size()==originalPositions.size()&&item.position.sourceStart==originalPositions[i].sourceStart)item.position=originalPositions[i];
                else item.position.location=L"原文页码未知 · 正文";
            }
        }
        if (guiWorker) emitGuiScan(scan);
        // 识别拒绝项在初始化接口前发送，后续整份任务失败时也能显示逐式中文原因。
        if(guiWorker)for(size_t index=0;index<scan.items.size();++index){const auto& item=scan.items[index];
            if(!item.ready&&!item.existingDisplay&&!item.existingReference)std::printf("UI_ITEM_REJECTED %zu %ld %ld %ld %s %s %s\n",index+1,item.paragraph,item.span.start,item.span.end,diagnosticExcerpt(item.source).c_str(),wireHex(utf8(item.reason)).c_str(),wireHex(utf8(item.position.location)).c_str());
        }
        std::puts("UI_PHASE INITIALIZING");
        if(options.addRightNumbers)for(const auto& item:scan.items)if(item.existingDisplay||item.existingReference)numbers.remember(word.document.Get(),item.span);
        const auto tables = options.scan.fast?std::vector<long>():tableStructure(word.document.Get()); const auto fields = options.scan.fast?std::vector<std::wstring>():numbers.fields(word.document.Get());
        auto maths = !options.scan.fast&&(reverse||scan.bodyOmml) ? nativeMathSnapshot(word.document.Get(),timings,!reverse) : std::vector<std::wstring>();
        const bool needsMathType=std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.ready;});
        if(needsMathType&&!api.initialized){require(api.init(0,10)>0,"MathType initialization");api.initialized=true;taskMathTypeConnected();}
        const auto prefs = reverse||followSize||fixedSize||!needsMathType ? std::string() : loadPrefs(api, eqp), globalPrefs = needsMathType&&!options.scan.fast?readDefaultPrefs(api):std::string();
        std::vector<unsigned char> naturalEmpty;
        if(!reverse&&std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.ready;}))naturalEmpty=emptyEquation(api,word,nullptr,true,options.scan.fast);
        std::map<double,std::string> sizePrefs;
        std::map<std::pair<double,bool>,std::vector<unsigned char>> emptyStyles;
        // 正文按目标之间的不变片段只读取一次，独立构造预期结果，避免逐式读取全文。
        std::vector<std::wstring> bodyGaps;long previousEnd{};
        std::wstring originalBody;
        if(!options.scan.fast){
        for(const auto& item:scan.items){bodyGaps.push_back(text(docRange(word.document.Get(),previousEnd,item.span.start).Get(),L"Text"));previousEnd=item.span.end;}
        const long contentEnd=integer(get(word.document.Get(),L"Content").Get(),L"End");
        // 末段公式覆盖到正文末尾时，没有尾片段；Word 不接受在排他末端创建折叠 Range。
        bodyGaps.push_back(previousEnd==contentEnd?std::wstring():text(docRange(word.document.Get(),previousEnd,contentEnd).Get(),L"Text"));
        for(size_t i=0;i<scan.items.size();++i)originalBody+=bodyGaps[i]+scan.items[i].source;
        originalBody+=bodyGaps.back();
        require(originalBody==text(get(word.document.Get(),L"Content").Get(),L"Text"),"Independent body snapshot uses exact Word text ranges");
        }
        outcomes.clear();
        for (const auto& item : scan.items){outcomes.push_back({item.span, item.ready ? L"NOT_ATTEMPTED" : item.existingDisplay?L"EXISTING_PENDING":item.existingReference?L"REFERENCE_PRESERVED":L"REJECTED_PRESERVED", item.reason, {}, 0});outcomes.back().replacementSpan=item.span;}
        auto expectedText=[&](){std::wstring body;for(size_t i=0;i<scan.items.size();++i)body+=bodyGaps[i]+(changedOutcome(outcomes[i])?outcomes[i].outputText:scan.items[i].source);return body+bodyGaps.back();};
        const bool existingDisplays=std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.existingDisplay;});
        const auto existingAssets=existingDisplays&&!options.scan.fast?bodyNativeAssets(word.document.Get()):std::vector<std::vector<unsigned char>>();
        DisplayLayout displayLayout{word.document.Get(),originalPackage.displayStyle,options.addRightNumbers,options.numbering};
        ComPtr<IXSLTemplate> ommlStylesheet;
        total = static_cast<size_t>(std::count_if(scan.items.begin(), scan.items.end(), [](const ScanItem& item) { return item.ready; }));completed=0;
        if (!guiWorker && !taskShared) require(SetConsoleCtrlHandler(requestStop, TRUE) != 0, "Install safe stop handler");
        std::puts("UI_PHASE CONVERTING");
        for (size_t index = scan.items.size(); index-- > 0;) {
            const auto& item = scan.items[index];
            if(item.existingReference)continue;
            if(std::any_of(skippedBody.begin(),skippedBody.end(),[&](const ScanItem& skipped){return skipped.span.start==item.span.start&&skipped.span.end==item.span.end&&skipped.source==item.source;})){
                outcomes[index].reason=L"BODY_RANGE_SKIPPED";
                if(item.existingDisplay){auto shape=collectionItem(get(docRange(word.document.Get(),item.span.start,item.span.end).Get(),L"InlineShapes").Get(),1);auto range=get(shape.Get(),L"Range");outcomes[index].finalSpan=bounds(range.Get());outcomes[index].native=existingAssets.empty()?nativePayload(rangeNativeAsset(range.Get()).equation):existingAssets.at(static_cast<size_t>(item.existingNativeOrdinal-1));outcomes[index].status=L"EXISTING_PRESERVED";}
                continue;
            }
            if(!item.ready&&!item.existingDisplay){std::printf("ITEM_REJECTED ID=%zu PARAGRAPH=%ld START=%ld END=%ld REASON=%s\n",index+1,item.paragraph,item.span.start,item.span.end,utf8(item.reason).c_str());continue;}
            require(!isStopping(),"TASK_CANCELLED");
            std::printf("ITEM_BEGIN ID=%zu PARAGRAPH=%ld START=%ld END=%ld TYPE=%s\n",index+1,item.paragraph,item.span.start,item.span.end,item.display?"DISPLAY":"INLINE");
            if(guiWorker)std::printf("UI_ITEM_BEGIN %zu %ld %ld %ld %s %s\n",index+1,item.paragraph,item.span.start,item.span.end,diagnosticExcerpt(item.source).c_str(),wireHex(utf8(item.position.location)).c_str());
            long beforeEnd{};
            try{
            beforeEnd=integer(get(word.document.Get(), L"Content").Get(), L"End");
            if(item.existingDisplay){outcomes[index]=formatExistingDisplay(word,item,displayLayout,existingAssets);
                if(outcomes[index].status==L"EXISTING_PRESERVED"){issues.push_back({L"body:"+std::to_wstring(index+1),L"正文排版",L"FORMATTING",outcomes[index].reason,item.position,item.source});emitProcessingIssue(issues.back());}}
            else if(reverse)outcomes[index]=reverseOne(api,word,item,eqp==L"dollar",selfTest,abortTest,expectedText);
            else {
                const double points=followSize?item.points:fixedSize?requestedPoints:0;
                const auto key=std::make_pair(points,item.display);
                try{
                require(!conversionFault(L"DOCX_MATH_BODY_FAULT",std::to_wstring(item.paragraph),L"PREPARE"),"Injected formula style preparation failure");
                if(emptyStyles.find(key)==emptyStyles.end()) {
                    if(points!=0&&sizePrefs.find(points)==sizePrefs.end())sizePrefs[points]=sizedPreferences(api,points,sizePreferencePath,sizePreferenceOwned);
                    emptyStyles[key]=styledNative(api,naturalEmpty,points==0?prefs:sizePrefs.at(points),item.display);
                }
                }catch(const std::exception& error){outcomes[index]=failedPreserved(item,error);}
                if(outcomes[index].status!=L"FAILED_PRESERVED")outcomes[index]=convertOne(api,word,item,emptyStyles.at(key),displayLayout,ommlStylesheet,selfTest,abortTest);
            }
            }catch(const std::exception& error){if(abortTest)throw;require(!isStopping(),"TASK_CANCELLED");
                if(options.scan.fast){const std::string message=error.what();outcomes[index]={item.span,L"FAILED_UNVERIFIED",std::wstring(message.begin(),message.end())};}
                else{if(needsMathType)require(readDefaultPrefs(api)==globalPrefs,"Global MathType defaults unchanged");throw BodyRangeError(item,error);}}
            if(!displayOutcome(outcomes[index]))outcomes[index].replacementSpan=outcomes[index].finalSpan;
            const long delta = integer(get(word.document.Get(), L"Content").Get(), L"End") - beforeEnd;
            // 前面的替换只移动已经处理过的后方坐标；当前及更前方目标仍使用刚重扫的原位置。
            for (size_t later = index + 1; later < outcomes.size(); ++later) {
                outcomes[later].finalSpan.start += delta; outcomes[later].finalSpan.end += delta;
                outcomes[later].replacementSpan.start+=delta;outcomes[later].replacementSpan.end+=delta;
            }
            if(item.existingDisplay){if(guiWorker)std::puts("UI_ITEM_END");continue;}
            const auto itemState=outcomes[index].status==L"SUCCEEDED"?"SUCCEEDED":outcomes[index].status==L"FAILED_UNVERIFIED"?"FAILED_UNVERIFIED":"FAILED_PRESERVED";
            std::printf("CONVERT_PROGRESS %zu/%zu ID=%zu START=%ld STATUS=%s\n", ++completed, total, index + 1, item.span.start,
                itemState);
            std::printf("UI_ITEM_RESULT %zu %s %s\n", index + 1, itemState,
                        wireHex(utf8(outcomes[index].reason)).c_str());
            if (stopTest && completed == 3) {InterlockedExchange(&stopRequested, 1);if(taskShared)InterlockedExchange(&taskShared->cancelRequested,1);} // 自检同样丢弃结果。
        }
        if (!guiWorker && !taskShared) require(SetConsoleCtrlHandler(requestStop, FALSE) != 0, "Remove safe stop handler");
        require(!isStopping(),"TASK_CANCELLED");
        // 成功的正文 OMML 从保留清单移除，其余公式仍逐式核对内容和顺序。
        if(!options.scan.fast&&scan.bodyOmml)for(size_t i=scan.items.size();i-- >0;)if(scan.items[i].ommlOrdinal&&outcomes[i].status==L"SUCCEEDED")maths.erase(maths.begin()+scan.items[i].ommlOrdinal-1);
        std::puts("TIMING_STAGE NUMBERING");
        if(options.addRightNumbers&&std::any_of(outcomes.begin(),outcomes.end(),displayOutcome)){
            if(options.scan.fast){
                try{numbers.apply(word.document.Get(),scan,outcomes);}
                catch(const std::exception& error){require(!isStopping(),"TASK_CANCELLED");const std::string reason=error.what();issues.push_back({L"numbering",L"正文编号及交叉引用",L"NUMBERING",std::wstring(reason.begin(),reason.end()),{-1,0,L"正文编号及交叉引用（涉及多个位置）"}});emitProcessingIssue(issues.back());}
            }else{
            // 连续编号和 REF 必须一起成功；只在启用编号时保存一个阶段检查点。
            validateBeforeSave(word.document.Get(),scan,outcomes,expectedText(),tables,fields,maths);
            bindShapeOrdinals(word.document.Get(),outcomes);invoke(word.document.Get(),L"Save",DISPATCH_METHOD);
            ConversionCheckpoint checkpoint(candidate,L".numbering-checkpoint");const auto beforeNumbering=outcomes;auto numbersBefore=numbers;numbersBefore.freeze();const auto bodyBefore=expectedText();
            try{numbers.apply(word.document.Get(),scan,outcomes);}
            catch(const std::exception& error){require(!isStopping(),"TASK_CANCELLED");const std::string reason=error.what();
                displayLayout.style.Reset();displayLayout.formats.clear();numbers.tracked.clear();invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();checkpoint.restore(candidate);
                outcomes=beforeNumbering;numbers=numbersBefore;numbers.summary={};
                word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{candidate.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));
                displayLayout.document=word.document.Get();scanOptions.saved=true;validateDocument(api,word.document.Get(),scan,outcomes,bodyBefore,tables,fields,maths,true,false);scanOptions.saved=false;
                issues.push_back({L"numbering",L"正文编号及交叉引用",L"NUMBERING",std::wstring(reason.begin(),reason.end()),{-1,0,L"正文编号及交叉引用（涉及多个位置）"}});emitProcessingIssue(issues.back());std::puts("NUMBERING_CHECKPOINT_RESTORED");
            }
            }
        }
        displayLayout.style.Reset();
        displayLayout.formats.clear();
        const auto expectedBody=options.scan.fast?std::wstring():expectedText();
        if(!options.scan.fast){std::puts("UI_PHASE VERIFYING");
        if(selfTest) {
            if(reverse)validateReverse(word.document.Get(),scan,outcomes,expectedBody,tables,fields,maths);
            else validateDocument(api,word.document.Get(),scan,outcomes,expectedBody,tables,fields,maths,false);
        } else {
            // 单条转换已核对内容与恢复；完整重扫只在保存重开后执行一次。
            validateBeforeSave(word.document.Get(),scan,outcomes,expectedBody,tables,fields,maths);
            std::puts("TIMING_STAGE BIND_OBJECTS");
            if(!reverse)bindShapeOrdinals(word.document.Get(),outcomes);
        }
        }
        if(needsMathType&&!options.scan.fast)require(readDefaultPrefs(api) == globalPrefs, "Global MathType defaults unchanged");
        if(options.scan.fast&&experimentalOptions.enabled()){
            std::puts("UI_EXPERIMENTAL_PHASE PREPARING");
            // 快速模式沿用已打开的工作副本；页眉引用从只读原件读取，避免与 Word 的文件锁冲突。
            experimental=convertExperimental(word,api,input,experimentalOptions,reverse,eqp==L"dollar",scan.items.size()+1,originalPackage,selfTest,scanOptions,timings,result.experimentalSourceLocations);
        }
        std::puts("UI_PHASE SAVING");
        taskTestBoundary(L"BEFORE_SAVE");
        invoke(word.document.Get(), L"Save", DISPATCH_METHOD);
        if(!options.scan.fast)numbers.freeze();
        else numbers.tracked.clear();
        invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L}); word.document.Reset();
        taskTestBoundary(L"SAVED_DOCUMENT");
        succeeded = static_cast<size_t>(std::count_if(outcomes.begin(), outcomes.end(), [](const ItemOutcome& item) { return item.status == L"SUCCEEDED"; }));
        if(!options.scan.fast){
        std::puts("UI_PHASE REOPENING");
        word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD,
            {candidate.c_str(), false, true, false, Value(), Value(), false, Value(), Value(), Value(), Value(), false}));
        require(boolProperty(word.document.Get(), L"ReadOnly"), "Verification opens candidate read-only");
        scanOptions.saved=true;std::puts("UI_PHASE RESCANNING");
        if (reverse) validateReverse(word.document.Get(), scan, outcomes, expectedBody, tables, fields, maths);
        else validateDocument(api, word.document.Get(), scan, outcomes, expectedBody, tables, fields,maths, true,selfTest,&result.warnings);
        if (selfTest) {
            const auto pdf = candidate + L".pdf";
            previewOwned = true;
            if (taskShared) { TaskHandle preview(CreateFileW(pdf.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr)); require(preview.value != INVALID_HANDLE_VALUE, "Reserve test preview"); taskTrackFile(pdf, preview.value, false); }
            invoke(word.document.Get(), L"ExportAsFixedFormat", DISPATCH_METHOD, {pdf.c_str(), 17L});
        }
        std::puts("UI_PHASE CLOSING");
        // 实验项启用时先在已核验的只读正文副本上判断是否真的存在实验目标；没有就不再打开第二个 Word 文档。
        if(experimentalOptions.enabled()){
            std::puts("UI_EXPERIMENTAL_PHASE PREPARING");
            bool present=true;
            try{present=experimentalTargetsPresent(word.document.Get(),candidate,experimentalOptions,reverse,scanOptions);}
            catch(const std::exception& error){std::printf("EXPERIMENTAL_PRESENCE_FALLBACK %s\n",error.what());}
            if(!present){experimentalSkip=L"NO_TARGETS";std::puts("EXPERIMENTAL_NO_TARGETS");std::puts("UI_EXPERIMENTAL_SKIPPED NO_TARGETS");}
        }
        if(experimentalOptions.enabled()&&experimentalSkip.empty()){invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();}
        else{word.shutdown();if(taskShared)InterlockedExchange(&taskShared->mathReady,0);if(api.initialized){require(api.term()==0,"Terminate MathType API");api.initialized=false;}}
        std::puts("UI_PHASE PACKAGE_VERIFY");
        afterPackage = inspectPackage(candidate,"SAVED_PARTS",false,scanOptions.progress);
        std::puts("UI_PHASE PACKAGE_COMPARE");
        requireContains(afterPackage.otherEmbeddings,originalPackage.otherEmbeddings,"Unconverted embedded packages stay byte-identical");
        if (reverse) {
            auto retained = originalPackage;
            const std::set<std::vector<unsigned char>> knownNative(originalPackage.equations.begin(),originalPackage.equations.end()),
                knownMedia(originalPackage.media.begin(),originalPackage.media.end());
            for (const auto& outcome : outcomes) if (outcome.status == L"SUCCEEDED" && !outcome.native.empty()) {
                // 多个逻辑对象可共用一个包部件；核对来源，但每个物理资产至多扣到零。
                require(knownNative.count(outcome.native)!=0,"Exported native stream belonged to original package");
                const auto native = std::find(retained.equations.begin(), retained.equations.end(), outcome.native);
                if(native!=retained.equations.end())retained.equations.erase(native);
                for (const auto& preview : outcome.removedPreviews) {
                    require(knownMedia.count(preview)!=0,"Removed equation preview belonged to original package");
                    const auto media = std::find(retained.media.begin(), retained.media.end(), preview);
                    if(media!=retained.media.end())retained.media.erase(media);
                }
            }
            // Word 可保留共享部件或为剩余对象复制它；逻辑对象数已逐式核验，包资产只允许原字节。
            require(std::all_of(afterPackage.equations.begin(),afterPackage.equations.end(),[&](const auto& value){return knownNative.count(value)!=0;}),"Saved reverse native streams come from original package");
            require(std::all_of(afterPackage.media.begin(),afterPackage.media.end(),[&](const auto& value){return knownMedia.count(value)!=0;}),"Saved reverse media come from original package");
            requireContains(afterPackage.equations, retained.equations, "All retained native streams stay byte-identical");
            requireContains(afterPackage.media, retained.media, "Ordinary media and retained previews stay byte-identical");
        } else {
            require(afterPackage.equations.size() == originalPackage.equations.size() + succeeded, "Saved native stream increment equals successful targets");
            requireContains(afterPackage.equations, originalPackage.equations, "Every existing Equation Native stream stays byte-identical");
            requireContains(afterPackage.media, originalPackage.media, "Every existing image and equation preview stays byte-identical");
            // 直接读取磁盘上的所有 MTEF，精确核对新公式内容与样式，正常任务无需再次激活 OLE。
            std::vector<std::vector<unsigned char>> savedNative,expectedNative;
            for(const auto& equation:afterPackage.equations)savedNative.push_back(nativePayload(equation));
            for(const auto& equation:originalPackage.equations)expectedNative.push_back(nativePayload(equation));
            for(const auto& outcome:outcomes)if(outcome.status==L"SUCCEEDED")expectedNative.push_back(outcome.native);
            requireContains(savedNative,expectedNative,"Every saved MTEF payload matches original or converted bytes including duplicates");
        }
        }else{
            // 快速模式保存一次后直接交付，不重开、不重扫、不核对包资产。
            word.shutdown();if(taskShared)InterlockedExchange(&taskShared->mathReady,0);
            if(api.initialized){require(api.term()==0,"Terminate MathType API");api.initialized=false;}
            if(taskShared)InterlockedExchange(&taskShared->wordQuitComplete,1);
        }
        break;
        }catch(const BodyRangeError& error){
            require(!isStopping(),"TASK_CANCELLED");const auto& item=error.item;
            require(std::none_of(skippedBody.begin(),skippedBody.end(),[&](const ScanItem& prior){return prior.span.start==item.span.start&&prior.source==item.source;}),"Failed body range is not retried");
            skippedBody.push_back(item);const std::string message=error.what();recoveryIssues.push_back({L"body:"+std::to_wstring(item.span.start),L"正文公式",L"RECOVERY",std::wstring(message.begin(),message.end()),item.position,item.source});emitProcessingIssue(recoveryIssues.back());
            numbers.reset();
            if(word.document){invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();}
            require(CopyFileW(input.c_str(),candidate.c_str(),FALSE)!=0,"Rebuild owned body copy from locked original");setDocxLinkUpdates(candidate);
            taskTestBoundary(L"BODY_COPY_REBUILT");
            word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{candidate.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));
            std::puts("BODY_CHECKPOINT_RESTORED");std::puts("UI_BODY_RESET");
            for(const auto& issue:recoveryIssues)emitProcessingIssue(issue);
        }}
        result.experimentalSkip=experimentalSkip;result.numbering=numbers.summary;result.savedNative=batchSavedNative;
        return result;
    }
};
}
BodyResult convertBody(WordSession& word,MathTypeApi& api,const std::wstring& input,const std::wstring& candidate,
    const std::wstring& style,const PackageState& originalPackage,const BodyOptions& options,
    const ExperimentalOptions& experimentalOptions,const std::wstring& sizePreferencePath,bool& sizePreferenceOwned,bool& previewOwned,ConversionTimings& timings) {
    InterlockedExchange(&stopRequested,0);
    return BodyConversion(options,timings).run(word,api,input,candidate,style,originalPackage,experimentalOptions,sizePreferencePath,sizePreferenceOwned,previewOwned);
}
