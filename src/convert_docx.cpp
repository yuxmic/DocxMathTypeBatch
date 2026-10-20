// 阶段 3：只写新副本，重扫后倒序转换；逐条 Undo 恢复，保存重开通过后发布。
#include "docx_scan.h"
#include "mathtype_api.h"
#include "equation_objects.h"
#include "experimental_conversion.h"
#include "app_version.h"
#include <msopc.h>
#include <xmllite.h>
#include <wincrypt.h>
#include <cstring>
#include <map>
#include <set>
#include "native_numbering.h"

static volatile LONG stopRequested{};
static bool addRightNumbers{};
static long insertedNumberFields{};
static long updatedNumberFields{};
static long updatedNumberReferences{};
static bool batchSavedNative{};
static long addedNativeFields{};
static std::vector<std::wstring> numberBookmarks;
static std::wstring numberIdentity;
static EquationSpacing inlineSpacing,displaySpacing;
static EquationNumbering numbering;
static BOOL WINAPI requestStop(DWORD event) {
    if (event != CTRL_C_EVENT && event != CTRL_BREAK_EVENT) return FALSE;
    InterlockedExchange(&stopRequested, 1); return TRUE;
}

static double number(IDispatch* object, const wchar_t* property) {
    auto value = invoke(object, property, DISPATCH_PROPERTYGET);
    if (value.v.vt == VT_R4) return value.v.fltVal;
    if (value.v.vt == VT_R8) return value.v.dblVal;
    if (value.v.vt == VT_I4) return value.v.lVal;
    throw std::runtime_error("Automation numeric type");
}
static bool boolProperty(IDispatch* object, const wchar_t* property) {
    auto value = invoke(object, property, DISPATCH_PROPERTYGET);
    require(value.v.vt == VT_BOOL, "Automation Boolean type"); return value.v.boolVal != VARIANT_FALSE;
}
static void requireStyle(const EquationStyle& actual, const EquationStyle& expected) {
    require(actual.inlineEquation == expected.inlineEquation && actual.sizes == expected.sizes && actual.styles == expected.styles,
            "New equation native style matches task EQP");
}

// 使用厂商转换器在内存中应用样式，不设置共享服务的 NEXT_EQN 偏好。
static std::vector<unsigned char> styledNative(MathTypeApi& api,std::vector<unsigned char> source,const std::string& prefs,bool display) {
    using Reset=long (WINAPI*)();using Set=long (WINAPI*)(short,const char*);
    using Transform=long (WINAPI*)(short,short,const unsigned char*,long,short,short,unsigned char*,long,const char*,void*);
    using Status=long (WINAPI*)(short);
    const auto reset=reinterpret_cast<Reset>(GetProcAddress(api.library,"MTXFormReset"));
    const auto set=reinterpret_cast<Set>(GetProcAddress(api.library,"MTXFormSetPrefs"));
    const auto transform=reinterpret_cast<Transform>(GetProcAddress(api.library,"MTXFormEqn"));
    const auto status=reinterpret_cast<Status>(GetProcAddress(api.library,"MTXFormGetStatus"));
    require(reset&&set&&transform&&status,"Native preference transform exports");
    size_t options=5;while(source.at(options++)!=0){}source.at(options)=display?0:1;
    struct Dimensions {short baseline;RECT bounds;} dimensions{};
    require(reset()==0&&set(3,prefs.c_str())==0,"Select task transform preferences");
    unsigned char first{};const auto query=transform(-3,4,source.data(),static_cast<long>(source.size()),-3,4,&first,1,"",&dimensions);
    const long length=status(-1);require((query==-9||query==0)&&length>0&&length<1000000,"Query styled native length");
    std::vector<unsigned char> result(static_cast<size_t>(length));
    require(transform(-3,4,source.data(),static_cast<long>(source.size()),-3,4,result.data(),length,"",&dimensions)==0,"Transform task native preferences");
    require(reset()==0,"Restore native transform options");return result;
}

// 仅主字号变化；同字号只准备一次，厂商字体、相对字号和间距原样保留。
static std::string sizedPreferences(MathTypeApi& api,double points,const std::wstring& path,bool& owned) {
    require(std::isfinite(points)&&points>=1&&points<=256,"Source equation size must be 1..256 pt");
    FILE* raw{};require(_wfopen_s(&raw,(api.install+L"\\Preferences\\Times+Symbol 12.eqp").c_str(),L"rb")==0&&raw,"Read default EQP");
    std::unique_ptr<FILE,decltype(&fclose)> file(raw,&fclose);std::string prefs;char buffer[4096];
    for(size_t count;(count=fread(buffer,1,sizeof(buffer),file.get()))!=0;)prefs.append(buffer,count);
    require(!ferror(file.get()),"Read complete default EQP");
    const auto first=prefs.find("Full="),last=prefs.find_first_of("\r\n",first);require(first!=std::string::npos,"Default Full size key");
    char value[48]{};sprintf_s(value,"Full=%.2f pt",points);prefs.replace(first,last==std::string::npos?std::string::npos:last-first,value);
    if(!owned)taskReservePrivatePath(path);
    TaskHandle output(CreateFileW(path.c_str(),GENERIC_WRITE,FILE_SHARE_READ,nullptr,owned?TRUNCATE_EXISTING:CREATE_NEW,FILE_ATTRIBUTE_TEMPORARY,nullptr));
    require(output.value!=INVALID_HANDLE_VALUE,"Create private size EQP");owned=true;DWORD written{};
    require(prefs.size()<=MAXDWORD&&WriteFile(output.value,prefs.data(),static_cast<DWORD>(prefs.size()),&written,nullptr)&&written==prefs.size()&&FlushFileBuffers(output.value),"Write private size EQP");
    CloseHandle(output.value);output.value=nullptr;return loadPrefs(api,path);
}

// 临时空对象只通过厂商书签插入，不触发交互式编辑窗口。
static std::vector<unsigned char> emptyEquation(MathTypeApi& api, WordSession& word,
                                               const std::string* prefs = nullptr, bool display = true) {
    require(word.owned, "Own scratch Word session");
    auto scratch = asObject(invoke(get(word.application.Get(), L"Documents").Get(), L"Add", DISPATCH_METHOD, {Value(), false, Value(), false}));
    ComPtr<IDispatch> object;
    bool active = false;
    try {
        // 从厂商空公式书签插入 OLE，避免 Word.AddOLEObject 自动执行交互式 Open。
        auto blankRange=docRange(scratch.Get(),0,0);
        auto shape=insertBlankEquation(scratch.Get(),blankRange.Get(),api.install);
        // 本机厂商 Verb 2 为 RunForConversion；保留其转换服务初始化，不调用 Open/Edit。
        auto ole = get(shape.Get(), L"OLEFormat"); invoke(ole.Get(), L"DoVerb", DISPATCH_METHOD, {2L});
        object = get(ole.Get(), L"Object"); active = true;
        if(prefs){const auto styled=styledNative(api,readNative(api,object.Get()),*prefs,display);
            require(api.writeBytes(object.Get(),3,styled.data(),static_cast<long>(styled.size()))==0,"Apply task native preferences");}
        auto native = readNative(api, object.Get());
        require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close empty equation"); active = false;
        invoke(scratch.Get(), L"Close", DISPATCH_METHOD, {0L});
        invoke(word.document.Get(), L"Activate", DISPATCH_METHOD); // 自定义 Undo 必须绑定任务文档。
        return native;
    } catch (...) {
        if (active) require(api.close(OLECLOSE_NOSAVE, object.Get()) == 0, "Close failed empty equation");
        invoke(scratch.Get(), L"Close", DISPATCH_METHOD, {0L});
        throw;
    }
}
struct ItemOutcome {
    Span finalSpan;
    std::wstring status, reason;
    std::vector<unsigned char> native;
    long addedParagraphs{};
    long shapeOrdinal{};
    std::wstring outputText;
    std::vector<std::vector<unsigned char>> removedPreviews;
    long nativeOrdinal{};
    Span replacementSpan;
};
struct BodyRangeError : std::runtime_error {
    ScanItem item;
    BodyRangeError(const ScanItem& source,const std::exception& error):std::runtime_error(error.what()),item(source){}
};
static ItemOutcome failedPreserved(const ScanItem& item,const std::exception& error) {
    const std::string message=error.what();std::printf("ITEM_PREPARATION_PRESERVED %s\n",message.c_str());
    return {item.span,L"FAILED_PRESERVED",std::wstring(message.begin(),message.end())};
}
// 反向清单仍只覆盖主故事正文/表格；边界对象保留，不激活任何既有 OLE。
ScanResult scanReverseDocument(IDispatch* document) {
    // 反向按原生类型及导出的定界符核对，保留旧文档中独占段落的行内对象。
    auto result = scanDocumentText(document, false); result.reverse = true;
    if (result.protectedDocument) return result;
    if(guiStageProgress)std::puts(checkingSavedDocument?"UI_PHASE RESCANNING_OBJECTS":"UI_PHASE SCANNING_OBJECTS");
    auto content = get(document, L"Content"); auto shapes = get(content.Get(), L"InlineShapes");
    std::vector<Span> excluded;
    for (const auto name : {L"Revisions", L"ContentControls", L"OMaths"}) {
        auto collection = get(content.Get(), name);
        for (long index = 1; index <= integer(collection.Get(), L"Count"); ++index)
            excluded.push_back(bounds(get(collectionItem(collection.Get(), index).Get(), L"Range").Get()));
    }
    auto fields = get(content.Get(), L"Fields");
    for (long index = 1; index <= integer(fields.Get(), L"Count"); ++index) {
        auto field = collectionItem(fields.Get(), index); if (integer(field.Get(), L"Type") == 58) continue;
        excluded.push_back(fieldBounds(field.Get()));
    }
    const long shapeCount=integer(shapes.Get(),L"Count");
    StageProgress objectProgress(checkingSavedDocument?"RESCAN_OBJECTS":"SCAN_OBJECTS",static_cast<size_t>(shapeCount));
    for (long index = 1; index <= shapeCount; ++index) {
        auto shape = collectionItem(shapes.Get(), index); const auto type = integer(shape.Get(), L"Type");
        if (type != 1 && type != 2) {objectProgress.advance();continue;}
        const auto cls = text(get(shape.Get(), L"OLEFormat").Get(), L"ClassType");
        if (cls.find(L"Equation.DSMT") != 0) {objectProgress.advance();continue;}
        auto range = get(shape.Get(), L"Range"); ScanItem item;
        item.span = bounds(range.Get()); item.source = text(range.Get(), L"Text"); item.delimiter = L"MATH_TYPE";
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
        if(item.ready&&!selectedEquation(item.display)){objectProgress.advance();continue;}
        result.items.push_back(std::move(item));
        objectProgress.advance();
    }
    std::sort(result.items.begin(), result.items.end(), [](const ScanItem& a, const ScanItem& b) {return a.span.start < b.span.start;});
    long end = -1;
    for (const auto& item : result.items) {require(item.span.start >= end, "Non-overlapping reverse targets"); end = item.span.end;}
    return result;
}
static std::vector<std::wstring> nativeMathSnapshot(IDispatch* document,bool contentOnly=false) {
    std::vector<std::wstring> snapshot; auto maths = get(get(document, L"Content").Get(), L"OMaths");
    for (long index = 1; index <= integer(maths.Get(), L"Count"); ++index)
        snapshot.push_back(contentOnly?ommlSource(get(collectionItem(maths.Get(),index).Get(),L"Range").Get()):sourceSignature(get(collectionItem(maths.Get(), index).Get(), L"Range").Get()));
    return snapshot;
}
// 新导出只接收已经验收的 ASCII 单段 TeX；其他情况保持原对象并报告原因。
static ItemOutcome reverseOne(MathTypeApi& api, WordSession& word, const ScanItem& item, bool dollar,
                              bool selfTest, bool abortTest) {
    auto document = word.document.Get();ComPtr<IDispatch> range,shapes,undo;long count{},end{};
    std::wstring body,sourceXml,prefix,suffix;
    const bool object = item.delimiter == L"MATH_TYPE";
    try{
    range = docRange(document, item.span.start, item.span.end);
    require(text(range.Get(), L"Text") == item.source, "Reverse source rechecked before export");
    shapes = get(get(document, L"Content").Get(), L"InlineShapes");count = integer(shapes.Get(), L"Count");
    end = integer(get(document, L"Content").Get(), L"End");
    body = text(get(document, L"Content").Get(), L"Text");sourceXml = sourceSignature(range.Get(), object);
    prefix = text(docRange(document, 0, item.span.start).Get(), L"Text");suffix = text(docRange(document, item.span.end, end).Get(), L"Text");
    undo = get(word.application.Get(), L"UndoRecord");
    }catch(const std::exception& error){return failedPreserved(item,error);}
    NativeAsset asset;
    bool recordOpen = false;
    try {
        // 现有 LaTeX 只改定界符，保留原文中的 Word 换行及空白。
        const size_t delimiterLength=item.delimiter==L"$"?1:2;
        std::wstring inner = object?item.inner:item.source.substr(delimiterLength,item.source.size()-2*delimiterLength);
        if (object) {
            asset = rangeNativeAsset(range.Get()); const auto payload = nativePayload(asset.equation);
            require(inspectStyle(payload).inlineEquation == !item.display, "Reverse native type still matches scan");
            inner = exportedInner(exportNativeTex(api, payload), item.display);
        }
        const std::wstring open = item.display ? (dollar ? L"$$" : L"\\[") : L"$";
        const std::wstring close = item.display ? (dollar ? L"$$" : L"\\]") : L"$";
        const auto output = open + inner + close;
        invoke(document, L"Activate", DISPATCH_METHOD);
        invoke(undo.Get(), L"StartCustomRecord", DISPATCH_METHOD, {L"MathType 导出/定界符归一"}); recordOpen = true;
        // 已有 LaTeX 只改两端，保留内部原文和字符格式；MathType 成功导出后才替换对象。
        if (object) invoke(range.Get(), L"Text", DISPATCH_PROPERTYPUT, {output.c_str()});
        else if (output != item.source) {
            const long oldDelimiter = item.delimiter == L"$" ? 1 : 2;
            invoke(docRange(document, item.span.end - oldDelimiter, item.span.end).Get(), L"Text", DISPATCH_PROPERTYPUT, {close.c_str()});
            invoke(docRange(document, item.span.start, item.span.start + oldDelimiter).Get(), L"Text", DISPATCH_PROPERTYPUT, {open.c_str()});
        }
        require(!(selfTest && object && item.display), "Injected reverse replacement failure");
        require(!conversionFault(L"DOCX_MATH_BODY_FAULT",std::to_wstring(item.paragraph),L"RECOVERY"),"Injected reverse body replacement failure");
        const long afterEnd = integer(get(document, L"Content").Get(), L"End"), replacementEnd = item.span.end + afterEnd - end;
        require(text(docRange(document, 0, item.span.start).Get(), L"Text") == prefix &&
                text(docRange(document, replacementEnd, afterEnd).Get(), L"Text") == suffix &&
                text(docRange(document, item.span.start, replacementEnd).Get(), L"Text") == output &&
                integer(shapes.Get(), L"Count") == count - static_cast<long>(object), "Reverse replaces only the target range");
        if (item.row) {
            auto replaced = docRange(document, item.span.start, replacementEnd);
            require(information(replaced.Get(), 12) && information(replaced.Get(), 13) == item.row &&
                    information(replaced.Get(), 16) == item.column, "Reverse target remains in original cell");
        }
        invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD); recordOpen = false;
        ItemOutcome outcome{{item.span.start, replacementEnd}, L"SUCCEEDED", L"", std::move(asset.equation)};
        outcome.outputText = output; outcome.removedPreviews = std::move(asset.previews); return outcome;
    } catch (const std::exception& error) {
        if (recordOpen) invoke(undo.Get(), L"EndCustomRecord", DISPATCH_METHOD);
        const bool changed = text(get(document, L"Content").Get(), L"Text") != body || integer(shapes.Get(), L"Count") != count ||
                             sourceSignature(docRange(document, item.span.start, item.span.end).Get(), object) != sourceXml;
        if (changed) {
            const auto undone = invoke(document, L"Undo", DISPATCH_METHOD, {1L});
            require(undone.v.vt == VT_BOOL && undone.v.boolVal != VARIANT_FALSE, "Undo failed reverse replacement");
        }
        const auto recoveredXml = sourceSignature(docRange(document, item.span.start, item.span.end).Get(), object);
        require(text(get(document, L"Content").Get(), L"Text") == body && integer(shapes.Get(), L"Count") == count &&
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
static std::vector<long> tableStructure(IDispatch* document);
static std::vector<std::wstring> ordinaryFields(IDispatch* document,const std::vector<ScannedField>* cached=nullptr);
static void validateReverse(IDispatch* document, const ScanResult& before, std::vector<ItemOutcome>& outcomes,
                            const std::wstring& body, const std::vector<long>& tables,
                            const std::vector<std::wstring>& fields, const std::vector<std::wstring>& maths) {
    require(text(get(document, L"Content").Get(), L"Text") == body && tableStructure(document) == tables &&
            ordinaryFields(document) == fields && nativeMathSnapshot(document) == maths, "Reverse preserves body, table identities, fields and full OMML");
    const auto after = scanReverseDocument(document); long removed{};
    for (size_t index = 0; index < outcomes.size(); ++index)
        removed += before.items[index].delimiter == L"MATH_TYPE" && outcomes[index].status == L"SUCCEEDED";
    require(after.items.size() == before.items.size() && after.inlineShapes == before.inlineShapes - removed &&
            after.existingMathType == before.existingMathType - removed && after.nativeMath == before.nativeMath &&
            after.fields == before.fields - removed && after.revisions == before.revisions && after.controls == before.controls &&
            after.noteReferences == before.noteReferences, "Reverse target/OLE/EMBED and excluded counts");
    // 保存重开可能改变域坐标，用同一反向重扫的有序清单重新绑定真实位置。
    StageProgress itemProgress("VERIFY_ITEMS",outcomes.size());
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
// MathType 不支持 aligned；只改导入副本，并按顶层分列保留多组等式。
// 原生显示布局按文档初始化一次；宽度只在实际节变化时重新读取，不加载交互宏。
struct DisplayLayout {
    IDispatch* document{};bool stylePresent{},singleSection{};ComPtr<IDispatch> style;std::map<long,double> widths;std::map<std::wstring,ComPtr<IDispatch>> formats;
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
            double width=number(page.Get(),L"PageWidth")-number(page.Get(),L"LeftMargin")-number(page.Get(),L"RightMargin");
            auto columns=get(page.Get(),L"TextColumns");if(integer(columns.Get(),L"Count")>1)width=number(collectionItem(columns.Get(),1).Get(),L"Width");widths[section]=width;
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

static ItemOutcome convertOne(MathTypeApi& api, WordSession& word, const ScanItem& item,
                              const std::vector<unsigned char>& empty,DisplayLayout& displayLayout,ComPtr<IXMLDOMDocument2>& stylesheet,bool selfTest, bool abortTest) {
    auto document = word.document.Get();
    ComPtr<IDispatch> range,shapes,undo;long count{},end{};std::wstring body,sourceXml,sourcePackage,prefix,suffix;
    try{
    range = docRange(document, item.span.start, item.span.end);
    require(text(range.Get(), L"Text") == item.source, "Source rechecked immediately before conversion");
    shapes = get(get(document, L"Content").Get(), L"InlineShapes");
    count = integer(shapes.Get(), L"Count");end = integer(get(document, L"Content").Get(), L"End");
    body = selfTest ? text(get(document, L"Content").Get(), L"Text") : std::wstring();
    sourceXml = sourceSignature(range.Get());
    sourcePackage=item.ommlOrdinal?text(range.Get(),L"WordOpenXML"):std::wstring();
    prefix = selfTest ? text(docRange(document, 0, item.span.start).Get(), L"Text") : std::wstring();
    suffix = selfTest ? text(docRange(document, item.span.end, end).Get(), L"Text") : std::wstring();
    undo = get(word.application.Get(), L"UndoRecord");
    }catch(const std::exception& error){return failedPreserved(item,error);}
    bool active = false, recordOpen = false;
    ComPtr<IDispatch> object;
    try {
        invoke(document, L"Activate", DISPATCH_METHOD);
        invoke(undo.Get(), L"StartCustomRecord", DISPATCH_METHOD, {L"正文公式单条转换"});recordOpen=true;
        // ponytail: 本机 TeX Input 对中文实测乱码；可靠 Unicode 导入通过验收后再放开非 ASCII。
        const bool omml=item.ommlOrdinal!=0;
        require(omml||std::all_of(item.inner.begin(), item.inner.end(), [](wchar_t c) {return c <= 127;}), "UNSUPPORTED_NON_ASCII_TEX");
        require(!(selfTest && item.source == L"$beta_2$"), "Injected preparation failure");
        const auto inner=omml?std::wstring():compatibleTex(item.inner);
        const auto input=omml?ommlMathML(word,range.Get(),stylesheet,item.display):(item.display?L"\\["+inner+L"\\]":L"$"+inner+L"$");
        // 是否拆段按当前段落重算：同段后面的公式可能已经拆段，不能沿用扫描提示。
        auto paragraph = collectionItem(get(range.Get(), L"Paragraphs").Get(), 1);
        auto paragraphs=get(range.Get(),L"Paragraphs");
        auto last=collectionItem(paragraphs.Get(),integer(paragraphs.Get(),L"Count"));
        const Span para{integer(get(paragraph.Get(),L"Range").Get(),L"Start"),integer(get(last.Get(),L"Range").Get(),L"End")};
        const auto leading = item.display ? text(docRange(document, para.start, item.span.start).Get(), L"Text") : std::wstring();
        const bool splitBefore = item.display && !whitespaceOnly(leading);
        const bool splitAfter = item.display && !whitespaceOnly(text(docRange(document, item.span.end, para.end).Get(), L"Text"));
        const bool leadingTab=item.display;
        // 两个 Tab 与源码一次替换；对象插在中间，段尾预留右对齐位置供原生按钮插入编号。
        // OMML 先替换为对象，离开原生公式区域后再补 Tab，避免破坏可撤销的公式结构。
        invoke(range.Get(), L"Text", DISPATCH_PROPERTYPUT, {item.display&&!omml?L"\t\t":L""});
        if(item.display&&!omml)invoke(range.Get(),L"SetRange",DISPATCH_METHOD,{item.span.start+1,item.span.start+1});else invoke(range.Get(),L"Collapse",DISPATCH_METHOD,{1L});
        require(!(selfTest && item.source == L"$split_1+split_2$"), "Injected creation failure");
        auto shape=insertBlankEquation(document,range.Get(),api.install);
        auto ole=get(shape.Get(),L"OLEFormat");invoke(ole.Get(),L"DoVerb",DISPATCH_METHOD,{2L});
        object=get(ole.Get(),L"Object");active=true;
        require(api.writeBytes(object.Get(),3,empty.data(),static_cast<long>(empty.size()))==0,"Apply isolated equation preferences");
        taskTestBoundary(L"NATIVE_STYLE_WRITTEN");
        require(input.size()<=LONG_MAX&&api.write(object.Get(),omml?2:1,input.c_str(),static_cast<long>(input.size()))==0,omml?"Import OMML into new object":"Write TeX into new object");
        auto native=readNative(api,object.Get());require(native!=empty&&native.size()>empty.size(),"New equation contains native content");
        requireStyle(inspectStyle(native),inspectStyle(empty));
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
            requireDisplayLayout(document, shapeRange.Get());
        }
        const auto& spacing=item.display?displaySpacing:inlineSpacing;
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
        require((!selfTest || (text(docRange(document, 0, item.span.start).Get(), L"Text") == prefix &&
                text(docRange(document, replacementEnd, afterEnd).Get(), L"Text") == suffix)) &&
                text(docRange(document, item.span.start, replacementEnd).Get(), L"Text") == replaced &&
                integer(shapes.Get(), L"Count") == count + 1, "Only source text replaced by one OLE and required paragraph breaks");
        require(number(shape.Get(), L"Width") > 0 && number(shape.Get(), L"Height") > 0, "New equation dimensions");
        if (item.row) require(information(shapeRange.Get(), 12) && information(shapeRange.Get(), 13) == item.row &&
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
            sourceSignature(docRange(document, item.span.start, item.span.end).Get()) != sourceXml;
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
static bool displayOutcome(const ItemOutcome& outcome){return outcome.status==L"SUCCEEDED"||outcome.status==L"EXISTING_FORMATTED";}
static bool nativeOutcome(const ItemOutcome& outcome){return displayOutcome(outcome)||outcome.status==L"EXISTING_PRESERVED";}
static bool changedOutcome(const ItemOutcome& outcome){return displayOutcome(outcome)||outcome.status==L"REFERENCE_UPDATED";}
static void includeNumberReferences(IDispatch* document,ScanResult& scan){
    if(!std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.existingDisplay;}))return;
    ComPtr<IDispatch> bookmarks;std::vector<ScanItem> references;
    for(const auto& field:scan.fieldSnapshots){if(field.kind!=3)continue;
        const auto first=field.code.find_first_not_of(L" \t\r\n");if(first==std::wstring::npos||field.code.compare(first,4,L"REF "))continue;
        const auto start=field.code.find_first_not_of(L" \t\r\n",first+4);if(start==std::wstring::npos)continue;const auto end=field.code.find_first_of(L" \t\r\n",start);
        auto name=field.code.substr(start,end==std::wstring::npos?end:end-start);if(name.size()>1&&name.front()==L'"'&&name.back()==L'"')name=name.substr(1,name.size()-2);
        if(!bookmarks)bookmarks=get(document,L"Bookmarks");const auto exists=invoke(bookmarks.Get(),L"Exists",DISPATCH_METHOD,{name.c_str()});if(exists.v.vt!=VT_BOOL||exists.v.boolVal==VARIANT_FALSE)continue;
        const auto mark=asObject(invoke(bookmarks.Get(),L"Item",DISPATCH_METHOD,{name.c_str()}));const auto span=bounds(get(mark.Get(),L"Range").Get());
        const bool target=std::any_of(scan.items.begin(),scan.items.end(),[&](const ScanItem& item){return item.existingDisplay&&span.start>=item.span.start&&span.end<=item.span.end;});
        const bool number=std::any_of(scan.fieldSnapshots.begin(),scan.fieldSnapshots.end(),[&](const ScannedField& value){return value.kind==51&&value.code.find(L"MACROBUTTON MTPlaceRef")!=std::wstring::npos&&span.start>=value.span.start&&span.end<=value.span.end;});
        if(target&&number){ScanItem item;item.span=field.span;item.source=text(docRange(document,field.span.start,field.span.end).Get(),L"Text");item.existingReference=true;item.delimiter=L"EXISTING_REFERENCE";references.push_back(std::move(item));}
    }
    scan.items.insert(scan.items.end(),references.begin(),references.end());std::sort(scan.items.begin(),scan.items.end(),[](const ScanItem& a,const ScanItem& b){return a.span.start<b.span.start;});
    for(size_t i=1;i<scan.items.size();++i)require(scan.items[i-1].span.end<=scan.items[i].span.start,"Number reference changes do not overlap conversion targets");
}
static void rememberNumberRange(IDispatch* document,Span span){
    auto range=docRange(document,span.start,span.end);const auto name=numberIdentity+L"_"+std::to_wstring(numberBookmarks.size());
    invoke(get(document,L"Bookmarks").Get(),L"Add",DISPATCH_METHOD,{name.c_str(),range.Get()});numberBookmarks.push_back(name);
}
static ItemOutcome formatExistingDisplay(WordSession& word,const ScanItem& item,DisplayLayout& layout,const std::vector<std::vector<unsigned char>>& assets){
    auto document=word.document.Get();auto previousLayout=layout;
    auto range=docRange(document,item.span.start,item.span.end);require(text(range.Get(),L"Text")==item.source,"Existing display source remains exact before formatting");
    auto shape=collectionItem(get(range.Get(),L"InlineShapes").Get(),1);auto objectRange=get(shape.Get(),L"Range");const auto before=integer(get(document,L"Content").Get(),L"End");
    const auto sourceXml=sourceSignature(range.Get(),true);const auto native=assets.empty()?nativePayload(rangeNativeAsset(objectRange.Get()).equation):assets.at(static_cast<size_t>(item.existingNativeOrdinal-1));
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
        if(integer(get(document,L"Content").Get(),L"End")!=before||sourceSignature(docRange(document,item.span.start,item.span.end).Get(),true)!=sourceXml){
            auto undone=invoke(document,L"Undo",DISPATCH_METHOD,{1L});require(undone.v.vt==VT_BOOL&&undone.v.boolVal,"Undo failed existing display formatting");}
        require(integer(get(document,L"Content").Get(),L"End")==before&&sourceSignature(docRange(document,item.span.start,item.span.end).Get(),true)==sourceXml,"Existing display recovery restores original paragraph");
        auto original=collectionItem(get(docRange(document,item.span.start,item.span.end).Get(),L"InlineShapes").Get(),1);auto originalRange=get(original.Get(),L"Range");
        require(nativePayload(rangeNativeAsset(originalRange.Get()).equation)==native,"Existing display recovery preserves original native contents");
        layout=std::move(previousLayout);const std::string message=error.what();ItemOutcome outcome{bounds(originalRange.Get()),L"EXISTING_PRESERVED",std::wstring(message.begin(),message.end()),native};outcome.replacementSpan=item.span;return outcome;
    }
}
static std::vector<std::wstring> ordinaryFields(IDispatch* document,const std::vector<ScannedField>* cached) {
    std::vector<std::wstring> snapshot;std::vector<Span> owned;
    for(const auto& name:numberBookmarks){auto mark=asObject(invoke(get(document,L"Bookmarks").Get(),L"Item",DISPATCH_METHOD,{name.c_str()}));owned.push_back(bounds(get(mark.Get(),L"Range").Get()));}
    auto fields=cached?ComPtr<IDispatch>():get(get(document,L"Content").Get(),L"Fields");
    const long count=cached?static_cast<long>(cached->size()):integer(fields.Get(),L"Count");
    for(long index=0;index<count;++index){
        ScannedField read;const auto& value=cached?cached->at(static_cast<size_t>(index)):read;ComPtr<IDispatch> field,code;
        if(!cached){field=collectionItem(fields.Get(),index+1);read.kind=integer(field.Get(),L"Type");}
        if(value.kind==58)continue;
        if(!cached){code=get(field.Get(),L"Code");if(!owned.empty())read.codeSpan=bounds(code.Get());}
        if(std::any_of(owned.begin(),owned.end(),[&](const Span& range){return range.start<=value.codeSpan.start&&range.end>=value.codeSpan.end;}))continue;
        if(!cached){read.code=text(code.Get(),L"Text");read.result=text(get(field.Get(),L"Result").Get(),L"Text");}
        snapshot.push_back(std::to_wstring(value.kind));snapshot.push_back(value.code);snapshot.push_back(value.result);
    }return snapshot;
}
// 创建兼容字段而不加载交互宏；所有外层括号和内层序列都归 MathType 管理。
static void insertRightNumbers(IDispatch* document,const ScanResult& scan,std::vector<ItemOutcome>& outcomes) {
    std::vector<ComPtr<IDispatch>> native(outcomes.size());size_t first=outcomes.size(),count{};
    for(size_t i=0;i<outcomes.size();++i)if(scan.items[i].display&&displayOutcome(outcomes[i])){first=std::min(first,i);++count;}
    if(!count)return;
    const long originalFields=integer(get(get(document,L"Content").Get(),L"Fields").Get(),L"Count");
    nativeProperty(document,L"MTEquationNumber2",numbering.format);nativeProperty(document,L"MTCustomEquationNumber",numbering.custom?L"1":L"0");nativeProperty(document,L"MTEquationSection",L"1");
    auto remember=[&](IDispatch* field){rememberNumberRange(document,nativeFieldSpan(field));};
    StageProgress progress("NUMBER_EQUATIONS",count);
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
                if(kind==51&&code.find(L"MACROBUTTON MTPlaceRef")!=std::wstring::npos){require(!native[i],"One native number per display paragraph");native[i]=field;}
                if(kind==12&&code.find(L"\\r")!=std::wstring::npos){const auto span=nativeFieldSpan(field.Get());if(native[i]){const auto outer=nativeFieldSpan(native[i].Get());if(span.start>=outer.start&&span.end<=outer.end)continue;}for(const auto* name:{L"MTEqn",L"MTChap",L"MTSec"})if(code.find(L"SEQ "+std::wstring(name)+L" ")!=std::wstring::npos)seeds[name]=field;}
            }
            // 旧起始值域不再重置中途序列；首个目标的隐藏种子域负责显式起始值。
            const auto seed=seeds.find(L"MTEqn");if(seed!=seeds.end()){auto code=get(seed->second.Get(),L"Code");invoke(code.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"SEQ MTEqn \\c \\h"});invoke(seed->second.Get(),L"Update",DISPATCH_METHOD);}
        }
        const long end=integer(get(shape.Get(),L"Range").Get(),L"End");
        if(!native[i]){native[i]=nativeNumber(document,end+1,numbering);++insertedNumberFields;}
        else{auto gap=docRange(document,end,nativeFieldSpan(native[i].Get()).start);if(seeds.empty()&&text(gap.Get(),L"Text")!=L"\t")invoke(gap.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"\t"});++updatedNumberFields;}
        remember(native[i].Get());
        if(i==first){
            // 只为显式高级设置添加空结果序列域，不创建 MathType 章节断点或可见说明文字。
            std::vector<std::pair<std::wstring,long>> counters;
            // 厂商格式化会重建编号内部域；显式起始值放在外部空结果域，下一编号自然递增。
            if(numbering.start!=1)counters.emplace_back(L"MTEqn",numbering.start-1);
            if(numbering.format.find(L"#C")!=std::wstring::npos)counters.emplace_back(L"MTChap",numbering.chapter);
            if(numbering.format.find(L"#S")!=std::wstring::npos)counters.emplace_back(L"MTSec",numbering.section);
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
        const long delta=integer(get(document,L"Content").Get(),L"End")-before;outcome.replacementSpan.end+=delta;outcome.status=L"REFERENCE_UPDATED";++updatedNumberReferences;
        for(size_t later=i+1;later<outcomes.size();++later){outcomes[later].finalSpan.start+=delta;outcomes[later].finalSpan.end+=delta;outcomes[later].replacementSpan.start+=delta;outcomes[later].replacementSpan.end+=delta;}
    }
    for(size_t i=0;i<outcomes.size();++i)if(native[i]||scan.items[i].existingReference)outcomes[i].outputText=text(docRange(document,outcomes[i].replacementSpan.start,outcomes[i].replacementSpan.end).Get(),L"Text");
    addedNativeFields=integer(get(get(document,L"Content").Get(),L"Fields").Get(),L"Count")-originalFields;
}
static void bindShapeOrdinals(IDispatch* document,std::vector<ItemOutcome>& outcomes) {
    auto shapes=get(get(document,L"Content").Get(),L"InlineShapes");
    std::map<std::pair<long,long>,long> ordinals;
    std::map<std::pair<long,long>,long> nativeOrdinals;long nativeOrdinal{};
    const long count=integer(shapes.Get(),L"Count");
    StageProgress objectProgress("BIND_OBJECTS",static_cast<size_t>(count));
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
static void validateBeforeSave(IDispatch* document,const ScanResult& before,const std::vector<ItemOutcome>& outcomes,
                               const std::wstring& body,const std::vector<long>& tables,
                               const std::vector<std::wstring>& fields,const std::vector<std::wstring>& maths) {
    require(text(get(document,L"Content").Get(),L"Text")==body && tableStructure(document)==tables && ordinaryFields(document)==fields,
            "Conversion preserves body, table identities and ordinary fields");
    long delta{},removedMath{};
    for(size_t index=0;index<outcomes.size();++index) if(outcomes[index].status==L"SUCCEEDED")
        {delta+=before.reverse ? (before.items[index].delimiter==L"MATH_TYPE"?-1:0) : 1;removedMath+=before.items[index].ommlOrdinal!=0;}
    auto content=get(document,L"Content");
    require(integer(get(content.Get(),L"InlineShapes").Get(),L"Count")==before.inlineShapes+delta &&
            integer(get(content.Get(),L"Fields").Get(),L"Count")==before.fields+delta+addedNativeFields &&
            integer(get(content.Get(),L"OMaths").Get(),L"Count")==before.nativeMath-removedMath &&
            integer(get(content.Get(),L"Revisions").Get(),L"Count")==before.revisions &&
            integer(get(content.Get(),L"ContentControls").Get(),L"Count")==before.controls &&
            integer(get(document,L"Footnotes").Get(),L"Count")+integer(get(document,L"Endnotes").Get(),L"Count")==before.noteReferences,
            "Conversion object and excluded boundary counts");
    if(before.reverse)require(nativeMathSnapshot(document)==maths,"Conversion preserves full OMML");
    if(before.bodyOmml)require(nativeMathSnapshot(document,true)==maths,"Unconverted OMML contents and order preserved");
}
static void validateDocument(MathTypeApi& api, IDispatch* document, const ScanResult& before,
                             std::vector<ItemOutcome>& outcomes, const std::wstring& expectedBody,
                             const std::vector<long>& tables, const std::vector<std::wstring>& fields,const std::vector<std::wstring>& maths, bool reopen, bool readback=true) {
    std::puts("TIMING_STAGE SAVED_STRUCTURE");
    require(text(get(document, L"Content").Get(), L"Text") == expectedBody, "Saved and reopened body equals verified conversion body");
    require(tableStructure(document) == tables, "Table cells and row/column identities unchanged");
    std::puts("TIMING_STAGE SAVED_SCAN");
    const auto after = scanDocumentText(document,true,true,false,before.bodyOmml);
    require(ordinaryFields(document,&after.fieldSnapshots) == fields, "Existing non-embedded field code and results unchanged");
    const auto succeeded = std::count_if(outcomes.begin(), outcomes.end(), [](const ItemOutcome& item) { return item.status == L"SUCCEEDED"; });
    long removedMath{};for(size_t i=0;i<outcomes.size();++i)removedMath+=before.items[i].ommlOrdinal&&outcomes[i].status==L"SUCCEEDED";
    require(after.inlineShapes == before.inlineShapes + succeeded && after.existingMathType == before.existingMathType + succeeded &&
            after.nativeMath == before.nativeMath-removedMath && after.fields == before.fields + succeeded + addedNativeFields && after.revisions == before.revisions &&
            after.controls == before.controls && after.noteReferences == before.noteReferences, "Object and excluded boundary counts unchanged");
    if(before.bodyOmml)require(nativeMathSnapshot(document,true)==maths,"Saved unconverted OMML contents and order preserved");
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
    StageProgress itemProgress("VERIFY_ITEMS",outcomes.size());
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
        if (item.display&&!addRightNumbers&&!item.existingDisplay) requireDisplayLayout(document, shapeRange.Get());
        if (item.row) require(information(shapeRange.Get(), 12) && information(shapeRange.Get(), 13) == item.row &&
                              information(shapeRange.Get(), 16) == item.column, "Saved new OLE remains in original cell");
        require(number(shape.Get(), L"Width") > 0 && number(shape.Get(), L"Height") > 0, "Saved OLE dimensions");
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
        require(text(get(document, L"Content").Get(), L"Text") == expectedBody && ordinaryFields(document) == fields,
                "Native verification leaves ordinary text and fields unchanged");
    }
}
static void writeNewText(const std::wstring& path, const std::string& contents) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(file != INVALID_HANDLE_VALUE, "Create new batch report"); DWORD written{};
    taskTrackFile(path, file, false);
    const bool saved = contents.size() <= MAXDWORD && WriteFile(file, contents.data(), static_cast<DWORD>(contents.size()), &written, nullptr) && written == contents.size();
    const bool flushed = FlushFileBuffers(file) != 0, closed = CloseHandle(file) != 0;
    if (!saved || !flushed || !closed) { DeleteFileW(path.c_str()); throw std::runtime_error("Write complete batch report"); }
}
static std::string outcomeReport(const std::wstring& input, const std::wstring& output, const std::wstring& eqp,
                                 const ScanResult& scan, const std::vector<ItemOutcome>& outcomes, bool fullReadback,const ExperimentalResult& experimental,const std::vector<ProcessingIssue>& issues) {
    size_t ready{}, succeeded{}, failed{}, unattempted{};
    for (size_t index = 0; index < outcomes.size(); ++index) {
        ready += scan.items[index].ready; succeeded += outcomes[index].status == L"SUCCEEDED";
        failed += outcomes[index].status == L"FAILED_PRESERVED"; unattempted += outcomes[index].status == L"NOT_ATTEMPTED";
    }
    const auto existing=std::count_if(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.existingDisplay||item.existingReference;});
    const size_t rejected=outcomes.size()-ready-static_cast<size_t>(existing)+experimental.rejected;
    ready+=experimental.succeeded+experimental.failed+experimental.unattempted;succeeded+=experimental.succeeded;failed+=experimental.failed;unattempted+=experimental.unattempted;
    const auto status = failed || rejected || unattempted || !issues.empty() || experimental.partial() ? L"PARTIAL" : L"COMPLETE";
    std::string json = "{\n  \"SchemaVersion\":1, \"Application\":\"DocxMathTypeBatch\", \"ApplicationVersion\":\"" DMTB_VERSION "\", \"CliContractVersion\":" DMTB_CLI_CONTRACT ", \"Scope\":"+quote(experimental.enabled?L"BODY_WITH_EXPERIMENTAL_EXTENSIONS":L"MAIN_STORY_BODY_ONLY")+", \"Positions\":\"WORD_SNAPSHOTS_BEFORE_AND_AFTER\",\n";
    json += "  \"EquationMode\":"+quote(equationScope==EquationScope::Inline?L"Inline":equationScope==EquationScope::Display?L"Display":L"All")+", \"NumbersAdded\":"+std::to_string(insertedNumberFields)+",\n";
    const auto displays=std::count_if(outcomes.begin(),outcomes.end(),[](const ItemOutcome& item){return item.status==L"EXISTING_FORMATTED";});
    json += "  \"NumbersUpdated\":"+std::to_string(updatedNumberFields)+", \"NumberReferencesUpdated\":"+std::to_string(updatedNumberReferences)+", \"ExistingDisplaysFormatted\":"+std::to_string(displays)+",\n";
    json += "  \"InlineSpacing\":"+quote(spacingArgument(inlineSpacing))+", \"DisplaySpacing\":"+quote(spacingArgument(displaySpacing))+", \"NumberFormat\":"+quote(numbering.format)+",\n";
    json += "  \"SavedNativeBatch\":"+std::string(batchSavedNative?"true":"false")+",\n";
    json += "  \"Input\":" + quote(input) + ", \"Output\":" + quote(output) + ", \"Preferences\":" + quote(eqp) + ",\n";
    if (scan.reverse) {
        long exported{}, normalized{};
        for (size_t index = 0; index < outcomes.size(); ++index) if (outcomes[index].status == L"SUCCEEDED") {
            if (scan.items[index].delimiter == L"MATH_TYPE") ++exported; else ++normalized;
        }
        json += "  \"Direction\":\"REVERSE\", \"DisplayDelimiter\":" + quote(eqp) +
                ", \"ExportedMathType\":" + std::to_string(exported) + ", \"NormalizedLatex\":" + std::to_string(normalized) + ",\n";
    }
    json += "  \"Status\":" + quote(status) + ", \"Targets\":" + std::to_string(ready) + ", \"Succeeded\":" + std::to_string(succeeded) +
            ", \"FailedPreserved\":" + std::to_string(failed) + ", \"NotAttempted\":" + std::to_string(unattempted) +
            ", \"RejectedPreserved\":" + std::to_string(rejected) + ", \"ExistingMathType\":" + std::to_string(scan.existingMathType) + ",\n";
    json += "  \"Validation\":\"" + std::string(scan.reverse ? "NATIVE_TRANSLATOR_CONTENT_RESIDUALS_OMML_PACKAGE_AND_WORD_REOPEN" : "SAVED_NATIVE_BYTES_STYLE_RESIDUALS_PACKAGE_AND_WORD_REOPEN") + "\", \"MathematicalEquivalenceVerified\":false,\n";
    json += "  \"FullDevelopmentReadback\":" + std::string(fullReadback?"true":"false") + ", \"ReopenedEquationActivations\":" + std::to_string(fullReadback&&!scan.reverse?succeeded-experimental.succeeded:0) + ",\n";
    json += "  \"PostConversionScans\":"+std::to_string((fullReadback?2:1)+(experimental.enabled&&experimental.skipped.empty()?2:0))+",\n";
    json += scan.reverse ? "  \"OriginalFileBytesUnchanged\":true, \"RetainedNativeStreamsUnchanged\":true, \"RetainedMediaBytesPreserved\":true, \"OMMLUnchanged\":true, \"WordReopenVerified\":true,\n  \"Items\":[\n" :
                          "  \"OriginalFileBytesUnchanged\":true, \"ExistingNativeStreamsUnchanged\":true, \"ExistingMediaBytesPreserved\":true, \"WordReopenVerified\":true,\n  \"Items\":[\n";
    size_t written{};for (size_t index = 0; index < outcomes.size(); ++index) {
        const auto& item = scan.items[index]; const auto& result = outcomes[index];
        if(item.existingDisplay||item.existingReference)continue;
        const auto fullSize=result.native.empty()||scan.reverse?std::string():inspectStyle(result.native).sizes.at(0).second;
        json += (written++?",\n":"")+std::string("    {\"Id\":") + std::to_string(index + 1) + ",\"InitialStart\":" + std::to_string(item.span.start) + ",\"InitialEnd\":" + std::to_string(item.span.end) +
            ",\"FinalStart\":" + std::to_string(result.finalSpan.start) + ",\"FinalEnd\":" + std::to_string(result.finalSpan.end) +
            ",\"Region\":\"" + (item.row ? "TABLE" : "BODY") + std::string("\",\"Row\":") + std::to_string(item.row) + ",\"Column\":" + std::to_string(item.column) +
            ",\"Display\":" + (item.display ? "true" : "false") + ",\"Status\":" + quote(result.status) + ",\"Reason\":" + quote(result.reason) +
            ",\"NativeFullSize\":" + (fullSize.empty()?std::string("null"):quote(std::wstring(fullSize.begin(),fullSize.end()))) +
            ",\"AddedParagraphs\":" + std::to_string(result.addedParagraphs) + ",\"Source\":" + quote(item.source) +
            (scan.reverse ? ",\"Kind\":" + quote(item.delimiter == L"MATH_TYPE" ? L"MATH_TYPE" : L"LATEX") + ",\"OutputText\":" + quote(result.outputText) : ",\"Kind\":"+quote(item.ommlOrdinal?L"OMML":L"LATEX")+",\"FormattingApplied\":"+(result.status==L"SUCCEEDED"?"true":"false")) + "}";
    }
    for(const auto& item:experimental.items)json+=(written++?",\n":"")+item.json();
    json+="\n  ],\n  \"Issues\":[";size_t issueIndex{};for(const auto& issue:issues)json+=(issueIndex++?",":"")+issue.json();for(const auto& issue:experimental.issues)json+=(issueIndex++?",":"")+issue.json();
    return json + "],\n  \"Experimental\":"+experimental.json()+"\n}\n";
}

int convertMain(int argc, wchar_t** argv) {
    setvbuf(stdout, nullptr, _IONBF, 0);
    const std::wstring mode = argc > 1 ? argv[1] : L"";
    const bool reverse = mode == L"--reverse" || mode == L"--worker-reverse" || mode == L"--self-test-reverse" ||
                         mode == L"--self-test-reverse-stop" || mode == L"--self-test-reverse-abort";
    const bool guiWorker = mode == L"--worker-convert" || mode == L"--worker-reverse";
    guiStageProgress=guiWorker;checkingSavedDocument=false;
    if ((guiWorker ? argc < 7 : argc < 6) ||
        (mode != L"--convert" && !reverse && !guiWorker && mode != L"--self-test" && mode != L"--self-test-abort" && mode != L"--self-test-stop")) {
        std::puts("usage: docx-math-convert --convert <input.docx> <new-output.docx> <new-report.json> <paragraph|size=pt|task.eqp> [MathType-directory] [--scope=all|inline|display] [--numbers]"); return 2;
    }
    const bool selfTest = mode == L"--self-test" || mode == L"--self-test-abort" || mode == L"--self-test-reverse" || mode == L"--self-test-reverse-abort";
    if (selfTest) {
        require(compatibleTex(L"\\begin{aligned}a&=b\\\\c&=d\\end{aligned}") == L"\\begin{array}{rl}a&=b\\\\c&=d\\end{array}", "Aligned two columns");
        require(compatibleTex(L"\\begin{aligned}[t]a&=b,&c&=d\\end{aligned}") == L"\\begin{array}{rlrl}a&=b,&c&=d\\end{array}", "Aligned four columns and position");
        require(compatibleTex(L"\\begin{aligned}a&=\\begin{matrix}x&y\\\\z&w\\end{matrix}\\end{aligned}") == L"\\begin{array}{rl}a&=\\begin{matrix}x&y\\\\z&w\\end{matrix}\\end{array}", "Nested matrix columns stay separate");
        require(compatibleTex(L"\\text{aligned} a\\&b") == L"\\text{aligned} a\\&b", "Ordinary TeX stays exact");
        require(parseSpacing(L"multiple:1.25").value==1.25&&parseSpacing(L"exact:24").mode==4,"Spacing modes and values");
        EquationNumbering testNumber;testNumber.format=L"[#CI.#Sa-#EA]";testNumber.chapter=4;testNumber.section=2;
        const auto parts=numberParts(testNumber,28);require(parts.before==L"[IV.b-"&&parts.digits==L"BB"&&parts.after==L"]","Number format tokens and alphabetic boundary");
    }
    const bool abortTest = mode == L"--self-test-abort" || mode == L"--self-test-reverse-abort";
    const bool stopTest = mode == L"--self-test-stop" || mode == L"--self-test-reverse-stop";
    const auto initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED); if (FAILED(initialized)) return 1;
    std::wstring candidate, stagedReport, output, sizePreferencePath;
    bool sizePreferenceOwned{};
    bool candidateOwned = false, reportOwned = false, outputOwned = false, previewOwned = false, publishedPreviewOwned = false;
    HANDLE sourceLock = INVALID_HANDLE_VALUE;
    int code = 1;
    try {
        std::puts("UI_PHASE PREPARING");
        std::wstring installProblem;
        equationScope=EquationScope::All;addRightNumbers=false;insertedNumberFields=updatedNumberFields=updatedNumberReferences=0;batchSavedNative=false;
        inlineSpacing={};displaySpacing={};numbering={};addedNativeFields=0;numberBookmarks.clear();
        std::wstring explicitInstall;ExperimentalOptions experimentalOptions;
        for(int i=guiWorker?7:6;i<argc;++i) {
            const std::wstring argument=argv[i];
            if(argument==L"--scope=all")equationScope=EquationScope::All;
            else if(argument==L"--experimental-tables")experimentalOptions.tables=true;
            else if(argument==L"--experimental-textboxes")experimentalOptions.textboxes=true;
            else if(argument==L"--experimental-headers")experimentalOptions.headers=true;
            else if(argument==L"--experimental-notes")experimentalOptions.notes=true;
            else if(argument==L"--scope=inline")equationScope=EquationScope::Inline;
            else if(argument==L"--scope=display")equationScope=EquationScope::Display;
            else if(argument==L"--numbers")addRightNumbers=true;
            else if(argument==L"--number-custom")numbering.custom=true;
            else if(argument.rfind(L"--inline-spacing=",0)==0)inlineSpacing=parseSpacing(argument.substr(17));
            else if(argument.rfind(L"--display-spacing=",0)==0)displaySpacing=parseSpacing(argument.substr(18));
            else if(argument.rfind(L"--number-format=",0)==0)numbering.format=argument.substr(16);
            else if(argument.rfind(L"--number-start=",0)==0){const auto value=settingNumber(argument.substr(15),1,999999);require(value==std::floor(value),"Integer start required");numbering.start=static_cast<long>(value);}
            else if(argument.rfind(L"--number-chapter=",0)==0){const auto value=settingNumber(argument.substr(17),1,999999);require(value==std::floor(value),"Integer chapter required");numbering.chapter=static_cast<long>(value);}
            else if(argument.rfind(L"--number-section=",0)==0){const auto value=settingNumber(argument.substr(17),1,999999);require(value==std::floor(value),"Integer section required");numbering.section=static_cast<long>(value);}
            else {require(argument.rfind(L"--",0)!=0&&explicitInstall.empty(),"Unknown or duplicate conversion argument");explicitInstall=absolutePath(argv[i]);}
        }
        require(!reverse||!addRightNumbers,"Reverse cannot add equation numbers");
        require(!reverse||(!inlineSpacing.mode&&!displaySpacing.mode),"Reverse preserves paragraph spacing");
        if(addRightNumbers)numberParts(numbering,numbering.start);
        numberIdentity=addRightNumbers?L"DMTBNum_"+std::to_wstring(GetCurrentProcessId())+L"_"+std::to_wstring(GetTickCount64()):L"";
        const auto install = explicitInstall.empty()?findMathTypeInstall(installProblem):explicitInstall;
        if (!install.empty()) installProblem=mathTypeInstallProblem(install);
        require(installProblem.empty(),utf8(installProblem).c_str());
        std::printf("MATHTYPE_DIRECTORY %s\n",wireHex(utf8(install)).c_str());
        const auto stopFile = guiWorker ? absolutePath(argv[6]) : L"";
        // 无控制台 GUI 通过自己的任务标记请求停止，仍只在公式完成/恢复边界检查。
        auto isStopping = [&]() {
            if (!stopFile.empty() && GetFileAttributesW(stopFile.c_str()) != INVALID_FILE_ATTRIBUTES) InterlockedExchange(&stopRequested, 1);
            return taskCancelled() || InterlockedCompareExchange(&stopRequested, 0, 0) != 0;
        };
        const auto input = absolutePath(argv[2]), report = absolutePath(argv[4]);
        const std::wstring styleArgument=argv[5];
        const bool followSize=!reverse&&styleArgument==L"paragraph", fixedSize=!reverse&&styleArgument.rfind(L"size=",0)==0;
        double requestedPoints{};
        if(fixedSize){wchar_t* end{};requestedPoints=wcstod(styleArgument.c_str()+5,&end);require(end!=styleArgument.c_str()+5&&!*end&&std::isfinite(requestedPoints)&&requestedPoints>=1&&requestedPoints<=256,"Specified size must be 1..256 pt");}
        const auto eqp = reverse||followSize||fixedSize?styleArgument:absolutePath(argv[5]);
        require(!reverse || eqp == L"bracket" || eqp == L"dollar", "Reverse display delimiter must be bracket or dollar");
        output = absolutePath(argv[3]);
        require(input.size() > 5 && _wcsicmp(input.c_str() + input.size() - 5, L".docx") == 0 &&
                output.size() > 5 && _wcsicmp(output.c_str() + output.size() - 5, L".docx") == 0, "DOCX input/output required");
        require(_wcsicmp(input.c_str(), output.c_str()) != 0 && _wcsicmp(output.c_str(), report.c_str()) != 0, "Distinct input/output/report paths required");
        require(GetFileAttributesW(output.c_str()) == INVALID_FILE_ATTRIBUTES, "Output already exists");
        require(GetFileAttributesW(report.c_str()) == INVALID_FILE_ATTRIBUTES, "Report already exists");
        if (selfTest) require(GetFileAttributesW((output + L".pdf").c_str()) == INVALID_FILE_ATTRIBUTES, "Test preview already exists");
        requireDocxZip(input);
        if (!reverse&&!followSize&&!fixedSize) require(GetFileAttributesW(eqp.c_str()) != INVALID_FILE_ATTRIBUTES, "Task EQP required");
        sourceLock = CreateFileW(input.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if(sourceLock==INVALID_HANDLE_VALUE)throw std::runtime_error("Lock source against concurrent writes during copy and conversion (Win32="+std::to_string(GetLastError())+")");
        const auto originalPackage = inspectPackage(input,"INPUT_PARTS");
        // 文件字节只读对比；输入同时保持只读共享锁，不能在扫描后被其他程序改写。
        auto readOriginal = [&](const std::vector<char>* expected=nullptr) {
            // 复用始终持有的原件句柄；末尾仍逐字节核对，只省去重新打开、重复扩容和第二份整文件缓冲。
            LARGE_INTEGER size{},zero{};require(GetFileSizeEx(sourceLock,&size)&&SetFilePointerEx(sourceLock,zero,nullptr,FILE_BEGIN),"Read locked original size and rewind");
            if(expected)require(static_cast<size_t>(size.QuadPart)==expected->size(),"Original DOCX size unchanged");
            std::vector<char> bytes(expected?0:static_cast<size_t>(size.QuadPart));char buffer[64*1024];
            for(size_t offset=0;offset<static_cast<size_t>(size.QuadPart);){
                const auto count=static_cast<DWORD>(std::min(static_cast<size_t>(size.QuadPart)-offset,sizeof(buffer)));DWORD read{};
                require(ReadFile(sourceLock,expected?buffer:bytes.data()+offset,count,&read,nullptr)&&read==count,"Complete original DOCX read");
                if(expected)require(std::equal(buffer,buffer+read,expected->begin()+offset),"Original DOCX bytes unchanged");offset+=read;
            }
            return bytes;
        };
        std::puts("TIMING_STAGE SOURCE_SNAPSHOT");
        const auto originalBytes = readOriginal();
        std::puts("TIMING_STAGE COPY_AND_LINKS");
        const auto suffix = L".work-" + std::to_wstring(GetCurrentProcessId()) + L"-" + std::to_wstring(GetTickCount64());
        candidate = output + suffix + L".docx"; stagedReport = report + suffix + L".tmp";
        sizePreferencePath=candidate+L".eqp";
        taskCopy(input, candidate); candidateOwned = true;
        const auto linkUpdates=setDocxLinkUpdates(candidate);
        std::puts("UI_PHASE OPENING");
        WordSession word;
        auto documents = get(word.application.Get(), L"Documents");
        word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD, {Value(), false, Value(), false})); word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
        invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
        invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L}); word.document.Reset();
        taskTrackWordOwner(candidate);
        word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD,
            {candidate.c_str(), false, false, false, Value(), Value(), false, Value(), Value(), Value(), Value(), false}));
        require(!boolProperty(word.document.Get(), L"ReadOnly"), "Task copy must be writable");
        require(!boolProperty(word.document.Get(), L"TrackRevisions"), "Disable revision tracking in source before batch conversion");
        MathTypeApi api(install);ScanResult scan;std::vector<ItemOutcome> outcomes;size_t total{},completed{},succeeded{};
        std::vector<ScanItem> skippedBody;std::vector<ProcessingIssue> recoveryIssues,issues;std::wstring experimentalSkip;PackageState afterPackage;
        for(;;){try{
        // 原件始终持有只读锁；只有单式恢复无法确认时才重建副本并跳过该式，正常路径不新增磁盘备份。
        issues=recoveryIssues;numberBookmarks.clear();insertedNumberFields=updatedNumberFields=updatedNumberReferences=addedNativeFields=0;
        std::puts("UI_PHASE SCANNING");
        scan = reverse ? scanReverseDocument(word.document.Get()) : scanDocumentText(word.document.Get(),true,false,true,true);
        if(addRightNumbers)includeNumberReferences(word.document.Get(),scan);
        require(!scan.protectedDocument, "Protected document cannot be converted");
        if (guiWorker) emitGuiScan(scan);
        // 识别拒绝项在初始化接口前发送，后续整份任务失败时也能显示逐式中文原因。
        if(guiWorker)for(size_t index=0;index<scan.items.size();++index){const auto& item=scan.items[index];
            if(!item.ready&&!item.existingDisplay&&!item.existingReference)std::printf("UI_ITEM_REJECTED %zu %ld %ld %ld %s %s\n",index+1,item.paragraph,item.span.start,item.span.end,diagnosticExcerpt(item.source).c_str(),wireHex(utf8(item.reason)).c_str());
        }
        std::puts("UI_PHASE INITIALIZING");
        if(addRightNumbers)for(const auto& item:scan.items)if(item.existingDisplay||item.existingReference)rememberNumberRange(word.document.Get(),item.span);
        const auto tables = tableStructure(word.document.Get()); const auto fields = ordinaryFields(word.document.Get());
        auto maths = reverse||scan.bodyOmml ? nativeMathSnapshot(word.document.Get(),!reverse) : std::vector<std::wstring>();
        const bool needsMathType=std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.ready;});
        if(needsMathType&&!api.initialized){require(api.init(0,10)>0,"MathType initialization");api.initialized=true;taskMathTypeConnected();}
        const auto prefs = reverse||followSize||fixedSize||!needsMathType ? std::string() : loadPrefs(api, eqp), globalPrefs = needsMathType?readDefaultPrefs(api):std::string();
        std::vector<unsigned char> naturalEmpty;
        if(!reverse&&std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.ready;}))naturalEmpty=emptyEquation(api,word);
        std::map<double,std::string> sizePrefs;
        std::map<std::pair<double,bool>,std::vector<unsigned char>> emptyStyles;
        // 正文按目标之间的不变片段只读取一次，独立构造预期结果，避免逐式读取全文。
        std::vector<std::wstring> bodyGaps;long previousEnd{};
        for(const auto& item:scan.items){bodyGaps.push_back(text(docRange(word.document.Get(),previousEnd,item.span.start).Get(),L"Text"));previousEnd=item.span.end;}
        const long contentEnd=integer(get(word.document.Get(),L"Content").Get(),L"End");
        // 末段公式覆盖到正文末尾时，没有尾片段；Word 不接受在排他末端创建折叠 Range。
        bodyGaps.push_back(previousEnd==contentEnd?std::wstring():text(docRange(word.document.Get(),previousEnd,contentEnd).Get(),L"Text"));
        std::wstring originalBody;
        for(size_t i=0;i<scan.items.size();++i)originalBody+=bodyGaps[i]+scan.items[i].source;
        originalBody+=bodyGaps.back();
        require(originalBody==text(get(word.document.Get(),L"Content").Get(),L"Text"),"Independent body snapshot uses exact Word text ranges");
        outcomes.clear();
        for (const auto& item : scan.items){outcomes.push_back({item.span, item.ready ? L"NOT_ATTEMPTED" : item.existingDisplay?L"EXISTING_PENDING":item.existingReference?L"REFERENCE_PRESERVED":L"REJECTED_PRESERVED", item.reason, {}, 0});outcomes.back().replacementSpan=item.span;}
        const bool existingDisplays=std::any_of(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return item.existingDisplay;});
        const auto existingAssets=existingDisplays?bodyNativeAssets(word.document.Get()):std::vector<std::vector<unsigned char>>();
        DisplayLayout displayLayout{word.document.Get(),originalPackage.displayStyle};
        ComPtr<IXMLDOMDocument2> ommlStylesheet;
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
            if(guiWorker)std::printf("UI_ITEM_BEGIN %zu %ld %ld %ld %s\n",index+1,item.paragraph,item.span.start,item.span.end,diagnosticExcerpt(item.source).c_str());
            long beforeEnd{};
            try{
            beforeEnd=integer(get(word.document.Get(), L"Content").Get(), L"End");
            if(item.existingDisplay){outcomes[index]=formatExistingDisplay(word,item,displayLayout,existingAssets);
                if(outcomes[index].status==L"EXISTING_PRESERVED")issues.push_back({L"body:"+std::to_wstring(index+1),L"正文第 "+std::to_wstring(item.paragraph)+L" 段",L"FORMATTING",outcomes[index].reason});}
            else if(reverse)outcomes[index]=reverseOne(api,word,item,eqp==L"dollar",selfTest,abortTest);
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
            }catch(const std::exception& error){if(abortTest)throw;require(!isStopping(),"TASK_CANCELLED");if(needsMathType)require(readDefaultPrefs(api)==globalPrefs,"Global MathType defaults unchanged");throw BodyRangeError(item,error);}
            if(!displayOutcome(outcomes[index]))outcomes[index].replacementSpan=outcomes[index].finalSpan;
            const long delta = integer(get(word.document.Get(), L"Content").Get(), L"End") - beforeEnd;
            // 前面的替换只移动已经处理过的后方坐标；当前及更前方目标仍使用刚重扫的原位置。
            for (size_t later = index + 1; later < outcomes.size(); ++later) {
                outcomes[later].finalSpan.start += delta; outcomes[later].finalSpan.end += delta;
                outcomes[later].replacementSpan.start+=delta;outcomes[later].replacementSpan.end+=delta;
            }
            if(item.existingDisplay){if(guiWorker)std::puts("UI_ITEM_END");continue;}
            std::printf("CONVERT_PROGRESS %zu/%zu ID=%zu START=%ld STATUS=%s\n", ++completed, total, index + 1, item.span.start,
                outcomes[index].status == L"SUCCEEDED" ? "SUCCEEDED" : "FAILED_PRESERVED");
            std::printf("UI_ITEM_RESULT %zu %s %s\n", index + 1, outcomes[index].status == L"SUCCEEDED" ? "SUCCEEDED" : "FAILED_PRESERVED",
                        wireHex(utf8(outcomes[index].reason)).c_str());
            if (stopTest && completed == 3) {InterlockedExchange(&stopRequested, 1);if(taskShared)InterlockedExchange(&taskShared->cancelRequested,1);} // 自检同样丢弃结果。
        }
        if (!guiWorker && !taskShared) require(SetConsoleCtrlHandler(requestStop, FALSE) != 0, "Remove safe stop handler");
        require(!isStopping(),"TASK_CANCELLED");
        // 成功的正文 OMML 从保留清单移除，其余公式仍逐式核对内容和顺序。
        if(scan.bodyOmml)for(size_t i=scan.items.size();i-- >0;)if(scan.items[i].ommlOrdinal&&outcomes[i].status==L"SUCCEEDED")maths.erase(maths.begin()+scan.items[i].ommlOrdinal-1);
        auto expectedText=[&](){std::wstring body;for(size_t i=0;i<scan.items.size();++i)body+=bodyGaps[i]+(changedOutcome(outcomes[i])?outcomes[i].outputText:scan.items[i].source);return body+bodyGaps.back();};
        std::puts("TIMING_STAGE NUMBERING");
        if(addRightNumbers&&std::any_of(outcomes.begin(),outcomes.end(),displayOutcome)){
            // 连续编号和 REF 必须一起成功；只在启用编号时保存一个阶段检查点。
            validateBeforeSave(word.document.Get(),scan,outcomes,expectedText(),tables,fields,maths);
            bindShapeOrdinals(word.document.Get(),outcomes);invoke(word.document.Get(),L"Save",DISPATCH_METHOD);
            ConversionCheckpoint checkpoint(candidate,L".numbering-checkpoint");const auto beforeNumbering=outcomes;const auto bookmarksBefore=numberBookmarks;const auto bodyBefore=expectedText();
            try{insertRightNumbers(word.document.Get(),scan,outcomes);}
            catch(const std::exception& error){require(!isStopping(),"TASK_CANCELLED");const std::string reason=error.what();
                displayLayout.style.Reset();displayLayout.formats.clear();invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();checkpoint.restore(candidate);
                outcomes=beforeNumbering;numberBookmarks=bookmarksBefore;insertedNumberFields=updatedNumberFields=updatedNumberReferences=addedNativeFields=0;
                word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{candidate.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));
                displayLayout.document=word.document.Get();checkingSavedDocument=true;validateDocument(api,word.document.Get(),scan,outcomes,bodyBefore,tables,fields,maths,true,false);checkingSavedDocument=false;
                issues.push_back({L"numbering",L"正文编号及交叉引用",L"NUMBERING",std::wstring(reason.begin(),reason.end())});std::puts("NUMBERING_CHECKPOINT_RESTORED");
            }
        }
        displayLayout.style.Reset();
        displayLayout.formats.clear();
        const auto expectedBody=expectedText();
        std::puts("UI_PHASE VERIFYING");
        if(selfTest) {
            if(reverse)validateReverse(word.document.Get(),scan,outcomes,expectedBody,tables,fields,maths);
            else validateDocument(api,word.document.Get(),scan,outcomes,expectedBody,tables,fields,maths,false);
        } else {
            // 单条转换已核对内容与恢复；完整重扫只在保存重开后执行一次。
            validateBeforeSave(word.document.Get(),scan,outcomes,expectedBody,tables,fields,maths);
            std::puts("TIMING_STAGE BIND_OBJECTS");
            if(!reverse)bindShapeOrdinals(word.document.Get(),outcomes);
        }
        if(needsMathType)require(readDefaultPrefs(api) == globalPrefs, "Global MathType defaults unchanged");
        std::puts("UI_PHASE SAVING");
        taskTestBoundary(L"BEFORE_SAVE");
        invoke(word.document.Get(), L"Save", DISPATCH_METHOD);
        invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L}); word.document.Reset();
        taskTestBoundary(L"SAVED_DOCUMENT");
        std::puts("UI_PHASE REOPENING");
        word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD,
            {candidate.c_str(), false, true, false, Value(), Value(), false, Value(), Value(), Value(), Value(), false}));
        require(boolProperty(word.document.Get(), L"ReadOnly"), "Verification opens candidate read-only");
        checkingSavedDocument=true;std::puts("UI_PHASE RESCANNING");
        if (reverse) validateReverse(word.document.Get(), scan, outcomes, expectedBody, tables, fields, maths);
        else validateDocument(api, word.document.Get(), scan, outcomes, expectedBody, tables, fields,maths, true,selfTest);
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
            try{present=experimentalTargetsPresent(word.document.Get(),candidate,experimentalOptions,reverse);}
            catch(const std::exception& error){std::printf("EXPERIMENTAL_PRESENCE_FALLBACK %s\n",error.what());}
            if(!present){experimentalSkip=L"NO_TARGETS";std::puts("EXPERIMENTAL_NO_TARGETS");std::puts("UI_EXPERIMENTAL_SKIPPED NO_TARGETS");}
        }
        if(experimentalOptions.enabled()&&experimentalSkip.empty()){invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();}
        else{word.shutdown();if(taskShared)InterlockedExchange(&taskShared->mathReady,0);if(api.initialized){require(api.term()==0,"Terminate MathType API");api.initialized=false;}}
        std::puts("UI_PHASE PACKAGE_VERIFY");
        afterPackage = inspectPackage(candidate,"SAVED_PARTS");
        std::puts("UI_PHASE PACKAGE_COMPARE");
        succeeded = static_cast<size_t>(std::count_if(outcomes.begin(), outcomes.end(), [](const ItemOutcome& item) { return item.status == L"SUCCEEDED"; }));
        if (reverse) {
            auto retained = originalPackage;
            for (const auto& outcome : outcomes) if (outcome.status == L"SUCCEEDED" && !outcome.native.empty()) {
                const auto native = std::find(retained.equations.begin(), retained.equations.end(), outcome.native);
                require(native != retained.equations.end(), "Exported native stream belonged to original package"); retained.equations.erase(native);
                for (const auto& preview : outcome.removedPreviews) {
                    const auto media = std::find(retained.media.begin(), retained.media.end(), preview);
                    require(media != retained.media.end(), "Removed equation preview belonged to original package"); retained.media.erase(media);
                }
            }
            require(afterPackage.equations.size() == retained.equations.size(), "Saved reverse native stream decrement equals exported objects");
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
            requireContains(std::move(savedNative),expectedNative,"Every saved MTEF payload matches original or converted bytes including duplicates");
        }
        break;
        }catch(const BodyRangeError& error){
            require(!isStopping(),"TASK_CANCELLED");const auto& item=error.item;
            require(std::none_of(skippedBody.begin(),skippedBody.end(),[&](const ScanItem& prior){return prior.span.start==item.span.start&&prior.source==item.source;}),"Failed body range is not retried");
            skippedBody.push_back(item);const std::string message=error.what();recoveryIssues.push_back({L"body:"+std::to_wstring(item.span.start),L"正文第 "+std::to_wstring(item.paragraph)+L" 段",L"RECOVERY",std::wstring(message.begin(),message.end())});
            if(word.document){invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();}
            require(CopyFileW(input.c_str(),candidate.c_str(),FALSE)!=0,"Rebuild owned body copy from locked original");setDocxLinkUpdates(candidate);
            taskTestBoundary(L"BODY_COPY_REBUILT");
            word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{candidate.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));
            std::puts("BODY_CHECKPOINT_RESTORED");std::puts("UI_BODY_RESET");
        }}
        ExperimentalResult experimental;experimental.options=experimentalOptions;experimental.enabled=experimentalOptions.enabled();
        if(experimentalOptions.enabled()){
            checkingSavedDocument=false;
            if(experimentalSkip.empty()){
                // 正文阶段已核验的输出包直接作为实验基线，避免实验分支再整包读取一次已有公式。
                experimental=convertExperimental(word,api,candidate,experimentalOptions,reverse,eqp==L"dollar",scan.items.size()+1,afterPackage,selfTest);
                // 表格前后的正文片段已核验不变；使用各表格的累计位移，不把多个累计值相加。
                for(auto& outcome:outcomes){
                    const auto after=std::upper_bound(experimental.bodyPositionShifts.begin(),experimental.bodyPositionShifts.end(),outcome.finalSpan.start,
                        [](long position,const BodyPositionShift& shift){return position<shift.boundary;});
                    if(after!=experimental.bodyPositionShifts.begin()){const long offset=(after-1)->offset;outcome.finalSpan.start+=offset;outcome.finalSpan.end+=offset;}
                }
                word.shutdown();if(taskShared)InterlockedExchange(&taskShared->mathReady,0);
                if(api.initialized){require(api.term()==0,"Terminate experimental MathType API");api.initialized=false;}
            }else experimental.skipped=experimentalSkip;
        }
        // 链接更新方式在正文与实验阶段共用同一次设置，实验保存后再统一恢复，避免重复读写整包。
        setDocxLinkUpdates(candidate,&linkUpdates);
        std::puts("TIMING_STAGE SOURCE_COMPARE");
        readOriginal(&originalBytes);
        require(CloseHandle(sourceLock) != 0, "Release original DOCX lock"); sourceLock = INVALID_HANDLE_VALUE;
        std::puts("UI_PHASE PUBLISHING");
        taskTestBoundary(L"BEFORE_PUBLISH");taskReservePrivatePath(stagedReport);
        writeNewText(stagedReport, outcomeReport(input, output, eqp, scan, outcomes,selfTest,experimental,issues)); reportOwned = true;
        taskMove(candidate, output);
        taskTestBoundary(L"AFTER_OUTPUT_MOVE");
        candidateOwned = false; outputOwned = true;
        if (selfTest) {
            taskMove(candidate + L".pdf", output + L".pdf");
            previewOwned = false; publishedPreviewOwned = true;
        }
        taskMove(stagedReport, report); reportOwned = false;
        taskCommit();
        taskTestBoundary(L"AFTER_COMMIT");
        outputOwned = false; publishedPreviewOwned = false;
        const auto failed = std::count_if(outcomes.begin(), outcomes.end(), [](const ItemOutcome& item) { return item.status == L"FAILED_PRESERVED"; });
        const auto rejected=std::count_if(scan.items.begin(),scan.items.end(),[](const ScanItem& item){return !item.ready&&!item.existingDisplay&&!item.existingReference;});
        for(const auto& issue:issues)std::printf("UI_PROCESSING_ISSUE %s %s %s\n",wireHex(utf8(issue.region)).c_str(),wireHex(utf8(issue.stage)).c_str(),wireHex(utf8(issue.reason)).c_str());
        for(const auto& issue:experimental.issues)std::printf("UI_PROCESSING_ISSUE %s %s %s\n",wireHex(utf8(issue.region)).c_str(),wireHex(utf8(issue.stage)).c_str(),wireHex(utf8(issue.reason)).c_str());
        std::printf("UI_CONVERT_SUMMARY %s %zu %zu %zu %zu %zu\n", failed || rejected || !issues.empty() || experimental.partial() ? "PARTIAL" : "COMPLETE",
                     static_cast<size_t>(total)+experimental.succeeded+experimental.failed+experimental.unattempted,static_cast<size_t>(succeeded)+experimental.succeeded,
                     static_cast<size_t>(failed)+experimental.failed,static_cast<size_t>(total)-completed+experimental.unattempted,static_cast<size_t>(rejected)+experimental.rejected);
        std::puts("BATCH_CONVERSION_PASSED"); code = 0;
    } catch (const std::exception& error) {
        std::printf("UI_ERROR %s\n", wireHex(error.what()).c_str());
        std::printf("BATCH_CONVERSION_FAILED: %s\n", error.what());
    }
    if (sourceLock != INVALID_HANDLE_VALUE) CloseHandle(sourceLock);
    if(sizePreferenceOwned&&!DeleteFileW(sizePreferencePath.c_str()))std::puts("OWNED_SIZE_PREFERENCE_CLEANUP_FAILED");
    // 服务断连时先释放 COM 公寓，避免 OLE 存储的延迟释放阻止工作副本清理。
    CoUninitialize();
    // 仅删除本次新建、尚未交付的工作副本/报告，不触碰任何已有用户文件。
    if (candidateOwned && !DeleteFileW(candidate.c_str())) std::printf("OWNED_BATCH_COPY_CLEANUP_FAILED WIN32=%lu\n", GetLastError());
    if (reportOwned && !DeleteFileW(stagedReport.c_str())) std::puts("OWNED_BATCH_REPORT_CLEANUP_FAILED");
    if (outputOwned && !DeleteFileW(output.c_str())) std::puts("OWNED_BATCH_OUTPUT_ROLLBACK_FAILED");
    if (previewOwned && !DeleteFileW((candidate + L".pdf").c_str())) std::puts("OWNED_BATCH_PREVIEW_CLEANUP_FAILED");
    if (publishedPreviewOwned && !DeleteFileW((output + L".pdf").c_str())) std::puts("OWNED_BATCH_PREVIEW_ROLLBACK_FAILED");
    return code;
}

// 有界恢复进程只消费未知的 NEXT_EQN，再确认自然默认和全局偏好保持，绝不覆盖用户默认。
int recoverPreferencesMain(const std::wstring& originalName) {
    const auto initialized = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED); if (FAILED(initialized)) return 1;
    int code = 1;
    try {
        TaskMapping original(originalName, false); auto& originalState = *original.state;
        WordSession word; auto documents = get(word.application.Get(), L"Documents");
        word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD)); word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false}); invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
        if (originalState.linksChanged) {
            invoke(get(word.application.Get(), L"Options").Get(), L"UpdateLinksAtOpen", DISPATCH_PROPERTYPUT, {originalState.originalLinks});
            InterlockedExchange(&originalState.linksChanged, 0); std::puts("WORD_OPTIONS_RECOVERY_PASSED");
        }
        if (originalState.pendingPrefs) {
            const auto& record = originalState.recovery;
            require(record.nativeLength && record.nativeLength <= TaskNativeCapacity, "Recover only recorded preference baseline");
            MathTypeApi api(record.install); require(api.init(0, 10) > 0, "Initialize preference recovery API"); api.initialized = true;
            taskMathTypeConnected();
            require(readDefaultPrefs(api) == record.defaults, "User defaults changed; refuse preference recovery");
            const auto consumed = inspectStyle(emptyEquation(api, word));
            if (originalState.expectedPendingLength) {
                const auto expected = inspectStyle(std::vector<unsigned char>(originalState.expectedPendingNative, originalState.expectedPendingNative + originalState.expectedPendingLength));
                const auto natural = inspectStyle(std::vector<unsigned char>(record.native, record.native + record.nativeLength));
                require((consumed.inlineEquation == expected.inlineEquation && consumed.sizes == expected.sizes && consumed.styles == expected.styles) ||
                    (consumed.inlineEquation == natural.inlineEquation && consumed.sizes == natural.sizes && consumed.styles == natural.styles), "Recovered pending equation has task or natural style");
                std::puts("PENDING_STYLE_CONSUMPTION_VERIFIED");
            }
            requireStyle(inspectStyle(emptyEquation(api, word)), inspectStyle(std::vector<unsigned char>(record.native, record.native + record.nativeLength)));
            require(readDefaultPrefs(api) == record.defaults, "Global preferences unchanged after recovery");
            if (taskShared) InterlockedExchange(&taskShared->mathReady, 0);
            require(api.term() == 0, "Terminate preference recovery API"); api.initialized = false;
            InterlockedExchange(&originalState.pendingPrefs, 0); std::puts("PREFERENCE_RECOVERY_PASSED");
        }
        word.shutdown(); code = 0;
    } catch (const std::exception& error) { taskError(error.what()); std::printf("PREFERENCE_RECOVERY_FAILED: %s\n", error.what()); }
    CoUninitialize(); return code;
}

#ifndef DOCX_MATH_GUI
// 独立批处理入口保留；GUI 链接同一实现，工作模式不复制转换流程。
int wmain(int argc, wchar_t** argv) {
    if (argc > 3 && std::wstring(argv[1]) == L"--task-run") {
        try { TaskMapping mapping(argv[2], false); taskShared = mapping.state;
            const int code = std::wstring(argv[3]) == L"--worker-recover" && argc == 5 ? recoverPreferencesMain(argv[4]) : convertMain(argc - 2, argv + 2);
            taskShared = nullptr; return code;
        } catch (const std::exception& error) { taskError(error.what()); return 1; }
    }
    return superviseTask(argc, argv);
}
#endif
