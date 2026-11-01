// 只读扫描与批处理实际共用的主故事范围、定界符及清单检查。
#pragma once
#include "word_automation.h"
#include "latex_rules.h"
#include "worker_protocol.h"
#include "numbering_rules.h"
#include <utility>

struct Span { long start{}, end{}; };
// 原文定位在首次修改前冻结；0 页表示未能取得页码，绝不借用转换后的排版。
struct SourcePosition {
    long sourceStart{-1},page{};
    std::wstring location;
};
struct ScannedShape {Span span;std::wstring classType;long type{};};
struct ScannedField {Span codeSpan;long kind{};std::wstring code,result;Span span;};
enum class EquationScope { All, Inline, Display };
struct ScanOptions {
    EquationScope scope{EquationScope::All};
    bool fast{},saved{},progress{};
    bool capturePoints{true};
    bool convertOmml{true};
};
inline bool selectedEquation(bool display,EquationScope scope=EquationScope::All) {
    return scope == EquationScope::All || (display ? scope == EquationScope::Display : scope == EquationScope::Inline);
}
struct ScanItem {
    Span span;
    long paragraph{}, row{}, column{}, tableStart{};
    std::wstring source, inner, delimiter, reason;
    bool ready{}, display{};
    double points{};
    bool existingDisplay{};
    long existingNativeOrdinal{};
    bool existingReference{};
    long ommlOrdinal{};
    SourcePosition position;
};
struct ScanResult {
    std::vector<ScanItem> items;
    std::vector<long> paragraphStarts;
    std::vector<ScannedShape> shapes;
    std::vector<ScannedField> fieldSnapshots;
    long paragraphs{}, existingMathType{}, inlineShapes{}, nativeMath{}, fields{}, revisions{}, controls{}, noteReferences{};
    bool protectedDocument{};
    bool tracksChanges{};
    bool reverse{};
    bool bodyOmml{};
};

inline ComPtr<IDispatch> docRange(IDispatch* document, long start, long end) {
    require(start >= 0 && end >= start, "Valid Word range bounds");
    return asObject(invoke(document, L"Range", DISPATCH_METHOD, {start, end}));
}
inline ComPtr<IDispatch> collectionItem(IDispatch* collection, long index) {
    return asObject(invoke(collection, L"Item", DISPATCH_METHOD, {index}));
}
inline long information(IDispatch* range, long type) {
    auto value = invoke(range, L"Information", DISPATCH_PROPERTYGET, {type});
    if (value.v.vt == VT_BOOL) return value.v.boolVal != VARIANT_FALSE;
    require(value.v.vt == VT_I4, "Word Information type");
    return value.v.lVal;
}
inline SourcePosition originalSourcePosition(IDispatch* document,long start,const std::wstring& region,bool paginated) {
    SourcePosition result;result.sourceStart=start;
    if(paginated&&start>=0){
        try{auto point=docRange(document,start,start);result.page=std::max(0L,information(point.Get(),3));}
        catch(const std::exception& error){require(!taskCancelled()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");}
    }
    result.location=(result.page?L"原文第 "+std::to_wstring(result.page)+L" 页":L"原文页码未知")+L" · "+region;
    return result;
}
inline std::string sourcePositionJson(const SourcePosition& position) {
    return ",\"SourceLocation\":"+quote(position.location.empty()?L"原文页码未知":position.location)+",\"Page\":"+(position.page?std::to_string(position.page):"null");
}
inline Span bounds(IDispatch* range) { return {integer(range, L"Start"), integer(range, L"End")}; }
inline Span fieldBounds(IDispatch* field,ScannedField* snapshot=nullptr) {
    auto codeRange=get(field,L"Code"),resultRange=get(field,L"Result");const auto code=bounds(codeRange.Get()),value=bounds(resultRange.Get());
    const long kind=snapshot||value.start==value.end?integer(field,L"Type"):0;
    // 保存后扫描已取得域范围，顺便记录核验所需的代码和缓存，避免再遍历整份域集合。
    if(snapshot){snapshot->codeSpan=code;snapshot->kind=kind;if(kind!=58){snapshot->code=text(codeRange.Get(),L"Text");snapshot->result=text(resultRange.Get(),L"Text");}}
    Span span{code.start-1,value.end+1};
    if(value.start==value.end&&kind==51){
        const auto instruction=snapshot?snapshot->code:text(codeRange.Get(),L"Text");
        if(nativeNumberWithoutSeparator(instruction))
            span.end=code.end+1; // 原生编号没有结果分隔符，避免多占相邻内容。
    }
    if(snapshot)snapshot->span=span;return span;
}

// 子范围从所属文本区域复制，避免将页眉、文本框或注释坐标当成正文坐标。
inline ComPtr<IDispatch> sliceRange(IDispatch* root,long start,long end) {
    require(start>=integer(root,L"Start")&&end>=start&&end<=integer(root,L"End"),"Range remains inside owning story");
    auto range=get(root,L"Duplicate");invoke(range.Get(),L"SetRange",DISPATCH_METHOD,{start,end});return range;
}
// 正文继续使用原来的 Document.Range；额外区域才采用所属 Range 的坐标。
struct ScanRangeContext { IDispatch* document{}; IDispatch* story{};bool captureSize{true};ScanOptions options{}; };
inline ComPtr<IDispatch> docRange(const ScanRangeContext& context,long start,long end) {
    return context.story?sliceRange(context.story,start,end):docRange(context.document,start,end);
}
// 所有定位从 Word Find 返回的实际位置取得，绝不把 Range.Text 下标加到 Start。
inline ComPtr<IDispatch> findLiteral(const ScanRangeContext& document, long start, long end, const wchar_t* needle) {
    // 折叠 Range 的 Find 会从插入点向故事末尾扩展；空范围必须直接结束，不能继续查找。
    if (start >= end) return {};
    auto range = docRange(document, start, end);
    auto find = get(range.Get(), L"Find");
    invoke(find.Get(), L"ClearFormatting", DISPATCH_METHOD);
    std::wstring escaped;
    for (const wchar_t* character = needle; *character; ++character) {
        escaped += *character;
        if (*character == L'^') escaped += L'^';
    }
    invoke(find.Get(), L"Text", DISPATCH_PROPERTYPUT, {escaped.c_str()});
    invoke(find.Get(), L"Forward", DISPATCH_PROPERTYPUT, {true});
    invoke(find.Get(), L"Wrap", DISPATCH_PROPERTYPUT, {0L});
    invoke(find.Get(), L"MatchWildcards", DISPATCH_PROPERTYPUT, {false});
    invoke(find.Get(), L"MatchCase", DISPATCH_PROPERTYPUT, {true});
    invoke(find.Get(), L"MatchWholeWord", DISPATCH_PROPERTYPUT, {false});
    invoke(find.Get(), L"Format", DISPATCH_PROPERTYPUT, {false});
    auto found = invoke(find.Get(), L"Execute", DISPATCH_METHOD);
    require(found.v.vt == VT_BOOL, "Word Find return type");
    if (found.v.boolVal == VARIANT_FALSE) return {};
    const auto actual = bounds(range.Get());
    require(actual.start >= start && actual.end <= end && actual.end > actual.start && (document.options.fast||text(range.Get(), L"Text") == needle),
            "Literal Find remains inside requested range");
    return range;
}

inline bool escapedAt(const ScanRangeContext& document, long position, long boundary) {
    long slashes{};
    while (position > boundary && text(docRange(document, position - 1, position).Get(), L"Text") == L"\\") {
        --position; ++slashes;
    }
    return slashes % 2 != 0;
}
enum class TokenKind { Dollar, DoubleDollar, BracketOpen, BracketClose, ParenOpen, ParenClose, RepairedBracketOpen, AmbiguousDollars, EmptyDisplay };
struct Token { Span span; TokenKind kind; };

inline bool hasTexMarkers(const std::wstring& source) {
    return source.find(L'$') != std::wstring::npos || source.find(L"\\[") != std::wstring::npos ||
        source.find(L"\\]") != std::wstring::npos || source.find(L"\\(") != std::wstring::npos || source.find(L"\\)") != std::wstring::npos;
}
inline bool tokensMatch(TokenKind opening, TokenKind closing) {
    return (opening == TokenKind::Dollar && closing == TokenKind::Dollar) ||
        (opening == TokenKind::DoubleDollar && closing == TokenKind::DoubleDollar) ||
        ((opening == TokenKind::BracketOpen || opening == TokenKind::RepairedBracketOpen) && closing == TokenKind::BracketClose) ||
        (opening == TokenKind::ParenOpen && closing == TokenKind::ParenClose);
}
inline std::vector<Token> tokens(const ScanRangeContext& document, Span segment, const std::wstring& source) {
    std::vector<Token> result;
    long cursor = segment.start;
    while (auto found = findLiteral(document, cursor, segment.end, L"$")) {
        auto dollar = bounds(found.Get());
        if (escapedAt(document, dollar.start, segment.start)) { cursor = dollar.end; continue; }
        long count = 1;
        while (dollar.end < segment.end && text(docRange(document, dollar.end, dollar.end + 1).Get(), L"Text") == L"$") {
            ++dollar.end; ++count;
        }
        const auto kind = count == 1 ? TokenKind::Dollar : count == 2 ? TokenKind::DoubleDollar :
                          count == 4 ? TokenKind::EmptyDisplay : TokenKind::AmbiguousDollars;
        result.push_back({dollar, kind});
        cursor = dollar.end;
    }
    for (const auto& marker : {std::make_pair(L"\\[", TokenKind::BracketOpen), std::make_pair(L"\\]", TokenKind::BracketClose),
                               std::make_pair(L"\\(", TokenKind::ParenOpen), std::make_pair(L"\\)", TokenKind::ParenClose)}) {
        if (source.find(marker.first) == std::wstring::npos) continue;
        cursor = segment.start;
        while (auto found = findLiteral(document, cursor, segment.end, marker.first)) {
            const auto span = bounds(found.Get());
            if (!escapedAt(document, span.start, segment.start)) result.push_back({span, marker.second});
            cursor = span.end;
        }
    }
    // 仅补整段开头缺失的反斜杠，不能把正文或 TeX 可选参数中的方括号当成定界符。
    const auto first = source.find_first_not_of(L" \t\r\n\v");
    if (result.size() == 1 && result.front().kind == TokenKind::BracketClose && first != std::wstring::npos && source[first] == L'[') {
        auto opening = findLiteral(document, segment.start, result.front().span.start, L"[");
        if (opening) result.push_back({bounds(opening.Get()), TokenKind::RepairedBracketOpen});
    }
    std::sort(result.begin(), result.end(), [](const Token& a, const Token& b) { return a.span.start < b.span.start; });
    return result;
}

inline double equationPoints(IDispatch* range,const ScanRangeContext& document,Span inner) {
    auto size=invoke(get(range,L"Font").Get(),L"Size",DISPATCH_PROPERTYGET);
    double points=size.v.vt==VT_R4?size.v.fltVal:size.v.vt==VT_R8?size.v.dblVal:size.v.vt==VT_I4?size.v.lVal:0;
    // 混合字号读取内容首个非空字符，避免将 Word 的 wdUndefined 当成字号。
    if(!(points>=1&&points<=256)&&inner.start>=0&&inner.end>inner.start){
        auto characters=get(docRange(document,inner.start,inner.end).Get(),L"Characters");
        const long characterCount=integer(characters.Get(),L"Count");
        for(long i=1;i<=characterCount;++i){auto character=collectionItem(characters.Get(),i);if(whitespaceOnly(text(character.Get(),L"Text")))continue;
            auto effective=invoke(get(character.Get(),L"Font").Get(),L"Size",DISPATCH_PROPERTYGET);
            points=effective.v.vt==VT_R4?effective.v.fltVal:effective.v.vt==VT_R8?effective.v.dblVal:effective.v.vt==VT_I4?effective.v.lVal:0;break;}
    }
    return points;
}
inline void addItem(const ScanRangeContext& document, Span span, Span inner, TokenKind kind, const std::wstring& reason,
                    long paragraph, ScanResult& result, const std::wstring& paragraphText, Span allowed) {
    auto range = docRange(document, span.start, span.end);
    ScanItem item;
    item.span = span; item.position.sourceStart=span.start;item.paragraph = paragraph; item.reason = reason; item.ready = reason.empty();
    item.source = text(range.Get(), L"Text");
    if (inner.end >= inner.start && inner.start >= span.start && inner.end <= span.end)
        item.inner = text(docRange(document, inner.start, inner.end).Get(), L"Text");
    item.delimiter = kind == TokenKind::Dollar ? L"$" : kind == TokenKind::DoubleDollar || kind == TokenKind::EmptyDisplay ? L"$$" :
                     kind == TokenKind::BracketOpen || kind == TokenKind::BracketClose ? L"\\[\\]" :
                     kind == TokenKind::RepairedBracketOpen ? L"[\\]" :
                     kind == TokenKind::ParenOpen || kind == TokenKind::ParenClose ? L"\\(\\)" : L"AMBIGUOUS";
    item.display = kind != TokenKind::Dollar && kind != TokenKind::ParenOpen && kind != TokenKind::ParenClose && kind != TokenKind::AmbiguousDollars;
    Span manualSpan{};bool manualNumber=false;
    if (item.ready && !paragraphText.empty()) {
        auto context = paragraphText;
        if (item.source.find(L'\r') != std::wstring::npos) {
            auto paragraphs = get(range.Get(), L"Paragraphs");
            const auto first = get(collectionItem(paragraphs.Get(), 1).Get(), L"Range");
            const auto last = get(collectionItem(paragraphs.Get(), integer(paragraphs.Get(), L"Count")).Get(), L"Range");
            context = text(docRange(document, integer(first.Get(), L"Start"), integer(last.Get(), L"End")).Get(), L"Text");
        }
        // 只比较文本判断上下文，不用字符串下标推算 Word Range 坐标。
        const auto position = context.find(item.source);
        const bool leadingOnly=position!=std::wstring::npos&&whitespaceOnly(context.substr(0,position));
        const auto trailing=leadingOnly?context.substr(position+item.source.size()):std::wstring();
        if(leadingOnly&&manualEquationNumber(trailing)){
            auto paragraphs=get(range.Get(),L"Paragraphs");auto first=get(collectionItem(paragraphs.Get(),1).Get(),L"Range");auto last=get(collectionItem(paragraphs.Get(),integer(paragraphs.Get(),L"Count")).Get(),L"Range");
            manualSpan={integer(first.Get(),L"Start"),integer(last.Get(),L"End")-1};
            // 编号和公式须在同一安全片段内，不能跨过域、控件或其他公式。
            manualNumber=manualSpan.start>=allowed.start&&manualSpan.end<=allowed.end;
        }
        item.display=leadingOnly&&(whitespaceOnly(trailing)||manualNumber);
        if (kind == TokenKind::RepairedBracketOpen && !item.display) { item.ready = false; item.reason = L"UNSAFE_MISSING_DELIMITER"; }
    }
    if (!selectedEquation(item.display,document.options.scope)) return;
    if(item.ready&&item.display&&!paragraphText.empty()){
        // 外侧空白仅在当前安全片段内扩展，不能吞掉空结果域或空白内容控件。
        if(manualNumber)item.span=manualSpan;
        else{auto paragraphs=get(range.Get(),L"Paragraphs");auto first=get(collectionItem(paragraphs.Get(),1).Get(),L"Range");auto last=get(collectionItem(paragraphs.Get(),integer(paragraphs.Get(),L"Count")).Get(),L"Range");
            item.span={std::max(integer(first.Get(),L"Start"),allowed.start),std::min(integer(last.Get(),L"End")-1,allowed.end)};}
        item.source=text(docRange(document,item.span.start,item.span.end).Get(),L"Text");
    }
    item.inner = normalizedTex(item.inner);
    if(!document.captureSize){result.items.push_back(std::move(item));return;}
    item.points=equationPoints(range.Get(),document,inner);
    result.items.push_back(std::move(item));
}

inline long scanSegment(const ScanRangeContext& document, Span segment, long paragraph, ScanResult& result, long boundary, const std::wstring& paragraphText) {
    const auto source=text(docRange(document,segment.start,segment.end).Get(),L"Text");
    if (!hasTexMarkers(source)) return segment.end;
    const Span allowed{segment.start,boundary};
    Token opening{};
    bool haveOpening = false;
    for (const auto& token : tokens(document, segment, source)) {
        if (token.kind == TokenKind::AmbiguousDollars || token.kind == TokenKind::EmptyDisplay) {
            addItem(document, {haveOpening ? opening.span.start : token.span.start, token.span.end}, {-1, -1},
                    token.kind, token.kind == TokenKind::EmptyDisplay && !haveOpening ? L"EMPTY" : L"AMBIGUOUS_DOLLARS", paragraph, result, paragraphText, allowed);
            haveOpening = false; continue;
        }
        if (!haveOpening) {
            if (token.kind == TokenKind::BracketClose || token.kind == TokenKind::ParenClose)
                addItem(document, token.span, {-1, -1}, token.kind, L"UNMATCHED_CLOSE", paragraph, result, paragraphText, allowed);
            else { opening = token; haveOpening = true; }
            continue;
        }
        const bool closes = tokensMatch(opening.kind, token.kind);
        if (closes) {
            const Span inner{opening.span.end, token.span.start};
            addItem(document, {opening.span.start, token.span.end}, inner, opening.kind,
                    lexicalReason(text(docRange(document, inner.start, inner.end).Get(), L"Text"),document.options.fast), paragraph, result, paragraphText, allowed);
        } else {
            // 混用定界符时消费冲突标记，保守拒绝，不生成互相覆盖的可转换范围。
            addItem(document, {opening.span.start, token.span.end}, {-1, -1}, opening.kind, L"MIXED_DELIMITERS", paragraph, result, paragraphText, allowed);
        }
        haveOpening = false;
    }
    if (haveOpening) {
        // 仅完整结构化公式可跨 Word 段落；下一起始标记及排除范围均为硬边界。
        const auto needle=opening.kind==TokenKind::BracketOpen || opening.kind==TokenKind::RepairedBracketOpen?L"\\]":
            opening.kind==TokenKind::ParenOpen?L"\\)":opening.kind==TokenKind::DoubleDollar?L"$$":L"$";
        auto close=findLiteral(document,segment.end,boundary,needle);
        while(close && escapedAt(document,integer(close.Get(),L"Start"),segment.end))close=findLiteral(document,integer(close.Get(),L"End"),boundary,needle);
        if(close) {
            const auto closing=bounds(close.Get());const Span inner{opening.span.end,closing.start};
            const auto body=text(docRange(document,inner.start,inner.end).Get(),L"Text");
            const auto all=tokens(document,{opening.span.start,closing.end},text(docRange(document,opening.span.start,closing.end).Get(),L"Text"));
            if(body.find_first_of(L"\r\n\v")!=std::wstring::npos && structuredMultiline(body) && all.size()==2 &&
                tokensMatch(opening.kind, all.back().kind)) {
                addItem(document,{opening.span.start,closing.end},inner,opening.kind,lexicalReason(body,document.options.fast),paragraph,result,paragraphText,allowed);return closing.end;
            }
        }
        addItem(document, {opening.span.start, segment.end}, {-1, -1}, opening.kind, L"UNMATCHED_OPEN", paragraph, result, paragraphText, allowed);
    }
    return segment.end;
}

// 复用现有扫描识别独立旧公式；没有已有对象的普通输入不增加任何段落检查。
inline bool addExistingDisplay(IDispatch* document,IDispatch* range,long paragraph,ScanResult& result,const std::vector<Span>& protectedRanges) {
    const auto para=bounds(range);const ScannedShape* candidate=nullptr;long nativeOrdinal{},candidateOrdinal{};
    for(const auto& shape:result.shapes){
        if(shape.classType.find(L"Equation.DSMT")==0)++nativeOrdinal;
        if(shape.span.start<para.start||shape.span.end>para.end)continue;
        if(candidate||shape.type!=1||shape.classType!=L"Equation.DSMT4")return false;
        candidate=&shape;candidateOrdinal=nativeOrdinal;
    }
    if(!candidate||std::any_of(protectedRanges.begin(),protectedRanges.end(),[&](const Span& span){return span.start<para.end&&span.end>para.start;}))return false;
    std::vector<Span> excluded{candidate->span};
    for(const auto& field:result.fieldSnapshots){
        if(field.span.end<=para.start||field.span.start>=para.end)continue;
        if(field.span.start<para.start||field.span.end>para.end)return false;
        const bool number=nativeNumberField(field.kind,field.code);
        const bool counter=nativeSequenceField(field.kind,field.code);
        if(field.kind!=58&&!number&&!counter)return false;
        excluded.push_back(field.span);
    }
    std::sort(excluded.begin(),excluded.end(),[](const Span& a,const Span& b){return a.start<b.start;});
    long cursor=para.start;
    for(const auto& span:excluded){if(span.start>cursor&&!whitespaceOnly(text(docRange(document,cursor,span.start).Get(),L"Text")))return false;cursor=std::max(cursor,span.end);}
    if(cursor<para.end&&!whitespaceOnly(text(docRange(document,cursor,para.end).Get(),L"Text")))return false;
    ScanItem item;item.span=para;item.position.sourceStart=candidate->span.start;item.paragraph=paragraph;item.source=text(range,L"Text");item.display=true;item.existingDisplay=true;item.existingNativeOrdinal=candidateOrdinal;item.delimiter=L"EXISTING_DISPLAY";
    result.items.push_back(std::move(item));return true;
}

inline ScanResult scanDocumentText(IDispatch* document, bool paragraphContext,bool captureFields=false,bool includeExisting=false,bool includeOmml=false,const ScanOptions& options={}) {
    if(options.saved)std::puts("TIMING_STAGE RESCAN_BOUNDARIES");
    ScanResult result;
    const bool capturePoints=options.capturePoints&&!options.saved;
    const bool scanOmml=includeOmml&&options.convertOmml;
    const auto tracking = invoke(document, L"TrackRevisions", DISPATCH_PROPERTYGET);
    require(tracking.v.vt == VT_BOOL, "Revision tracking property type");
    result.tracksChanges = tracking.v.boolVal != VARIANT_FALSE;
    result.protectedDocument = integer(document, L"ProtectionType") != -1;
    if (result.protectedDocument) return result;
    auto content = get(document, L"Content");
    require(integer(content.Get(), L"StoryType") == 1, "Main story only");
    const long contentEnd=integer(content.Get(),L"End");
    std::vector<Span> excluded,protectedRanges,ommlExcluded;
    auto exclude = [&](IDispatch* range,bool protect=true,bool blockOmml=true,bool control=false) {
        require(integer(range, L"StoryType") == 1, "Excluded main story range");
        excluded.push_back(bounds(range));
        // ContentControl.Range 只含内容，首尾各一位的控件标记也必须排除。
        if(control){excluded.back().start=std::max(0L,excluded.back().start-1);excluded.back().end=std::min(contentEnd,excluded.back().end+1);}
        if(protect)protectedRanges.push_back(excluded.back());
        if(scanOmml&&protect&&blockOmml)ommlExcluded.push_back(excluded.back());
        return excluded.back();
    };
    auto tables=get(document,L"Tables");
    const long tableCount=integer(tables.Get(),L"Count");
    for(long i=1;i<=tableCount;++i)exclude(get(collectionItem(tables.Get(),i).Get(),L"Range").Get());
    auto shapes = get(content.Get(), L"InlineShapes");
    result.inlineShapes = integer(shapes.Get(), L"Count");
    result.shapes.reserve(static_cast<size_t>(result.inlineShapes));
    for (long index = 1; index <= result.inlineShapes; ++index) {
        auto shape = collectionItem(shapes.Get(), index);
        auto range=get(shape.Get(),L"Range");exclude(range.Get(),false);
        ScannedShape snapshot{excluded.back(),L""};
        const auto type = integer(shape.Get(), L"Type");
        snapshot.type=type;
        if (type == 1 || type == 2) {
            const auto classType = text(get(shape.Get(), L"OLEFormat").Get(), L"ClassType");
            snapshot.classType=classType;
            if (classType.find(L"Equation.DSMT") == 0) ++result.existingMathType;
        }
        result.shapes.push_back(std::move(snapshot));
    }
    // 只读扫描内复用原生公式对象和真实范围，候选构造不再访问一次 COM 集合。
    struct OmmlRead {ComPtr<IDispatch> math,range;Span span;long ordinal;};
    std::vector<OmmlRead> ommlReads;
    for (const auto& entry : {std::make_pair(L"OMaths", &result.nativeMath), std::make_pair(L"Revisions", &result.revisions),
                              std::make_pair(L"ContentControls", &result.controls)}) {
        auto collection = get(content.Get(), entry.first);
        *entry.second = integer(collection.Get(), L"Count");
        const bool omml=wcscmp(entry.first,L"OMaths")==0;
        if(scanOmml&&omml)ommlReads.reserve(static_cast<size_t>(*entry.second));
        for (long index = 1; index <= *entry.second; ++index) {
            auto object=collectionItem(collection.Get(),index);auto range=get(object.Get(),L"Range");
            const auto span=exclude(range.Get(),true,!omml,wcscmp(entry.first,L"ContentControls")==0);
            if(scanOmml&&omml)ommlReads.push_back({std::move(object),std::move(range),span,index});
        }
    }
    // 关闭转换仍需为保留公式建立 XML 基线；候选开关不能关闭范围边界或核验。
    result.bodyOmml=includeOmml&&result.nativeMath>0;
    auto fields = get(content.Get(), L"Fields");
    result.fields = integer(fields.Get(), L"Count");
    captureFields=captureFields||(includeExisting&&result.existingMathType>0);
    if(captureFields)result.fieldSnapshots.resize(static_cast<size_t>(result.fields));
    for (long index = 1; index <= result.fields; ++index) {
        auto field = collectionItem(fields.Get(), index);
        excluded.push_back(fieldBounds(field.Get(),captureFields?&result.fieldSnapshots[static_cast<size_t>(index-1)]:nullptr));
        if(scanOmml)ommlExcluded.push_back(excluded.back());
    }
    for (const auto name : {L"Footnotes", L"Endnotes"}) {
        auto notes = get(document, name);
        const auto count = integer(notes.Get(), L"Count");
        for (long index = 1; index <= count; ++index) {
            exclude(get(collectionItem(notes.Get(), index).Get(), L"Reference").Get()); ++result.noteReferences;
        }
    }
    std::sort(excluded.begin(), excluded.end(), [](const Span& a, const Span& b) { return a.start < b.start; });
    std::vector<Span> merged;
    for(const auto& span:excluded) {if(!merged.empty()&&span.start<=merged.back().end)merged.back().end=std::max(merged.back().end,span.end);else merged.push_back(span);}
    excluded=std::move(merged);
    if(options.saved)std::puts("TIMING_STAGE RESCAN_PARAGRAPHS");
    auto paragraphs = get(content.Get(), L"Paragraphs");
    result.paragraphs = integer(paragraphs.Get(), L"Count");
    std::vector<ComPtr<IDispatch>> savedParagraphs;
    if(options.saved){
        // 只读重扫按批顺序枚举段落，避免每段再经 Item(index) 查找；Word 坐标仍逐范围读取。
        const auto source=invoke(paragraphs.Get(),L"_NewEnum",DISPATCH_PROPERTYGET);ComPtr<IEnumVARIANT> sequence;
        require(source.v.vt==VT_UNKNOWN&&source.v.punkVal,"Word paragraph enumerator");check(source.v.punkVal->QueryInterface(IID_PPV_ARGS(&sequence)),"Word paragraph enumeration interface");
        for(;;){
            struct ParagraphBatch {VARIANT values[32]{};~ParagraphBatch(){for(auto& value:values)VariantClear(&value);}} batch;
            ULONG count{};taskPulse();check(sequence->Next(32,batch.values,&count),"Read paragraph batch");
            for(ULONG i=0;i<count;++i){Value value;std::swap(value.v,batch.values[i]);savedParagraphs.push_back(asObject(value));}
            if(count<32)break;
        }
        require(savedParagraphs.size()==static_cast<size_t>(result.paragraphs),"Saved paragraph enumeration count");
    }
    StageProgress paragraphProgress(options.saved?"RESCAN_TEXT":"SCAN_TEXT",static_cast<size_t>(result.paragraphs),options.progress);
    // OMML 上下文复用本轮已经读取的段落文本，不为识别手动编号再次读取整段。
    std::vector<std::wstring> ommlParagraphText(scanOmml&&result.bodyOmml?static_cast<size_t>(result.paragraphs):0);
    long consumed{};size_t excludedIndex{};
    for (long index = 1; index <= result.paragraphs; ++index) {
        auto paragraph = options.saved?savedParagraphs[static_cast<size_t>(index-1)]:collectionItem(paragraphs.Get(), index);
        auto range = get(paragraph.Get(), L"Range");
        auto paragraphSpan = bounds(range.Get());
        result.paragraphStarts.push_back(paragraphSpan.start);
        if(includeExisting&&result.existingMathType>0&&selectedEquation(true,options.scope)&&addExistingDisplay(document,range.Get(),index,result,protectedRanges)){paragraphProgress.advance();continue;}
        paragraphSpan.start=std::max(paragraphSpan.start,consumed);
        while(excludedIndex<excluded.size()&&excluded[excludedIndex].end<=paragraphSpan.start)++excludedIndex;
        if(paragraphSpan.start>=paragraphSpan.end || (excludedIndex<excluded.size()&&excluded[excludedIndex].start<=paragraphSpan.start&&excluded[excludedIndex].end>=paragraphSpan.end)) {paragraphProgress.advance();continue;}
        const auto paragraphText=text(range.Get(),L"Text");
        if(scanOmml&&result.bodyOmml)ommlParagraphText[static_cast<size_t>(index-1)]=paragraphText;
        if (!hasTexMarkers(paragraphText)) {paragraphProgress.advance();continue;}
        // Word 单元格末尾的文本可能同时含 \r\a，一个真实字符范围可以返回两个标记。
        while (paragraphSpan.end > paragraphSpan.start) {
            const auto last = text(docRange(document, paragraphSpan.end - 1, paragraphSpan.end).Get(), L"Text");
            if (last.empty() || !std::all_of(last.begin(), last.end(), [](wchar_t c) { return c == L'\r' || c == L'\a'; })) break;
            --paragraphSpan.end;
        }
        long cursor = paragraphSpan.start;
        for (size_t blockIndex=excludedIndex;blockIndex<excluded.size();++blockIndex) {
            const auto& block=excluded[blockIndex];if(block.start>paragraphSpan.end)break;
            if (block.end < cursor || block.start > paragraphSpan.end) continue;
            if (block.start > cursor) consumed=std::max(consumed,scanSegment(ScanRangeContext{document,nullptr,capturePoints,options}, {cursor, std::min(block.start, paragraphSpan.end)}, index, result,block.start,paragraphContext?paragraphText:L""));
            cursor = std::max(cursor, std::min(block.end, paragraphSpan.end));
        }
        if (cursor < paragraphSpan.end) {
            long boundary=contentEnd;
            for(size_t i=excludedIndex;i<excluded.size();++i)if(excluded[i].start>=cursor){boundary=excluded[i].start;break;}
            consumed=std::max(consumed,scanSegment(ScanRangeContext{document,nullptr,capturePoints,options}, {cursor, paragraphSpan.end}, index, result,boundary,paragraphContext?paragraphText:L""));
        }
        paragraphProgress.advance();
    }
    // 仅开启 OMML 时把正文原生公式并入同一清单；表格等区域仍由实验分支处理。
    if(scanOmml&&result.bodyOmml){
        for(const auto& read:ommlReads){taskPulse();const auto& math=read.math;const auto& range=read.range;const auto span=read.span;
            if(std::any_of(ommlExcluded.begin(),ommlExcluded.end(),[&](Span block){return block.start<span.end&&block.end>span.start;}))continue;
            ScanItem item;item.span=span;item.position.sourceStart=span.start;item.ommlOrdinal=read.ordinal;item.delimiter=L"OMML";item.display=integer(math.Get(),L"Type")==0;item.ready=true;
            // 段落位置来自已记录的 Word 范围；文本匹配只判断上下文，不推算 Word 字符坐标。
            const auto paragraphIndex=static_cast<size_t>(std::upper_bound(result.paragraphStarts.begin(),result.paragraphStarts.end(),span.start)-result.paragraphStarts.begin()-1);
            item.paragraph=static_cast<long>(paragraphIndex+1);
            const Span limits{result.paragraphStarts[paragraphIndex],paragraphIndex+1<result.paragraphStarts.size()?result.paragraphStarts[paragraphIndex+1]:contentEnd};
            auto nativeText=text(range.Get(),L"Text");const auto& context=ommlParagraphText[paragraphIndex];const auto at=context.find(nativeText);
            const bool leadingOnly=!nativeText.empty()&&at!=std::wstring::npos&&whitespaceOnly(context.substr(0,at));
            const bool manual=leadingOnly&&manualEquationNumber(context.substr(at+nativeText.size()))&&
                std::none_of(protectedRanges.begin(),protectedRanges.end(),[&](Span block){return block.start<limits.end&&block.end>limits.start&&(block.start!=span.start||block.end!=span.end);})&&
                std::none_of(ommlExcluded.begin(),ommlExcluded.end(),[&](Span block){return block.start<limits.end&&block.end>limits.start;});
            if(manual){item.display=true;item.span={limits.start,limits.end-1};}
            else if(item.display){
                if(whitespaceOnly(text(docRange(document,limits.start,span.start).Get(),L"Text"))&&whitespaceOnly(text(docRange(document,span.end,limits.end).Get(),L"Text"))){
                    item.span={limits.start,limits.end-1};
                    // OMML 的段外空白同样不能跨过已记录的域和控件边界。
                    for(const auto block:ommlExcluded){if(block.end<=span.start)item.span.start=std::max(item.span.start,block.end);if(block.start>=span.end)item.span.end=std::min(item.span.end,block.start);}
                }}
            if(!selectedEquation(item.display,options.scope))continue;
            if(capturePoints)item.points=equationPoints(range.Get(),ScanRangeContext{document,nullptr,true,options},span);
            item.source=item.span.start==span.start&&item.span.end==span.end?std::move(nativeText):text(docRange(document,item.span.start,item.span.end).Get(),L"Text");result.items.push_back(std::move(item));
        }
    }
    std::sort(result.items.begin(), result.items.end(), [](const ScanItem& a, const ScanItem& b) { return a.span.start < b.span.start; });
    if(!options.fast){long previousEnd = -1;
    if(options.saved)std::puts("TIMING_STAGE RESCAN_RANGES");
    StageProgress rangeProgress(options.saved?"RESCAN_RANGES":"SCAN_RANGES",result.items.size(),options.progress);
    for (const auto& item : result.items) {
        require(item.span.start >= previousEnd && item.span.end > item.span.start, "Non-overlapping scan items");
        require(text(docRange(document, item.span.start, item.span.end).Get(), L"Text") == item.source, "All reported ranges rechecked");
        previousEnd = item.span.end;
        rangeProgress.advance();
    }
    }
    return result;
}
inline ScanResult scanDocument(IDispatch* document,const ScanOptions& options={}) { return scanDocumentText(document,true,false,false,true,options); }
