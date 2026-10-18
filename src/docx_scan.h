// 只读扫描与批处理实际共用的主故事范围、定界符及清单检查。
#pragma once
#include "word_automation.h"
#include <climits>
#include <cwctype>
#include <msopc.h>
#include <msxml6.h>
#include <map>
#include <cmath>

using DocxLinkUpdates = std::map<std::wstring, std::vector<std::wstring>>;
// 只修改私有副本的 OLE 外链更新方式，正常保存后恢复，不修改 Word 全局选项。
inline DocxLinkUpdates setDocxLinkUpdates(const std::wstring& path, const DocxLinkUpdates* restore = nullptr) {
    if (restore && restore->empty()) return {};
    ComPtr<IOpcFactory> factory; check(CoCreateInstance(__uuidof(OpcFactory),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&factory)),"Link settings package factory");
    ComPtr<IStream> input; check(factory->CreateStreamOnFile(path.c_str(),OPC_STREAM_IO_READ,nullptr,0,&input),"Read private link settings");
    ComPtr<IOpcPackage> package; check(factory->ReadPackageFromStream(input.Get(),OPC_VALIDATE_ON_LOAD,&package),"Read private package for links");
    DocxLinkUpdates updates; size_t restored{};
    {
        ComPtr<IOpcPartSet> parts;check(package->GetPartSet(&parts),"Link settings parts");
        ComPtr<IOpcPartEnumerator> entries;check(parts->GetEnumerator(&entries),"Link settings part enumeration");BOOL more{};HRESULT advanced{};
        while((advanced=entries->MoveNext(&more))==S_OK&&more){
            taskPulse();ComPtr<IOpcPart> part;check(entries->GetCurrent(&part),"Link settings current part");
            ComPtr<IOpcPartUri> uri;check(part->GetName(&uri),"Link settings part name");Value name;name.v.vt=VT_BSTR;check(uri->GetAbsoluteUri(&name.v.bstrVal),"Link settings part URI");
            const std::wstring partName=name.v.bstrVal;
            if(partName.rfind(L"/word/",0)!=0||(partName.size()<4)||(partName.substr(partName.size()-4)!=L".xml"&&partName.substr(partName.size()-4)!=L".vml"))continue;
            if(restore&&restore->find(partName)==restore->end())continue;
            ComPtr<IStream> content;check(part->GetContentStream(&content),"Read XML link settings");
            ComPtr<IXMLDOMDocument2> document;check(CoCreateInstance(__uuidof(DOMDocument60),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&document)),"System XML link settings reader");
            check(document->put_async(VARIANT_FALSE),"Synchronous link settings XML");check(document->put_preserveWhiteSpace(VARIANT_TRUE),"Keep XML whitespace");
            check(document->put_resolveExternals(VARIANT_FALSE),"Disable XML external resolution");check(document->setProperty(Value(L"ProhibitDTD").v.bstrVal,Value(true).v),"Disable XML DTD");
            Value source;source.v.vt=VT_UNKNOWN;source.v.punkVal=content.Get();source.v.punkVal->AddRef();VARIANT_BOOL loaded{};
            check(document->load(source.v,&loaded),"Load private XML part");require(loaded!=VARIANT_FALSE,"Valid private XML part");
            ComPtr<IXMLDOMNodeList> links;check(document->selectNodes(Value(L"//*[local-name()='OLEObject' and namespace-uri()='urn:schemas-microsoft-com:office:office' and @Type='Link']").v.bstrVal,&links),"Find linked OLE objects");
            long count{};check(links->get_length(&count),"Count linked OLE objects");if(!count&&!restore)continue;
            if(restore){require(static_cast<size_t>(count)==restore->at(partName).size(),"Saved linked OLE objects preserved");++restored;}
            for(long index=0;index<count;++index){
                ComPtr<IXMLDOMNode> node;check(links->get_item(index,&node),"Read linked OLE node");ComPtr<IXMLDOMElement> element;check(node.As(&element),"Linked OLE element");
                if(!restore){Value mode;check(element->getAttribute(Value(L"UpdateMode").v.bstrVal,&mode.v),"Read original OLE update mode");
                    require(mode.v.vt==VT_NULL||mode.v.vt==VT_BSTR,"OLE update mode type");updates[partName].push_back(mode.v.vt==VT_BSTR?mode.v.bstrVal:L"");
                    check(element->setAttribute(Value(L"UpdateMode").v.bstrVal,Value(L"OnCall").v),"Disable private OLE automatic update");
                }else{const auto& mode=restore->at(partName)[static_cast<size_t>(index)];
                    if(mode.empty())check(element->removeAttribute(Value(L"UpdateMode").v.bstrVal),"Restore absent OLE update mode");
                    else check(element->setAttribute(Value(L"UpdateMode").v.bstrVal,Value(mode.c_str()).v),"Restore original OLE update mode");
                }
            }
            ULARGE_INTEGER zeroSize{};LARGE_INTEGER zero{};check(content->SetSize(zeroSize),"Replace private XML part");check(content->Seek(zero,STREAM_SEEK_SET,nullptr),"Rewind private XML output");
            check(document->save(source.v),"Write private OLE link settings");
        }
        check(advanced,"Complete linked OLE part enumeration");
    }
    if(restore)require(restored==restore->size(),"All original linked OLE parts preserved");
    if(!restore&&updates.empty())return updates;
    ComPtr<IStream> output;check(CreateStreamOnHGlobal(nullptr,TRUE,&output),"Private package rewrite buffer");
    check(factory->WritePackageToStream(package.Get(),OPC_WRITE_DEFAULT,output.Get()),"Write private link settings package");package.Reset();input.Reset();
    LARGE_INTEGER zero{};check(output->Seek(zero,STREAM_SEEK_SET,nullptr),"Rewind private package buffer");
    ComPtr<IStream> destination;check(factory->CreateStreamOnFile(path.c_str(),OPC_STREAM_IO_WRITE,nullptr,0,&destination),"Replace owned private package");
    ULARGE_INTEGER all{};all.QuadPart=~0ULL;check(output->CopyTo(destination.Get(),all,nullptr,nullptr),"Complete private package rewrite");return updates;
}

// 只在 GUI 工作进程输出计数；阶段首尾立即发送，中途最多每 150 ms 更新一次。
inline bool guiStageProgress{},checkingSavedDocument{};
struct StageProgress {
    const char* stage{};size_t total{},completed{};ULONGLONG lastUpdate{};
    StageProgress(const char* name,size_t count):stage(name),total(count) {
        if(guiStageProgress&&stage&&total){std::printf("UI_STAGE_PROGRESS %s 0 %zu\n",stage,total);lastUpdate=GetTickCount64();}
    }
    void advance(){++completed;if(completed==1&&stage){
        if(!strcmp(stage,"SCAN_TEXT"))taskTestBoundary(L"SCAN_PROGRESS");
        if(!strcmp(stage,"RESCAN_TEXT"))taskTestBoundary(L"RESCAN_PROGRESS");
    }if(!guiStageProgress||!stage||!total)return;const auto now=GetTickCount64();
        if(completed==total||now-lastUpdate>=150){std::printf("UI_STAGE_PROGRESS %s %zu %zu\n",stage,completed,total);lastUpdate=now;}}
};

struct Span { long start{}, end{}; };
struct ScannedShape {Span span;std::wstring classType;long type{};};
struct ScannedField {Span codeSpan;long kind{};std::wstring code,result;Span span;};
enum class EquationScope { All, Inline, Display };
inline EquationScope equationScope = EquationScope::All;
inline bool selectedEquation(bool display) {
    return equationScope == EquationScope::All || (display ? equationScope == EquationScope::Display : equationScope == EquationScope::Inline);
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
};
struct ScanResult {
    std::vector<ScanItem> items;
    std::vector<ScannedShape> shapes;
    std::vector<ScannedField> fieldSnapshots;
    long paragraphs{}, existingMathType{}, inlineShapes{}, nativeMath{}, fields{}, revisions{}, controls{}, noteReferences{};
    bool protectedDocument{};
    bool tracksChanges{};
    bool reverse{};
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
inline Span bounds(IDispatch* range) { return {integer(range, L"Start"), integer(range, L"End")}; }
inline Span fieldBounds(IDispatch* field,ScannedField* snapshot=nullptr) {
    auto codeRange=get(field,L"Code"),resultRange=get(field,L"Result");const auto code=bounds(codeRange.Get()),value=bounds(resultRange.Get());
    const long kind=snapshot||value.start==value.end?integer(field,L"Type"):0;
    // 保存后扫描已取得域范围，顺便记录核验所需的代码和缓存，避免再遍历整份域集合。
    if(snapshot){snapshot->codeSpan=code;snapshot->kind=kind;if(kind!=58){snapshot->code=text(codeRange.Get(),L"Text");snapshot->result=text(resultRange.Get(),L"Text");}}
    Span span{code.start-1,value.end+1};
    if(value.start==value.end&&kind==51){
        const auto instruction=snapshot?snapshot->code:text(codeRange.Get(),L"Text");
        if(instruction.find(L" MACROBUTTON MTPlaceRef ")==0||instruction.find(L" MACROBUTTON MTEditEquationSection")==0)
            span.end=code.end+1; // 原生编号没有结果分隔符，避免多占相邻内容。
    }
    if(snapshot)snapshot->span=span;return span;
}

// 所有定位从 Word Find 返回的实际位置取得，绝不把 Range.Text 下标加到 Start。
inline ComPtr<IDispatch> findLiteral(IDispatch* document, long start, long end, const wchar_t* needle) {
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
    require(actual.start >= start && actual.end <= end && actual.end > actual.start && text(range.Get(), L"Text") == needle,
            "Literal Find remains inside requested range");
    return range;
}

inline bool escapedAt(IDispatch* document, long position, long boundary) {
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
inline std::vector<Token> tokens(IDispatch* document, Span segment, const std::wstring& source) {
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

inline bool whitespaceOnly(const std::wstring& value) {
    return std::all_of(value.begin(), value.end(), [](wchar_t c) { return iswspace(c) || c == L'\a'; });
}
// Word 换行只是源码排版空白；TeX 的双反斜杠必须原样交给 MathType。
inline std::wstring normalizedTex(std::wstring value) {
    for (auto& c : value) if (c == L'\r' || c == L'\n' || c == L'\v') c = L' ';
    return value;
}
inline bool structuredMultiline(const std::wstring& value) {
    if (value.find(L"\\\\") != std::wstring::npos) return true;
    for (const auto* name : {L"align",L"aligned",L"split",L"gather",L"gathered",L"array",L"cases",L"matrix",L"pmatrix",L"bmatrix",L"vmatrix",L"Vmatrix",L"smallmatrix"})
        if (value.find(std::wstring(L"\\begin{") + name + L"}") != std::wstring::npos ||
            value.find(std::wstring(L"\\begin{") + name + L"*}") != std::wstring::npos) return true;
    return false;
}
// 只检查基本语法与环境配对，支持度仍由 MathType 的导入结果判定。
inline std::wstring lexicalReason(const std::wstring& inner) {
    if (whitespaceOnly(inner)) return L"EMPTY";
    if (inner.find_first_of(L"\f\a") != std::wstring::npos ||
        (inner.find_first_of(L"\r\n\v") != std::wstring::npos && !structuredMultiline(inner))) return L"UNSUPPORTED_LINE_BREAK";
    long depth{};
    std::vector<std::wstring> environments;
    for (size_t index = 0; index < inner.size(); ++index) {
        const auto character = inner[index];
        if (character == L'\\' && index + 1 < inner.size()) {
            const bool begin = inner.compare(index,7,L"\\begin{") == 0, end = inner.compare(index,5,L"\\end{") == 0;
            if (begin || end) {
                const auto first = index + (begin ? 7 : 5), last = inner.find(L'}',first);
                if (last == std::wstring::npos) return L"UNBALANCED_ENVIRONMENTS";
                const auto name = inner.substr(first,last-first);
                if (begin) environments.push_back(name);
                else { if (environments.empty() || environments.back()!=name) return L"UNBALANCED_ENVIRONMENTS"; environments.pop_back(); }
            }
            ++index; continue;
        }
        if (character == L'{') ++depth;
        if (character == L'}' && --depth < 0) return L"UNBALANCED_BRACES";
    }
    if (depth != 0) return L"UNBALANCED_BRACES";
    return environments.empty() ? L"" : L"UNBALANCED_ENVIRONMENTS";
}

inline void addItem(IDispatch* document, Span span, Span inner, TokenKind kind, const std::wstring& reason,
                    long paragraph, ScanResult& result, const std::wstring& paragraphText) {
    auto range = docRange(document, span.start, span.end);
    ScanItem item;
    item.span = span; item.paragraph = paragraph; item.reason = reason; item.ready = reason.empty();
    item.source = text(range.Get(), L"Text");
    if (inner.end >= inner.start && inner.start >= span.start && inner.end <= span.end)
        item.inner = text(docRange(document, inner.start, inner.end).Get(), L"Text");
    item.delimiter = kind == TokenKind::Dollar ? L"$" : kind == TokenKind::DoubleDollar || kind == TokenKind::EmptyDisplay ? L"$$" :
                     kind == TokenKind::BracketOpen || kind == TokenKind::BracketClose ? L"\\[\\]" :
                     kind == TokenKind::RepairedBracketOpen ? L"[\\]" :
                     kind == TokenKind::ParenOpen || kind == TokenKind::ParenClose ? L"\\(\\)" : L"AMBIGUOUS";
    item.display = kind != TokenKind::Dollar && kind != TokenKind::ParenOpen && kind != TokenKind::ParenClose && kind != TokenKind::AmbiguousDollars;
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
        item.display = position != std::wstring::npos && whitespaceOnly(context.substr(0, position)) &&
            whitespaceOnly(context.substr(position + item.source.size()));
        if (kind == TokenKind::RepairedBracketOpen && !item.display) { item.ready = false; item.reason = L"UNSAFE_MISSING_DELIMITER"; }
    }
    if (!selectedEquation(item.display)) return;
    if(item.ready&&item.display&&!paragraphText.empty()){
        // 独占段落的外侧空白一并替换为单个原生前导 Tab，避免原有多个 Tab 把公式推向右边。
        auto paragraphs=get(range.Get(),L"Paragraphs");auto first=get(collectionItem(paragraphs.Get(),1).Get(),L"Range");auto last=get(collectionItem(paragraphs.Get(),integer(paragraphs.Get(),L"Count")).Get(),L"Range");
        item.span={integer(first.Get(),L"Start"),integer(last.Get(),L"End")-1};item.source=text(docRange(document,item.span.start,item.span.end).Get(),L"Text");
    }
    item.inner = normalizedTex(item.inner);
    const auto font = get(range.Get(), L"Font");
    // 混合字号读取公式内容首字符，避免 Word 的 wdUndefined 被当成真实字号。
    auto size = invoke(font.Get(),L"Size",DISPATCH_PROPERTYGET);
    item.points = size.v.vt == VT_R4 ? size.v.fltVal : size.v.vt == VT_R8 ? size.v.dblVal : size.v.vt == VT_I4 ? size.v.lVal : 0;
    if (!(item.points >= 1 && item.points <= 256) && inner.start >= 0 && inner.end > inner.start) {
        auto characters=get(docRange(document,inner.start,inner.end).Get(),L"Characters");
        for(long i=1;i<=integer(characters.Get(),L"Count");++i) {
            auto character=collectionItem(characters.Get(),i); if(whitespaceOnly(text(character.Get(),L"Text")))continue;
            auto effective=invoke(get(character.Get(),L"Font").Get(),L"Size",DISPATCH_PROPERTYGET);
            item.points=effective.v.vt==VT_R4?effective.v.fltVal:effective.v.vt==VT_R8?effective.v.dblVal:effective.v.vt==VT_I4?effective.v.lVal:0;break;
        }
    }
    result.items.push_back(std::move(item));
}

inline long scanSegment(IDispatch* document, Span segment, long paragraph, ScanResult& result, long boundary, const std::wstring& paragraphText) {
    const auto source=text(docRange(document,segment.start,segment.end).Get(),L"Text");
    if (!hasTexMarkers(source)) return segment.end;
    Token opening{};
    bool haveOpening = false;
    for (const auto& token : tokens(document, segment, source)) {
        if (token.kind == TokenKind::AmbiguousDollars || token.kind == TokenKind::EmptyDisplay) {
            addItem(document, {haveOpening ? opening.span.start : token.span.start, token.span.end}, {-1, -1},
                    token.kind, token.kind == TokenKind::EmptyDisplay && !haveOpening ? L"EMPTY" : L"AMBIGUOUS_DOLLARS", paragraph, result, paragraphText);
            haveOpening = false; continue;
        }
        if (!haveOpening) {
            if (token.kind == TokenKind::BracketClose || token.kind == TokenKind::ParenClose)
                addItem(document, token.span, {-1, -1}, token.kind, L"UNMATCHED_CLOSE", paragraph, result, paragraphText);
            else { opening = token; haveOpening = true; }
            continue;
        }
        const bool closes = tokensMatch(opening.kind, token.kind);
        if (closes) {
            const Span inner{opening.span.end, token.span.start};
            addItem(document, {opening.span.start, token.span.end}, inner, opening.kind,
                    lexicalReason(text(docRange(document, inner.start, inner.end).Get(), L"Text")), paragraph, result, paragraphText);
        } else {
            // 混用定界符时消费冲突标记，保守拒绝，不生成互相覆盖的可转换范围。
            addItem(document, {opening.span.start, token.span.end}, {-1, -1}, opening.kind, L"MIXED_DELIMITERS", paragraph, result, paragraphText);
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
                addItem(document,{opening.span.start,closing.end},inner,opening.kind,lexicalReason(body),paragraph,result,paragraphText);return closing.end;
            }
        }
        addItem(document, {opening.span.start, segment.end}, {-1, -1}, opening.kind, L"UNMATCHED_OPEN", paragraph, result, paragraphText);
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
        const bool number=field.kind==51&&field.code.find(L"MACROBUTTON MTPlaceRef")!=std::wstring::npos;
        const bool counter=field.kind==12&&(field.code.find(L"SEQ MTEqn ")!=std::wstring::npos||field.code.find(L"SEQ MTChap ")!=std::wstring::npos||field.code.find(L"SEQ MTSec ")!=std::wstring::npos);
        if(field.kind!=58&&!number&&!counter)return false;
        excluded.push_back(field.span);
    }
    std::sort(excluded.begin(),excluded.end(),[](const Span& a,const Span& b){return a.start<b.start;});
    long cursor=para.start;
    for(const auto& span:excluded){if(span.start>cursor&&!whitespaceOnly(text(docRange(document,cursor,span.start).Get(),L"Text")))return false;cursor=std::max(cursor,span.end);}
    if(cursor<para.end&&!whitespaceOnly(text(docRange(document,cursor,para.end).Get(),L"Text")))return false;
    ScanItem item;item.span=para;item.paragraph=paragraph;item.source=text(range,L"Text");item.display=true;item.existingDisplay=true;item.existingNativeOrdinal=candidateOrdinal;item.delimiter=L"EXISTING_DISPLAY";
    result.items.push_back(std::move(item));return true;
}

inline ScanResult scanDocumentText(IDispatch* document, bool paragraphContext,bool captureFields=false,bool includeExisting=false) {
    if(checkingSavedDocument)std::puts("TIMING_STAGE RESCAN_BOUNDARIES");
    ScanResult result;
    const auto tracking = invoke(document, L"TrackRevisions", DISPATCH_PROPERTYGET);
    require(tracking.v.vt == VT_BOOL, "Revision tracking property type");
    result.tracksChanges = tracking.v.boolVal != VARIANT_FALSE;
    result.protectedDocument = integer(document, L"ProtectionType") != -1;
    if (result.protectedDocument) return result;
    auto content = get(document, L"Content");
    require(integer(content.Get(), L"StoryType") == 1, "Main story only");
    std::vector<Span> excluded,protectedRanges;
    auto exclude = [&](IDispatch* range,bool protect=true) {
        require(integer(range, L"StoryType") == 1, "Excluded main story range");
        excluded.push_back(bounds(range));
        if(protect)protectedRanges.push_back(excluded.back());
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
    for (const auto& entry : {std::make_pair(L"OMaths", &result.nativeMath), std::make_pair(L"Revisions", &result.revisions),
                              std::make_pair(L"ContentControls", &result.controls)}) {
        auto collection = get(content.Get(), entry.first);
        *entry.second = integer(collection.Get(), L"Count");
        for (long index = 1; index <= *entry.second; ++index) exclude(get(collectionItem(collection.Get(), index).Get(), L"Range").Get());
    }
    auto fields = get(content.Get(), L"Fields");
    result.fields = integer(fields.Get(), L"Count");
    captureFields=captureFields||(includeExisting&&result.existingMathType>0);
    if(captureFields)result.fieldSnapshots.resize(static_cast<size_t>(result.fields));
    for (long index = 1; index <= result.fields; ++index) {
        auto field = collectionItem(fields.Get(), index);
        excluded.push_back(fieldBounds(field.Get(),captureFields?&result.fieldSnapshots[static_cast<size_t>(index-1)]:nullptr));
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
    if(checkingSavedDocument)std::puts("TIMING_STAGE RESCAN_PARAGRAPHS");
    auto paragraphs = get(content.Get(), L"Paragraphs");
    result.paragraphs = integer(paragraphs.Get(), L"Count");
    std::vector<ComPtr<IDispatch>> savedParagraphs;
    if(checkingSavedDocument){
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
    StageProgress paragraphProgress(checkingSavedDocument?"RESCAN_TEXT":"SCAN_TEXT",static_cast<size_t>(result.paragraphs));
    const long contentEnd=integer(content.Get(),L"End");long consumed{};size_t excludedIndex{};
    for (long index = 1; index <= result.paragraphs; ++index) {
        auto paragraph = checkingSavedDocument?savedParagraphs[static_cast<size_t>(index-1)]:collectionItem(paragraphs.Get(), index);
        auto range = get(paragraph.Get(), L"Range");
        auto paragraphSpan = bounds(range.Get());
        if(includeExisting&&result.existingMathType>0&&selectedEquation(true)&&addExistingDisplay(document,range.Get(),index,result,protectedRanges)){paragraphProgress.advance();continue;}
        paragraphSpan.start=std::max(paragraphSpan.start,consumed);
        while(excludedIndex<excluded.size()&&excluded[excludedIndex].end<=paragraphSpan.start)++excludedIndex;
        if(paragraphSpan.start>=paragraphSpan.end || (excludedIndex<excluded.size()&&excluded[excludedIndex].start<=paragraphSpan.start&&excluded[excludedIndex].end>=paragraphSpan.end)) {paragraphProgress.advance();continue;}
        const auto paragraphText=text(range.Get(),L"Text");
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
            if (block.start > cursor) consumed=std::max(consumed,scanSegment(document, {cursor, std::min(block.start, paragraphSpan.end)}, index, result,block.start,paragraphContext?paragraphText:L""));
            cursor = std::max(cursor, std::min(block.end, paragraphSpan.end));
        }
        if (cursor < paragraphSpan.end) {
            long boundary=contentEnd;
            for(size_t i=excludedIndex;i<excluded.size();++i)if(excluded[i].start>=cursor){boundary=excluded[i].start;break;}
            consumed=std::max(consumed,scanSegment(document, {cursor, paragraphSpan.end}, index, result,boundary,paragraphContext?paragraphText:L""));
        }
        paragraphProgress.advance();
    }
    std::sort(result.items.begin(), result.items.end(), [](const ScanItem& a, const ScanItem& b) { return a.span.start < b.span.start; });
    long previousEnd = -1;
    if(checkingSavedDocument)std::puts("TIMING_STAGE RESCAN_RANGES");
    StageProgress rangeProgress(checkingSavedDocument?"RESCAN_RANGES":"SCAN_RANGES",result.items.size());
    for (const auto& item : result.items) {
        require(item.span.start >= previousEnd && item.span.end > item.span.start, "Non-overlapping scan items");
        require(text(docRange(document, item.span.start, item.span.end).Get(), L"Text") == item.source, "All reported ranges rechecked");
        previousEnd = item.span.end;
        rangeProgress.advance();
    }
    return result;
}
inline ScanResult scanDocument(IDispatch* document) { return scanDocumentText(document, true); }

inline std::string utf8(const std::wstring& value) {
    if (value.empty()) return {};
    require(value.size() <= INT_MAX, "UTF-8 input bounds");
    const int count = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, value.data(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    require(count > 0, "UTF-8 length");
    std::string output(static_cast<size_t>(count), '\0');
    require(WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, value.data(), static_cast<int>(value.size()), output.data(), count, nullptr, nullptr) == count, "UTF-8 encoding");
    return output;
}
inline std::string quote(const std::wstring& value) {
    std::string output = "\"";
    for (const unsigned char c : utf8(value)) {
        if (c == '"' || c == '\\') { output += '\\'; output += static_cast<char>(c); }
        else if (c < 32) {
            char escaped[7]{}; sprintf_s(escaped, "\\u%04X", static_cast<unsigned int>(c)); output += escaped;
        } else output += static_cast<char>(c);
    }
    return output + '"';
}
inline void writeReport(const std::wstring& path, const std::wstring& input, const ScanResult& result) {
    const auto ready = std::count_if(result.items.begin(), result.items.end(), [](const ScanItem& item) { return item.ready; });
    std::string json = "{\n  \"Input\": " + quote(input) + ",\n  \"Scope\": \"MAIN_STORY_BODY_ONLY\",\n";
    json += "  \"Status\": \"" + std::string(result.protectedDocument ? "SKIPPED_PROTECTED_DOCUMENT" : "SCANNED") + "\",\n";
    json += "  \"SchemaVersion\": 1, \"ReadOnly\": true, \"Positions\": \"WORD_MAIN_STORY_SNAPSHOT\",\n";
    if (result.reverse) json += "  \"Direction\":\"REVERSE\", \"MathTypeTypeSource\":\"READ_ONLY_EQUATION_NATIVE\",\n";
    json += "  \"Validation\": \"BOUNDARIES_AND_BASIC_SYNTAX_ONLY\", \"MustRescanBeforeConvert\": true,\n";
    // 整篇保护时没有扫描，计数为未知，不能以 0 冒充检测结果。
    auto count = [&](size_t value) { return result.protectedDocument ? std::string("null") : std::to_string(value); };
    json += "  \"Ready\": " + count(static_cast<size_t>(ready)) + ", \"Rejected\": " + count(result.items.size() - static_cast<size_t>(ready)) + ",\n";
    json += "  \"ExistingMathType\": " + count(result.existingMathType) + ", \"InlineShapes\": " + count(result.inlineShapes) + ",\n";
    json += "  \"Excluded\": {\"OMML\":" + count(result.nativeMath) + ",\"Fields\":" + count(result.fields) +
            ",\"Revisions\":" + count(result.revisions) + ",\"ContentControls\":" + count(result.controls) +
            ",\"NoteReferences\":" + count(result.noteReferences) + "},\n  \"Items\": [\n";
    for (size_t index = 0; index < result.items.size(); ++index) {
        const auto& item = result.items[index];
        json += "    {\"Id\":" + std::to_string(index + 1) + ",\"Start\":" + std::to_string(item.span.start) + ",\"End\":" + std::to_string(item.span.end) +
            ",\"Paragraph\":" + std::to_string(item.paragraph) + ",\"Region\":\"" + (item.row ? std::string("TABLE") : std::string("BODY")) +
            "\",\"TableStart\":" + std::to_string(item.tableStart) + ",\"Row\":" + std::to_string(item.row) + ",\"Column\":" + std::to_string(item.column) +
            ",\"Delimiter\":" + quote(item.delimiter) + ",\"Display\":" + (item.display ? "true" : "false") +
            ",\"Status\":\"" + (item.ready ? "READY" : "REJECTED") + "\",\"Reason\":" + quote(item.reason) +
            ",\"Source\":" + quote(item.source) + ",\"Inner\":" + quote(item.inner) + "}" + (index + 1 == result.items.size() ? "\n" : ",\n");
    }
    json += "  ]\n}\n";
    // CREATE_NEW 防止覆盖任何已存在的清单；关闭失败或短写入不宣布通过。
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(file != INVALID_HANDLE_VALUE, "Create new scan report");
    taskTrackFile(path, file, true);
    DWORD written{};
    const bool saved = json.size() <= MAXDWORD && WriteFile(file, json.data(), static_cast<DWORD>(json.size()), &written, nullptr) && written == json.size();
    const bool closed = CloseHandle(file) != 0;
    if (!saved || !closed) { DeleteFileW(path.c_str()); throw std::runtime_error("Write complete scan report"); }
}

inline void requireDocxZip(const std::wstring& path) {
    FILE* file{};
    const auto opened=_wfopen_s(&file,path.c_str(),L"rb");
    if(opened!=0||!file){unsigned long systemError{};_get_doserrno(&systemError);throw std::runtime_error("Open DOCX input (errno="+std::to_string(opened)+" Win32="+std::to_string(systemError)+")");}
    unsigned char prefix[4]{};
    const bool zip = fread(prefix, 1, 4, file) == 4 && prefix[0] == 'P' && prefix[1] == 'K' && prefix[2] == 3 && prefix[3] == 4;
    const bool closed = fclose(file) == 0;
    // 加密 DOCX 通常是 CFB 容器，提前拒绝，避免 Word 弹密码或格式转换对话框。
    require(zip && closed, "Unencrypted DOCX ZIP required");
}

inline std::wstring absolutePath(const wchar_t* path) {
    std::vector<wchar_t> buffer(32768);
    const DWORD length = GetFullPathNameW(path, static_cast<DWORD>(buffer.size()), buffer.data(), nullptr);
    require(length > 0 && length < buffer.size(), "Resolve full input/output path");
    return std::wstring(buffer.data(), length);
}

// GUI 与命令行共用同一只读扫描；调用方所在工作进程必须先初始化 STA。
inline ScanResult scanReadOnlyFile(const std::wstring& input, const std::wstring& report,
                                  ScanResult (*scanner)(IDispatch*) = scanDocument) {
    require(input.size() > 5 && _wcsicmp(input.c_str() + input.size() - 5, L".docx") == 0, "DOCX input required");
    require(GetFileAttributesW(report.c_str()) == INVALID_FILE_ATTRIBUTES, "Report already exists");
    const auto copy = report + L".scan-" + std::to_wstring(GetCurrentProcessId()) + L".docx";
    HANDLE source = CreateFileW(input.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    require(source != INVALID_HANDLE_VALUE, "Lock source during read-only scan");
    bool owned = false;
    try {
        // 在只读共享锁内检查输入，避免检查后被替换成加密文件而触发 Word 对话框。
        requireDocxZip(input);
        taskCopy(input, copy); owned = true;setDocxLinkUpdates(copy);
        WordSession word;
        auto documents = get(word.application.Get(), L"Documents");
        word.document = asObject(invoke(documents.Get(), L"Add", DISPATCH_METHOD)); word.verifyOwnership();
        invoke(word.application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
        invoke(word.application.Get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0L});
        invoke(word.application.Get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3L});
        invoke(word.document.Get(), L"Close", DISPATCH_METHOD, {0L}); word.document.Reset();
        taskTrackWordOwner(copy);
        word.document = asObject(invoke(documents.Get(), L"Open", DISPATCH_METHOD, {copy.c_str(), false, true, false}));
        const auto readOnly = invoke(word.document.Get(), L"ReadOnly", DISPATCH_PROPERTYGET);
        require(readOnly.v.vt == VT_BOOL && readOnly.v.boolVal != VARIANT_FALSE, "Scan document opened read-only");
        const auto before = text(get(word.document.Get(), L"Content").Get(), L"Text");
        auto result = scanner(word.document.Get());
        require(text(get(word.document.Get(), L"Content").Get(), L"Text") == before, "Scanning leaves Word text unchanged");
        word.shutdown();
        require(DeleteFileW(copy.c_str()) != 0, "Remove owned scan copy"); owned = false;
        require(CloseHandle(source) != 0, "Release scan source lock"); source = INVALID_HANDLE_VALUE;
        const auto staged=report+L".work-"+std::to_wstring(GetCurrentProcessId());taskReservePrivatePath(staged);
        writeReport(staged, input, result);taskMove(staged,report);
        taskCommit();
        return result;
    } catch (...) {
        if (owned && !DeleteFileW(copy.c_str())) std::puts("OWNED_SCAN_COPY_CLEANUP_FAILED");
        if (source != INVALID_HANDLE_VALUE) CloseHandle(source);
        throw;
    }
}

// 工作进程只发送 UTF-8 文本；十六进制字段避免源式中的控制符破坏行协议。
inline std::string wireHex(const std::string& bytes) {
    if (bytes.empty()) return "-";
    constexpr char digits[] = "0123456789ABCDEF";
    std::string result;
    for (const unsigned char value : bytes) { result += digits[value >> 4]; result += digits[value & 15]; }
    return result;
}

// GUI 只接收统计，公式原文和定位始终留在转换进程中。
inline void emitGuiScan(const ScanResult& result) {
    const auto ready=std::count_if(result.items.begin(),result.items.end(),[](const ScanItem& item){return item.ready;});
    const auto inlines=std::count_if(result.items.begin(),result.items.end(),[](const ScanItem& item){return item.ready&&!item.display;});
    const auto rejected=std::count_if(result.items.begin(),result.items.end(),[](const ScanItem& item){return !item.ready&&!item.existingDisplay&&!item.existingReference;});
    std::printf("UI_SCAN_COUNTS %ld %td %td %td\n",result.existingMathType,ready,inlines,rejected);
}
