// 实验分支只替换公式，不读取或应用正文的字号、段落、编号设置。
#include "experimental_conversion.h"
#include "omml_adapter.h"
#include "equation_objects.h"
#include "mathtype_api.h"
#include "conversion_timing.h"
#include "layout_warnings.h"
#include <map>
#include <memory>
#include <set>
#include <string_view>

namespace {
struct Region {std::wstring key,label;ComPtr<IDispatch> range;bool main{};std::vector<Span> tables;ComPtr<IDispatch> owner;};
struct Target {ScanItem item;std::wstring kind;};
struct Snapshot {std::wstring key,body;std::vector<std::wstring> layout,maths,fields;std::vector<std::vector<unsigned char>> native;};
struct Binding {std::wstring key;ComPtr<IDispatch> owner;Span span;std::vector<unsigned char> native;std::wstring output;size_t itemIndex{};};
struct HeaderSlot {long section,kind;bool header;};
struct DocumentState {std::vector<std::wstring> layout,maths;};
struct RegionError : std::runtime_error {
    ProcessingIssue issue;
    std::vector<ExperimentalItem> preserved;
    RegionError(const std::wstring& key,const std::wstring& label,const std::wstring& stage,const std::exception& error)
        :std::runtime_error(error.what()),issue{key,label,stage,{}}{
        const std::string reason=error.what();issue.reason.assign(reason.begin(),reason.end());
        issue.position.location=L"原文页码未知 · "+label;
    }
};
// 测试只用明确的环境变量注入边界故障，不依赖用户公式内容。
bool regionFault(const std::wstring& key,const wchar_t* stage) {
    return conversionFault(L"DOCX_MATH_REGION_FAULT",key,stage);
}
bool boolean(IDispatch* object,const wchar_t* name) {auto value=invoke(object,name,DISPATCH_PROPERTYGET);require(value.v.vt==VT_BOOL,"Experimental Boolean property");return value.v.boolVal!=VARIANT_FALSE;}
bool inTable(const Region& region,long position) {return std::any_of(region.tables.begin(),region.tables.end(),[&](Span span){return position>=span.start&&position<span.end;});}

std::vector<std::wstring> xmlValues(IXMLDOMDocument2* dom,const wchar_t* query) {
    ComPtr<IXMLDOMNodeList> nodes;check(dom->selectNodes(Value(query).v.bstrVal,&nodes),"Select experimental XML");long count{};check(nodes->get_length(&count),"Experimental XML count");
    std::vector<std::wstring> result;for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> node;check(nodes->get_item(i,&node),"Experimental XML node");Value xml;xml.v.vt=VT_BSTR;check(node->get_xml(&xml.v.bstrVal),"Experimental node XML");result.emplace_back(xml.v.bstrVal);}return result;
}
// 一次 WordOpenXML 导出同时取得段落/表格版式和 OMML，避免同一区域导出两次整包样式。
DocumentState documentState(IDispatch* range,ConversionTimings& timings,bool captureMaths=true) {
    ScopedConversionTiming timing(timings,ConversionMetric::ExperimentalRegionState);
    auto dom=xmlDocument(text(range,L"WordOpenXML"));
    // 页眉部件引用由 Word 重命名或补齐，不属于段落版式；内容另按节与类型核验。
    ComPtr<IXMLDOMNodeList> references;check(dom->selectNodes(Value(L"//*[local-name()='body']//*[local-name()='sectPr']/*[local-name()='headerReference' or local-name()='footerReference']").v.bstrVal,&references),"Layout metadata references");
    long count{};check(references->get_length(&count),"Layout reference count");for(long i=count;i-- >0;){ComPtr<IXMLDOMNode> node,parent,removed;check(references->get_item(i,&node),"Layout reference");check(node->get_parentNode(&parent),"Layout reference owner");check(parent->removeChild(node.Get(),&removed),"Ignore header reference identifiers in layout snapshot");}
    DocumentState state;
    state.layout=xmlValues(dom.Get(),L"//*[local-name()='body']//*[namespace-uri()='http://schemas.openxmlformats.org/wordprocessingml/2006/main' and (local-name()='pPr' or local-name()='tblPr' or local-name()='trPr' or local-name()='tcPr')]");
    for(auto& value:state.layout)value=withoutWordRevisionIds(std::move(value));
    if(captureMaths)state.maths=xmlValues(dom.Get(),L"//*[local-name()='body']//*[namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/math' and local-name()='oMath']");
    return state;
}
std::vector<std::wstring> layoutSnapshot(IDispatch* range,ConversionTimings& timings) {return documentState(range,timings,false).layout;}
Region* addRegion(std::vector<Region>& regions,const std::wstring& key,const std::wstring& label,IDispatch* range,bool main=false) {
    for(const auto& prior:regions){auto same=invoke(range,L"IsEqual",DISPATCH_METHOD,{prior.range.Get()});if(same.v.vt==VT_BOOL&&same.v.boolVal)return nullptr;}
    regions.push_back({key,label,get(range,L"Duplicate"),main});
    return &regions.back();
}
void textboxes(std::vector<Region>& regions,IDispatch* shapes,const std::wstring& owner,const std::wstring& label) {
    // 只读遍历期间集合不变，数量在进入循环时读取一次。
    const long count=integer(shapes,L"Count");
    for(long i=1;i<=count;++i){taskPulse();auto shape=collectionItem(shapes,i);const long type=integer(shape.Get(),L"Type");
        if(type==6){textboxes(regions,get(shape.Get(),L"GroupItems").Get(),owner+L":group:"+std::to_wstring(i),label+L"组合 "+std::to_wstring(i)+L" ");continue;}
        if(type!=1&&type!=17)continue;auto frame=get(shape.Get(),L"TextFrame");if(!integer(frame.Get(),L"HasText"))continue;
        // Word 保存时会重新分配 Shape.ID；使用所属区域内的对象路径并核对完整内容。
        auto range=get(frame.Get(),L"ContainingRange");if(auto region=addRegion(regions,owner+L":box:"+std::to_wstring(i),label+L"文本框 "+std::to_wstring(i),range.Get()))region->owner=shape;
    }
}
// WordOpenXML 会导出虚拟空槽位；从尚未打开的 DOCX 读取实际引用，按节和类型定位。
std::vector<HeaderSlot> storedHeaders(const std::wstring& path) {
    ComPtr<IOpcFactory> factory;check(CoCreateInstance(__uuidof(OpcFactory),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&factory)),"Header reference package factory");
    ComPtr<IStream> input;check(factory->CreateStreamOnFile(path.c_str(),OPC_STREAM_IO_READ,nullptr,0,&input),"Read stored header references");
    ComPtr<IOpcPackage> package;check(factory->ReadPackageFromStream(input.Get(),OPC_VALIDATE_ON_LOAD,&package),"Read header reference package");
    ComPtr<IOpcPartSet> parts;check(package->GetPartSet(&parts),"Header reference parts");ComPtr<IOpcPartUri> uri;
    check(factory->CreatePartUri(L"/word/document.xml",&uri),"Document part URI");ComPtr<IOpcPart> part;check(parts->GetPart(uri.Get(),&part),"Stored document part");
    ComPtr<IStream> content;check(part->GetContentStream(&content),"Stored section XML");auto dom=xmlDocument(L"<root/>");
    Value source;source.v.vt=VT_UNKNOWN;source.v.punkVal=content.Get();source.v.punkVal->AddRef();VARIANT_BOOL loaded{};
    check(dom->load(source.v,&loaded),"Read stored section properties");require(loaded!=VARIANT_FALSE,"Valid stored section XML");
    ComPtr<IXMLDOMNodeList> properties;check(dom->selectNodes(Value(L"/*[local-name()='document']/*[local-name()='body']//*[local-name()='sectPr' and not(ancestor::*[local-name()='sectPrChange']) and not(ancestor::*[local-name()='txbxContent'])]").v.bstrVal,&properties),"Stored section references");
    long count{};check(properties->get_length(&count),"Stored section count");std::vector<HeaderSlot> result;
    for(long section=0;section<count;++section){ComPtr<IXMLDOMNode> property;check(properties->get_item(section,&property),"Stored section");ComPtr<IXMLDOMNodeList> references;
        check(property->selectNodes(Value(L"*[local-name()='headerReference' or local-name()='footerReference']").v.bstrVal,&references),"Stored header and footer references");long referenceCount{};check(references->get_length(&referenceCount),"Stored header reference count");
        for(long i=0;i<referenceCount;++i){ComPtr<IXMLDOMNode> reference,type;check(references->get_item(i,&reference),"Stored header reference");Value name,value;name.v.vt=value.v.vt=VT_BSTR;
            check(reference->get_baseName(&name.v.bstrVal),"Stored header kind");check(reference->selectSingleNode(Value(L"@*[local-name()='type']").v.bstrVal,&type),"Stored header type");require(type!=nullptr,"Stored header reference has type");check(type->get_text(&value.v.bstrVal),"Stored header variant");
            result.push_back({section+1,!wcscmp(value.v.bstrVal,L"default")?1:!wcscmp(value.v.bstrVal,L"first")?2:3,!wcscmp(name.v.bstrVal,L"headerReference")});
        }
    }
    return result;
}
std::vector<Region> collectRegions(IDispatch* document,const ExperimentalOptions& options,const std::vector<HeaderSlot>& headers) {
    std::vector<Region> regions;addRegion(regions,L"body",L"正文",get(document,L"Content").Get(),true);
    // 正文只作保留核验；每个顶层表格单独处理，嵌套表格由其所属顶层表格覆盖。
    if(options.tables){auto tables=get(document,L"Tables");const long count=integer(tables.Get(),L"Count");for(long i=1;i<=count;++i){auto range=get(collectionItem(tables.Get(),i).Get(),L"Range");
        bool nested=false;for(const auto& prior:regions)if(prior.key.rfind(L"body:table:",0)==0){const auto a=bounds(prior.range.Get()),b=bounds(range.Get());if(a.start<=b.start&&a.end>=b.end){nested=true;break;}}
        if(!nested)addRegion(regions,L"body:table:"+std::to_wstring(i),L"正文表格 "+std::to_wstring(i),range.Get());}}
    if(options.textboxes)textboxes(regions,get(document,L"Shapes").Get(),L"body",L"正文");
    if(options.headers){
        auto sections=get(document,L"Sections");
        for(const auto& slot:headers){auto item=collectionItem(sections.Get(),slot.section);const wchar_t* name=slot.header?L"Headers":L"Footers";
            auto header=collectionItem(get(item.Get(),name).Get(),slot.kind);if(slot.section>1&&boolean(header.Get(),L"LinkToPrevious"))continue;
            const auto key=std::wstring(name)+L":"+std::to_wstring(slot.section)+L":"+std::to_wstring(slot.kind);
            const auto label=L"第 "+std::to_wstring(slot.section)+L" 节"+(slot.header?L"页眉":L"页脚")+(slot.kind==1?L"（主页）":slot.kind==2?L"（首页）":L"（偶数页）");
            addRegion(regions,key,label,get(header.Get(),L"Range").Get());if(options.textboxes)textboxes(regions,get(header.Get(),L"Shapes").Get(),key,label);
        }
    }
    if(options.notes)for(const auto* name:{L"Footnotes",L"Endnotes"}){auto notes=get(document,name);const long count=integer(notes.Get(),L"Count");
        for(long i=1;i<=count;++i){auto note=collectionItem(notes.Get(),i);if(auto region=addRegion(regions,std::wstring(name)+L":"+std::to_wstring(i),std::wstring(!wcscmp(name,L"Footnotes")?L"脚注 ":L"尾注 ")+std::to_wstring(i),get(note.Get(),L"Range").Get()))region->owner=note;}
    }
    return regions;
}
bool locationSelected(const Region& region,IDispatch* range,const ExperimentalOptions& options) {
    const bool table=inTable(region,integer(range,L"Start"));
    return !region.main&&(!table||options.tables);
}
bool intersects(Span span,const std::vector<Span>& excluded) {
    return std::any_of(excluded.begin(),excluded.end(),[&](Span block){return block.start<span.end&&block.end>span.start;});
}
std::vector<Target> scanRegion(IDispatch* document,const Region& region,const ExperimentalOptions& options,bool reverse,const ScanOptions& scanOptions) {
    const auto root=region.range.Get();const auto limits=bounds(root);std::vector<Span> protectedRanges,excluded;std::vector<Target> targets;
    for(const auto* name:{L"Revisions",L"ContentControls"}){auto collection=get(root,name);const long count=integer(collection.Get(),L"Count");for(long i=1;i<=count;++i){auto span=bounds(get(collectionItem(collection.Get(),i).Get(),L"Range").Get());
        if(!wcscmp(name,L"ContentControls")){span.start=std::max(limits.start,span.start-1);span.end=std::min(limits.end,span.end+1);}protectedRanges.push_back(span);}}
    auto fields=get(root,L"Fields");const long fieldCount=integer(fields.Get(),L"Count");for(long i=1;i<=fieldCount;++i){auto field=collectionItem(fields.Get(),i);if(integer(field.Get(),L"Type")!=58)protectedRanges.push_back(fieldBounds(field.Get()));}
    excluded=protectedRanges;
    auto maths=get(root,L"OMaths");const long mathCount=integer(maths.Get(),L"Count");for(long i=1;i<=mathCount;++i){auto math=collectionItem(maths.Get(),i);auto range=get(math.Get(),L"Range");const auto span=bounds(range.Get());excluded.push_back(span);
        if(reverse||!scanOptions.convertOmml||!locationSelected(region,range.Get(),options)||intersects(span,protectedRanges))continue;
        ScanItem item;item.span=span;item.position.sourceStart=span.start;item.source=text(range.Get(),L"Text");item.display=integer(math.Get(),L"Type")==0;item.ready=true;
        if(selectedEquation(item.display,scanOptions.scope))targets.push_back({std::move(item),L"OMML"});
    }
    auto shapes=get(root,L"InlineShapes");const long shapeCount=integer(shapes.Get(),L"Count");for(long i=1;i<=shapeCount;++i){auto shape=collectionItem(shapes.Get(),i);auto range=get(shape.Get(),L"Range");const auto span=bounds(range.Get());excluded.push_back(span);
        if(!reverse||!locationSelected(region,range.Get(),options)||intersects(span,protectedRanges))continue;
        const auto type=integer(shape.Get(),L"Type");if(type!=1&&type!=2)continue;const auto cls=text(get(shape.Get(),L"OLEFormat").Get(),L"ClassType");if(cls.find(L"Equation.DSMT")!=0)continue;
        ScanItem item;item.span=span;item.position.sourceStart=span.start;item.source=text(range.Get(),L"Text");item.ready=type==1&&cls==L"Equation.DSMT4";
        if(item.ready){try{item.display=!inspectStyle(nativePayload(rangeNativeAsset(range.Get()).equation)).inlineEquation;}
            catch(const std::exception&){item.ready=false;item.reason=L"UNREADABLE_MATHTYPE_NATIVE";}}
        else item.reason=L"UNSUPPORTED_MATHTYPE_CLASS";
        if(!item.ready||selectedEquation(item.display,scanOptions.scope))targets.push_back({std::move(item),L"MATH_TYPE"});
    }
    if(reverse)return targets;
    // 注释引用和单元格末尾作为边界，LaTeX 识别复用正文词法器。
    if(region.main)for(const auto* name:{L"Footnotes",L"Endnotes"}){auto notes=get(document,name);const long count=integer(notes.Get(),L"Count");for(long i=1;i<=count;++i)excluded.push_back(bounds(get(collectionItem(notes.Get(),i).Get(),L"Reference").Get()));}
    auto paragraphs=get(root,L"Paragraphs");const long paragraphCount=integer(paragraphs.Get(),L"Count");
    for(long i=1;i<=paragraphCount;++i){auto range=get(collectionItem(paragraphs.Get(),i).Get(),L"Range");auto span=bounds(range.Get());
        if(!locationSelected(region,range.Get(),options))excluded.push_back(span);
        if(inTable(region,span.start)&&locationSelected(region,range.Get(),options)){auto cells=get(range.Get(),L"Cells");
            // Word 的行末段落仍报告在表格内，但不属于任何单元格。
            if(!integer(cells.Get(),L"Count")){excluded.push_back(span);continue;}
            auto cell=get(collectionItem(cells.Get(),1).Get(),L"Range");const long end=integer(cell.Get(),L"End");excluded.push_back({end-1,end});}
    }
    std::sort(excluded.begin(),excluded.end(),[](Span a,Span b){return a.start<b.start;});std::vector<Span> blocks;
    for(auto span:excluded){span.start=std::max(span.start,limits.start);span.end=std::min(span.end,limits.end);if(span.end<=span.start)continue;if(!blocks.empty()&&span.start<=blocks.back().end)blocks.back().end=std::max(blocks.back().end,span.end);else blocks.push_back(span);}
    ScanResult scan;long consumed=limits.start;ScanRangeContext context{document,root,false,scanOptions};
    for(long i=1;i<=paragraphCount;++i){taskPulse();auto range=get(collectionItem(paragraphs.Get(),i).Get(),L"Range");auto span=bounds(range.Get());span.start=std::max({span.start,limits.start,consumed});span.end=std::min(span.end,limits.end);
        if(span.end<=span.start||!locationSelected(region,range.Get(),options))continue;const auto paragraphText=text(range.Get(),L"Text");if(!hasTexMarkers(paragraphText))continue;
        while(span.end>span.start){const auto last=text(sliceRange(root,span.end-1,span.end).Get(),L"Text");if(last.empty()||!std::all_of(last.begin(),last.end(),[](wchar_t c){return c==L'\r'||c==L'\a';}))break;--span.end;}
        long cursor=span.start,boundary=limits.end;
        for(const auto block:blocks){if(block.end<=cursor)continue;if(block.start>span.end){boundary=block.start;break;}if(block.start>cursor)consumed=std::max(consumed,scanSegment(context,{cursor,std::min(block.start,span.end)},i,scan,block.start,paragraphText));cursor=std::max(cursor,block.end);}
        if(cursor<span.end)consumed=std::max(consumed,scanSegment(context,{cursor,span.end},i,scan,boundary,paragraphText));
    }
    for(auto& item:scan.items){
        // 实验替换只包含定界符与公式本身，保留段首段尾的空白和制表符。
        if(item.ready){const auto markers=tokens(context,item.span,item.source);require(markers.size()>=2,"Experimental formula delimiters");item.span={markers.front().span.start,markers.back().span.end};item.source=text(sliceRange(root,item.span.start,item.span.end).Get(),L"Text");}
        targets.push_back({std::move(item),L"LATEX"});
    }
    return targets;
}
void readRegionTables(Region& region) {
    auto tables=get(region.range.Get(),L"Tables");const long count=integer(tables.Get(),L"Count");
    for(long i=1;i<=count;++i)region.tables.push_back(bounds(get(collectionItem(tables.Get(),i).Get(),L"Range").Get()));
}
SourcePosition regionSourcePosition(IDispatch* document,const Region& region,bool paginated) {
    SourcePosition result;result.location=L"原文页码未知 · "+region.label;
    // 页眉页脚的同一内容可出现在多个页面，不给它编造一个唯一页码。
    if(region.key.rfind(L"Headers:",0)==0||region.key.rfind(L"Footers:",0)==0){result.location=region.label+L"（共用内容）";return result;}
    try{
        if(region.key.rfind(L"Footnotes:",0)==0||region.key.rfind(L"Endnotes:",0)==0){
            auto reference=get(region.owner.Get(),L"Reference");result=originalSourcePosition(document,integer(reference.Get(),L"Start"),region.label,paginated);
            result.location=(result.page?L"原文第 "+std::to_wstring(result.page)+L" 页引用的":L"原文引用页未知 · ")+region.label;return result;
        }
        if(region.key.find(L":box:")!=std::wstring::npos){
            auto frame=get(region.owner.Get(),L"TextFrame"),visible=get(frame.Get(),L"TextRange");const auto same=invoke(region.range.Get(),L"IsEqual",DISPATCH_METHOD,{visible.Get()});
            if(same.v.vt!=VT_BOOL||same.v.boolVal==VARIANT_FALSE){result.location=region.label+L"（串联文本框，原文页码未知）";return result;}
            auto anchor=get(region.owner.Get(),L"Anchor");return originalSourcePosition(document,integer(anchor.Get(),L"Start"),region.label,paginated);
        }
        result=originalSourcePosition(document,integer(region.range.Get(),L"Start"),region.label,paginated);
        if(result.page)result.location+=L"（区域起始页）";
    }catch(const std::exception& error){require(!taskCancelled()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");}
    return result;
}
const ExperimentalSourceRegion* originalRegion(const ExperimentalSourceLocations& locations,const std::wstring& key) {
    const auto found=std::find_if(locations.begin(),locations.end(),[&](const ExperimentalSourceRegion& region){return region.key==key;});
    return found==locations.end()?nullptr:&*found;
}
void attachSourceLocations(const Region& region,std::vector<Target>& targets,const ExperimentalSourceLocations& locations) {
    for(auto& target:targets){target.item.position.page=0;target.item.position.location=L"原文页码未知 · "+region.label;
        if(region.key.rfind(L"Headers:",0)==0||region.key.rfind(L"Footers:",0)==0)target.item.position.location=region.label+L"（共用内容）";}
    const auto original=originalRegion(locations,region.key);if(!original||targets.size()!=original->items.size())return;
    try{const long start=integer(region.range.Get(),L"Start");
        for(size_t i=0;i<targets.size();++i){auto& target=targets[i];const long relative=target.item.position.sourceStart-start;
            if(target.kind==original->items[i].kind&&relative==original->items[i].relativeStart)target.item.position=original->items[i].position;
        }
    }catch(const std::exception& error){require(!taskCancelled()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");}
}
// 预检与正式扫描使用同一套区域清单；区域内没有定界符、OMML 或 MathType 对象时不可能产生目标。
bool regionHasCandidate(const Region& region,bool reverse,const ScanOptions& scanOptions) {
    auto range=region.range.Get();
    if(reverse){auto shapes=get(range,L"InlineShapes");const long count=integer(shapes.Get(),L"Count");for(long i=1;i<=count;++i){auto shape=collectionItem(shapes.Get(),i);const long type=integer(shape.Get(),L"Type");
        if(type!=1&&type!=2)continue;if(text(get(shape.Get(),L"OLEFormat").Get(),L"ClassType").find(L"Equation.DSMT")!=0)continue;return true;}return false;}
    auto maths=get(range,L"OMaths");const long count=integer(maths.Get(),L"Count");
    if(count>0){
        if(scanOptions.convertOmml)return true;
        // 不把禁用的原生公式内部文本当作 LaTeX 候选，只读取真实 Word 范围之间的间隙。
        long cursor=integer(range,L"Start");const long end=integer(range,L"End");
        for(long index=1;index<=count;++index){const auto span=bounds(get(collectionItem(maths.Get(),index).Get(),L"Range").Get());
            if(span.start>cursor&&hasTexMarkers(text(sliceRange(range,cursor,std::min(span.start,end)).Get(),L"Text")))return true;
            cursor=std::max(cursor,span.end);if(cursor>=end)return false;
        }
        return cursor<end&&hasTexMarkers(text(sliceRange(range,cursor,end).Get(),L"Text"));
    }
    return hasTexMarkers(text(range,L"Text"));
}
// 正文原生流一次 Flat OPC 批量读取；对象数不一致时退回逐范围读取，避免把文本框等容器对象计入正文快照。
std::vector<std::vector<unsigned char>> confirmedBodyNative(IDispatch* document) {
    auto assets=bodyNativeAssets(document);if(assets.empty())return {};
    auto shapes=get(get(document,L"Content").Get(),L"InlineShapes");
    return assets.size()==static_cast<size_t>(integer(shapes.Get(),L"Count"))?assets:std::vector<std::vector<unsigned char>>();
}
std::vector<std::vector<unsigned char>> regionNativeFromShapes(IDispatch* root) {
    std::vector<std::vector<unsigned char>> native;auto shapes=get(root,L"InlineShapes");
    const long count=integer(shapes.Get(),L"Count");
    for(long i=1;i<=count;++i){auto shape=collectionItem(shapes.Get(),i);if(integer(shape.Get(),L"Type")!=1)continue;if(text(get(shape.Get(),L"OLEFormat").Get(),L"ClassType").find(L"Equation.DSMT")!=0)continue;native.push_back(nativePayload(rangeNativeAsset(get(shape.Get(),L"Range").Get()).equation));}
    return native;
}
Snapshot snapshotWithState(const Region& region,const DocumentState& state,std::vector<std::vector<unsigned char>> native) {
    Snapshot result;result.key=region.key;result.body=text(region.range.Get(),L"Text");result.layout=state.layout;result.maths=state.maths;
    // 批量读取为空时逐范围兜底；正常任务不会重复导出正文样式包。
    if(native.empty())native=regionNativeFromShapes(region.range.Get());
    result.native=std::move(native);
    auto fields=get(region.range.Get(),L"Fields");const long count=integer(fields.Get(),L"Count");for(long i=1;i<=count;++i){auto field=collectionItem(fields.Get(),i);if(integer(field.Get(),L"Type")==58)continue;result.fields.push_back(text(get(field.Get(),L"Code").Get(),L"Text")+L"\x1f"+text(get(field.Get(),L"Result").Get(),L"Text"));}
    return result;
}
Snapshot snapshot(const Region& region,std::vector<std::vector<unsigned char>> native,ConversionTimings& timings) {return snapshotWithState(region,documentState(region.range.Get(),timings),std::move(native));}
}
ExperimentalSourceLocations captureExperimentalSourceLocations(IDispatch* document,const std::wstring& path,
    const ExperimentalOptions& options,bool reverse,const ScanOptions& scanOptions,bool paginated) {
    ExperimentalSourceLocations result;if(!options.enabled())return result;
    try{
        const auto headers=options.headers?storedHeaders(path):std::vector<HeaderSlot>();auto regions=collectRegions(document,options,headers);
        for(auto& region:regions){taskPulse();require(!taskCancelled(),"TASK_CANCELLED");if(region.main)continue;
            ExperimentalSourceRegion original;original.key=region.key;original.position=regionSourcePosition(document,region,paginated);
            try{
                readRegionTables(region);auto targets=scanRegion(document,region,options,reverse,scanOptions);
                std::sort(targets.begin(),targets.end(),[](const Target& a,const Target& b){return a.item.span.start>b.item.span.start;});
                const long start=integer(region.range.Get(),L"Start");
                for(const auto& target:targets){taskPulse();require(!taskCancelled(),"TASK_CANCELLED");auto position=original.position;
                    if(region.key.rfind(L"body:table:",0)==0)position=originalSourcePosition(document,target.item.position.sourceStart,region.label,paginated);
                    position.sourceStart=target.item.position.sourceStart;original.items.push_back({position.sourceStart-start,target.kind,std::move(position)});
                }
            }catch(const std::exception& error){require(!taskCancelled()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");original.items.clear();}
            result.push_back(std::move(original));
        }
    }catch(const std::exception& error){require(!taskCancelled()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");}
    return result;
}
// 导入文本先准备好；真正写入只在目标文档的单式撤销事务内执行。
namespace {
std::wstring prepareInput(WordSession& word,const Target& target,const std::wstring& sourcePackage,ComPtr<IXSLTemplate>& stylesheet,bool fast,ConversionTimings& timings) {
    if(target.kind==L"OMML"){
        if(!stylesheet)stylesheet=compileOmmlStylesheet(text(word.application.Get(),L"Path")+L"\\OMML2MML.XSL",timings);
        return ommlMathML(sourcePackage,stylesheet.Get(),target.item.display,timings);
    }
    const auto tex=compatibleTex(target.item.inner);
    require(fast||std::all_of(tex.begin(),tex.end(),[](wchar_t c){return c<=127;}),"UNSUPPORTED_NON_ASCII_TEX");
    return target.item.display?L"\\["+tex+L"\\]":L"$"+tex+L"$";
}
}
void emitProcessingIssue(const ProcessingIssue& issue) {
    const auto location=issue.position.location.empty()?L"原文页码未知 · "+issue.region:issue.position.location;
    std::printf("UI_PROCESSING_ISSUE %s %s %s %s %s\n",wireHex(utf8(issue.region)).c_str(),wireHex(utf8(issue.stage)).c_str(),wireHex(utf8(issue.reason)).c_str(),wireHex(utf8(location)).c_str(),diagnosticExcerpt(issue.source).c_str());
}
std::string ExperimentalItem::json() const {
    return std::string("{\"Id\":")+std::to_string(id)+",\"Experimental\":true,\"Region\":"+quote(region)+",\"Kind\":"+quote(kind)+",\"Source\":"+quote(source)+",\"Status\":"+quote(status)+",\"Reason\":"+quote(reason)+sourcePositionJson(position)+"}";
}
std::string ProcessingIssue::json(bool fast) const {
    return "{\"Key\":"+quote(key)+",\"Region\":"+quote(region)+",\"Stage\":"+quote(stage)+",\"Reason\":"+quote(reason)+sourcePositionJson(position)+",\"Source\":"+quote(source)+",\"OriginalPreserved\":"+(fast?"null":"true")+"}";
}
std::string ExperimentalResult::json(bool fast) const {
    std::string result="{\"Enabled\":"+std::string(enabled?"true":"false")+",\"Skipped\":"+quote(skipped)+",\"FormattingApplied\":false,\"Succeeded\":"+std::to_string(succeeded)+",\"FailedPreserved\":"+std::to_string(failed-failedUnverified)+",\"FailedUnverified\":"+std::to_string(failedUnverified)+",\"RejectedPreserved\":"+std::to_string(rejected)+",\"NotAttempted\":"+std::to_string(unattempted)+",\"Items\":[";
    for(size_t i=0;i<items.size();++i)result+=(i?",":"")+items[i].json();
    result+="],\"Issues\":[";for(size_t i=0;i<issues.size();++i)result+=(i?",":"")+issues[i].json(fast);
    return result+"],\"Options\":{\"Tables\":"+(options.tables?"true":"false")+",\"Textboxes\":"+(options.textboxes?"true":"false")+",\"HeadersFooters\":"+(options.headers?"true":"false")+",\"FootnotesEndnotes\":"+(options.notes?"true":"false")+"}}";
}

bool experimentalTargetsPresent(IDispatch* document,const std::wstring& path,const ExperimentalOptions& options,bool reverse,const ScanOptions& scanOptions) {
    // Exists 会漏掉已保存但当前未显示的首页/偶数页页眉；预检必须与正式转换使用同一清单。
    const auto headers=options.headers?storedHeaders(path):std::vector<HeaderSlot>();
    const auto regions=collectRegions(document,options,headers);
    for(const auto& region:regions){if(region.main)continue;try{if(regionHasCandidate(region,reverse,scanOptions))return true;}catch(const std::exception&){return true;}}
    return false;
}

static ExperimentalResult convertPass(WordSession& word,MathTypeApi& api,const std::wstring& path,const ExperimentalOptions& options,bool reverse,bool dollar,size_t firstId,bool selfTest,const std::set<std::wstring>& skipped,std::unique_ptr<ConversionCheckpoint>& checkpoint,const PackageState& bodyPackage,const ScanOptions& scanOptions,ConversionTimings& timings,const ExperimentalSourceLocations& sourceLocations) {
    const bool fast=scanOptions.fast;
    ExperimentalResult result;result.options=options;result.enabled=true;
    const auto headers=options.headers?storedHeaders(path):std::vector<HeaderSlot>();
    auto documents=get(word.application.Get(),L"Documents");
    if(!fast){word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{path.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));word.verifyOwnership();}
    std::vector<Region> regions;try{regions=collectRegions(word.document.Get(),options,headers);}
    catch(const std::exception& error){throw RegionError(L"*",L"实验区",L"ENUMERATING",error);}
    ComPtr<IXSLTemplate> stylesheet;std::string defaults;bool connected=false;std::vector<Binding> bindings;
    std::map<std::wstring,std::vector<ExperimentalItem>> plans;
    std::wstring activeRegion;const ScanItem* activeItem=nullptr;
    auto failure=[&](const std::wstring& key,const std::wstring& stage,const std::exception& error){
        auto owner=std::find_if(regions.begin(),regions.end(),[&](const Region& region){return region.key==key;});RegionError resultError(key,owner==regions.end()?L"实验区":owner->label,stage,error);
        if(const auto original=originalRegion(sourceLocations,key))resultError.issue.position=original->position;
        else resultError.issue.position.location=L"原文页码未知 · "+resultError.issue.region;
        if(key==activeRegion&&stage==L"CONVERTING"&&activeItem){resultError.issue.position=activeItem->position;resultError.issue.source=activeItem->source;}
        if(key==L"*"||key==L"body"||skipped.count(key)){for(const auto& plan:plans)resultError.preserved.insert(resultError.preserved.end(),plan.second.begin(),plan.second.end());}
        else if(plans.count(key))resultError.preserved=plans.at(key);
        return resultError;
    };
    // 重资源只在出现可转换目标时准备：检查点、正文基线、整包状态都不能在无目标时白付成本。
    bool prepared=false,modified=false,converting=false;
    PackageState retained;std::vector<std::vector<unsigned char>> added;std::set<std::vector<unsigned char>> emptyNative;
    std::vector<long> bodyTables;std::map<std::wstring,long> bodyTableEnds;std::vector<std::wstring> bodyGaps,bodyLayout;std::wstring bodyOriginalText;
    auto prepare=[&](){
        if(prepared)return;
        if(fast){prepared=true;return;}
        try{
            auto body=regions.front().range.Get();const long end=integer(body,L"End");long cursor=integer(body,L"Start");std::wstring before;
            bodyLayout=layoutSnapshot(body,timings);bodyOriginalText=text(body,L"Text");
            // 表格单独处理后，仍按不变片段核验整段正文，不能放宽旧正文表格路径的保护范围。
            for(size_t i=0;i<regions.size();++i)if(regions[i].key.rfind(L"body:table:",0)==0){const auto span=bounds(regions[i].range.Get());
                bodyTables.push_back(static_cast<long>(i));bodyTableEnds.emplace(regions[i].key,span.end);bodyGaps.push_back(text(sliceRange(body,cursor,span.start).Get(),L"Text"));before+=bodyGaps.back()+text(regions[i].range.Get(),L"Text");cursor=span.end;}
            bodyGaps.push_back(cursor==end?std::wstring():text(sliceRange(body,cursor,end).Get(),L"Text"));before+=bodyGaps.back();
            require(before==bodyOriginalText,"Independent experimental body gaps use exact Word ranges");
            retained=bodyPackage;
        }catch(const std::exception& error){throw failure(L"body",L"VERIFYING",error);}
        prepared=true;
    };
    std::puts("TIMING_STAGE EXPERIMENTAL_SCAN_CONVERT");
    for(auto& region:regions){
        if(region.main||skipped.count(region.key))continue;
        activeRegion.clear();activeItem=nullptr;
        std::wstring stage=L"SCANNING";std::vector<Target> targets;
        DocumentState committedState;bool stateReady=false;
        try{
        // 同一脚注故事里的前一区域会移动后续坐标，扫描前再读取本区域的表格边界。
        readRegionTables(region);
        std::printf("EXPERIMENTAL_REGION %s\n",utf8(region.key).c_str());require(!regionFault(region.key,L"SCAN"),"Injected experimental region scan failure");
        targets=scanRegion(word.document.Get(),region,options,reverse,scanOptions);std::sort(targets.begin(),targets.end(),[](const Target& a,const Target& b){return a.item.span.start>b.item.span.start;});
        attachSourceLocations(region,targets,sourceLocations);
        for(const auto& target:targets)plans[region.key].push_back({0,region.label,target.kind,target.item.source,target.item.ready?L"NOT_ATTEMPTED":L"REJECTED_PRESERVED",target.item.ready?L"REGION_SKIPPED":target.item.reason,target.item.position});
        long last=integer(region.range.Get(),L"End");for(const auto& target:targets){require(target.item.span.end<=last,"Non-overlapping experimental targets");last=target.item.span.start;}
        for(const auto& target:targets){taskPulse();require(!taskCancelled(),"TASK_CANCELLED");const auto& item=target.item;
            activeRegion=region.key;activeItem=&item;
            stage=L"CONVERTING";const auto label=inTable(region,item.span.start)&&region.key.rfind(L"body:table:",0)!=0?region.label+L"·表格":region.label;
            ExperimentalItem report{firstId++,label,target.kind,item.source,item.ready?L"NOT_ATTEMPTED":L"REJECTED_PRESERVED",item.reason,item.position};
            if(!item.ready){std::printf("UI_EXPERIMENTAL_REJECTED %zu %s %s %s %s\n",report.id,wireHex(utf8(label)).c_str(),wireHex(utf8(item.source)).c_str(),wireHex(utf8(item.reason)).c_str(),wireHex(utf8(item.position.location)).c_str());++result.rejected;result.items.push_back(std::move(report));continue;}
            if(!prepared)prepare();
            if(!converting){std::puts("UI_EXPERIMENTAL_PHASE CONVERTING");converting=true;}
            std::printf("UI_EXPERIMENTAL_ITEM %zu %s %s %s\n",report.id,wireHex(utf8(region.label)).c_str(),diagnosticExcerpt(item.source).c_str(),wireHex(utf8(item.position.location)).c_str());
            auto root=region.range.Get();auto range=sliceRange(root,item.span.start,item.span.end);if(!fast)require(text(range.Get(),L"Text")==item.source,"Experimental source rechecked");
            const auto beforeText=fast?std::wstring():text(root,L"Text");std::wstring sourcePackage,beforeXml;
            if(target.kind==L"OMML"||!fast){
                ScopedConversionTiming timing(timings,ConversionMetric::SourceXml);
                if(target.kind==L"OMML")sourcePackage=text(range.Get(),L"WordOpenXML");
                if(!fast)beforeXml=target.kind==L"OMML"?sourceSignature(sourcePackage,true):sourceSignature(range.Get(),true);
            }
            // 区域内只复用已完整核验并提交的状态；失败恢复不推进基线，区域重跑会重新建立。
            if(!fast&&!stateReady){committedState=documentState(root,timings);stateReady=true;}
            const auto& beforeState=committedState;const auto& beforeLayout=beforeState.layout;const long rootEnd=integer(root,L"End");
            const auto prefix=fast?std::wstring():text(sliceRange(root,integer(root,L"Start"),item.span.start).Get(),L"Text"),suffix=fast?std::wstring():text(sliceRange(root,item.span.end,rootEnd).Get(),L"Text");
            auto undo=get(word.application.Get(),L"UndoRecord");bool record=false,active=false,wrote=false;ComPtr<IDispatch> activeObject;NativeAsset removed;std::vector<unsigned char> native;std::wstring output,input;DocumentState currentState;
            try{
                if(!connected){if(!api.initialized){require(api.init(0,10)>0,"Initialize experimental MathType");api.initialized=true;taskMathTypeConnected();}if(!fast)defaults=readDefaultPrefs(api);connected=true;}
                if(reverse){removed=rangeNativeAsset(range.Get());output=exportedInner(exportNativeTex(api,nativePayload(removed.equation)),item.display,fast);output=(item.display?(dollar?L"$$":L"\\["):L"$")+output+(item.display?(dollar?L"$$":L"\\]"):L"$");}
                else input=prepareInput(word,target,sourcePackage,stylesheet,fast,timings);
                // 首次真正写入前建立检查点；无目标或仅有拒绝项时不复制整份文件。
                if(!fast&&!checkpoint)checkpoint=std::make_unique<ConversionCheckpoint>(path,L".experimental-checkpoint");
                invoke(word.document.Get(),L"Activate",DISPATCH_METHOD);invoke(undo.Get(),L"StartCustomRecord",DISPATCH_METHOD,{L"实验性公式替换"});record=true;
                wrote=true;
                if(reverse)invoke(range.Get(),L"Text",DISPATCH_PROPERTYPUT,{output.c_str()});
                else{
                    invoke(range.Get(),L"Text",DISPATCH_PROPERTYPUT,{L""});invoke(range.Get(),L"Collapse",DISPATCH_METHOD,{1L});
                    auto shape=insertBlankEquation(word.document.Get(),range.Get(),api.install,root,fast);auto ole=get(shape.Get(),L"OLEFormat");invoke(ole.Get(),L"DoVerb",DISPATCH_METHOD,{2L});activeObject=get(ole.Get(),L"Object");active=true;
                    const auto empty=fast?std::vector<unsigned char>():readNative(api,activeObject.Get());if(!fast)emptyNative.insert(empty);
                    {ScopedConversionTiming timing(timings,target.kind==L"OMML"?ConversionMetric::MathmlImport:ConversionMetric::TexImport);require(api.write(activeObject.Get(),target.kind==L"OMML"?2:1,input.c_str(),static_cast<long>(input.size()))==0,"Experimental equation import");}
                    if(!fast){native=readNative(api,activeObject.Get());require(native!=empty&&native.size()>empty.size(),"Experimental equation contains native data");}
                    require(api.close(OLECLOSE_NOSAVE,activeObject.Get())==0,"Close experimental converted equation");active=false;
                }
                taskTestBoundary(L"EXPERIMENTAL_REPLACED");
                require(!regionFault(region.key,target.kind==L"OMML"?L"OMML":L"ITEM"),"Injected experimental formula replacement failure");
                require(!regionFault(region.key,L"RECOVERY"),"Injected experimental region recovery failure");
                // 开发自检在真实替换后注入失败，核对 Undo 能完整恢复公式与容器属性。
                require(!(selfTest&&item.inner==L"experimental_rollback"),"Injected experimental replacement failure");
                const long end=item.span.end+integer(root,L"End")-rootEnd;auto replacement=sliceRange(root,item.span.start,end);
                if(!reverse){auto shapes=get(replacement.Get(),L"InlineShapes");if(!fast)require(integer(shapes.Get(),L"Count")==1,"One experimental object inserted");auto objectRange=get(collectionItem(shapes.Get(),1).Get(),L"Range");output=text(objectRange.Get(),L"Text");if(!fast)require(text(replacement.Get(),L"Text")==output&&nativePayload(rangeNativeAsset(objectRange.Get()).equation)==native,"Experimental native data bound to target");}
                // 含控件的表格在整段和切片中返回不同的结束标记，前后片段须分别核对。
                if(!fast){
                const long currentEnd=integer(root,L"End");currentState=documentState(root,timings);const auto& currentLayout=currentState.layout;
                const bool textPreserved=text(sliceRange(root,integer(root,L"Start"),item.span.start).Get(),L"Text")==prefix&&
                    text(sliceRange(root,end,currentEnd).Get(),L"Text")==suffix&&text(replacement.Get(),L"Text")==output;
                if(!textPreserved||currentLayout!=beforeLayout)std::printf("EXPERIMENTAL_CHANGE text=%d layout=%d start=%ld end=%ld\n",textPreserved,currentLayout==beforeLayout,item.span.start,item.span.end);
                require(textPreserved&&currentLayout==beforeLayout,"Experimental replacement preserves surrounding text and layout");
                if(reverse||!scanOptions.convertOmml)require(currentState.maths==beforeState.maths,"Retained experimental OMML contents and order preserved");
                }
                invoke(undo.Get(),L"EndCustomRecord",DISPATCH_METHOD);record=false;
                if(!fast){if(reverse){require(std::find(bodyPackage.equations.begin(),bodyPackage.equations.end(),removed.equation)!=bodyPackage.equations.end(),"Experimental exported native belongs to source package");const auto payload=nativePayload(removed.equation);retained.equations.erase(std::remove_if(retained.equations.begin(),retained.equations.end(),[&](const auto& value){return nativePayload(value)==payload;}),retained.equations.end());
                    for(const auto& preview:removed.previews)retained.media.erase(std::remove(retained.media.begin(),retained.media.end(),preview),retained.media.end());
                }else{added.push_back(native);}
                const long delta=end-item.span.end,rootStart=integer(root,L"Start");
                for(auto& binding:bindings)if(binding.key==region.key){binding.span.start+=delta;binding.span.end+=delta;}
                bindings.push_back({region.key,region.range,{item.span.start-rootStart,end-rootStart},native,output,result.items.size()});
                }
                report.status=L"SUCCEEDED";++result.succeeded;modified=true;
                if(!fast)committedState=std::move(currentState);
            }catch(const std::exception& error){ScopedConversionTiming recoveryTiming(timings,ConversionMetric::ExperimentalRecovery);std::printf("EXPERIMENTAL_ITEM_FAILURE %s\n",error.what());if(active)require(api.close(OLECLOSE_NOSAVE,activeObject.Get())==0,"Close failed experimental equation");if(record)invoke(undo.Get(),L"EndCustomRecord",DISPATCH_METHOD);
                if(fast?wrote:(text(root,L"Text")!=beforeText||sourceSignature(sliceRange(root,item.span.start,item.span.end).Get(),true)!=beforeXml)){
                    if(target.kind==L"OMML"){
                        auto replacement=sliceRange(root,item.span.start,item.span.end+integer(root,L"End")-rootEnd);invoke(replacement.Get(),L"InsertXML",DISPATCH_METHOD,{sourcePackage.c_str()});
                        const long extra=integer(root,L"End")-rootEnd;
                        if(extra>0){auto trailing=sliceRange(root,item.span.end,item.span.end+extra);require(text(trailing.Get(),L"Text")==std::wstring(static_cast<size_t>(extra),L'\r'),"Experimental OMML restoration adds only trailing paragraph markers");invoke(trailing.Get(),L"Text",DISPATCH_PROPERTYPUT,{L""});}
                    }else{auto restored=invoke(word.document.Get(),L"Undo",DISPATCH_METHOD,{1L});require(restored.v.vt==VT_BOOL&&restored.v.boolVal,"Undo failed experimental replacement");}}
                if(!fast){const auto restoredState=documentState(root,timings);const bool textRestored=text(root,L"Text")==beforeText,layoutRestored=restoredState.layout==beforeLayout,sourceRestored=sourceSignature(sliceRange(root,item.span.start,item.span.end).Get(),true)==beforeXml;
                std::printf("EXPERIMENTAL_RECOVERY text=%d layout=%d source=%d\n",textRestored,layoutRestored,sourceRestored);
                require(textRestored&&layoutRestored&&sourceRestored,"Experimental recovery restores full original formula and layout");
                if(reverse||!scanOptions.convertOmml)require(restoredState.maths==beforeState.maths,"Retained experimental OMML recovery preserves contents and order");
                require(!regionFault(region.key,L"RECOVERY"),"Injected experimental recovery verification failure");
                }
                report.status=fast?L"FAILED_UNVERIFIED":L"FAILED_PRESERVED";result.failedUnverified+=fast;const std::string reason=error.what();report.reason=std::wstring(reason.begin(),reason.end());++result.failed;
            }
            std::printf("UI_ITEM_RESULT %zu %s %s\n",report.id,report.status==L"SUCCEEDED"?"SUCCEEDED":fast?"FAILED_UNVERIFIED":"FAILED_PRESERVED",wireHex(utf8(report.reason)).c_str());result.items.push_back(std::move(report));
        }
        }catch(const std::exception& error){require(!taskCancelled(),"TASK_CANCELLED");if(fast){result.issues.push_back(failure(region.key,stage,error).issue);emitProcessingIssue(result.issues.back());continue;}if(connected)require(readDefaultPrefs(api)==defaults,"Experimental conversion preserves global MathType preferences");throw failure(region.key,stage,error);}
    }
    if(connected&&!fast)require(readDefaultPrefs(api)==defaults,"Experimental conversion preserves global MathType preferences");
    if(fast){if(plans.empty())result.skipped=L"NO_TARGETS";else if(!modified)result.skipped=L"NO_CHANGES";return result;}
    // 没有实验目标：不保存、不重开，直接复用正文阶段已保存并核验的副本。
    if(plans.empty()){
        invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();
        // 重做轮次里全部区域都已被跳过时仍要保留先前的区域问题，不把它伪装成“无目标”。
        if(skipped.empty()){std::puts("EXPERIMENTAL_NO_TARGETS");std::puts("UI_EXPERIMENTAL_SKIPPED NO_TARGETS");result.skipped=L"NO_TARGETS";}
        return result;
    }
    // 有目标但没有任何成功替换：逐式恢复已经核对，这里再确认整段正文没有被全局改动，然后跳过保存重开。
    if(!modified){
        try{require(!prepared||(text(regions.front().range.Get(),L"Text")==bodyOriginalText&&layoutSnapshot(regions.front().range.Get(),timings)==bodyLayout),"Experimental attempts preserve body text and layout");}
        catch(const std::exception& error){throw failure(L"body",L"VERIFYING",error);}
        std::puts("EXPERIMENTAL_NO_CHANGES");std::puts("UI_EXPERIMENTAL_SKIPPED NO_CHANGES");
        invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();result.skipped=L"NO_CHANGES";return result;
    }
    std::puts("UI_EXPERIMENTAL_PHASE VERIFYING");
    std::puts("TIMING_STAGE EXPERIMENTAL_VERIFY");
    try{std::wstring expectedBody;for(size_t i=0;i<bodyTables.size();++i)expectedBody+=bodyGaps[i]+text(regions[bodyTables[i]].range.Get(),L"Text");expectedBody+=bodyGaps.back();
        const auto body=regions.front().range.Get();const auto bodyState=documentState(body,timings);
        require(text(body,L"Text")==expectedBody&&bodyState.layout==bodyLayout,"Experimental conversion preserves all surrounding body text and layout");
        // 先逐式核对预期内容，避免后续对象写入或区域操作串改已经成功的公式。
        for(const auto& binding:bindings){taskPulse();try{const long start=integer(binding.owner.Get(),L"Start");auto range=sliceRange(binding.owner.Get(),start+binding.span.start,start+binding.span.end);
            require(text(range.Get(),L"Text")==binding.output,"Experimental output stays at its bound source position");if(!binding.native.empty())require(nativePayload(rangeNativeAsset(range.Get()).equation)==binding.native,"Experimental native data stays bound to its own source formula");
            if(!reverse){const auto& report=result.items[binding.itemIndex];LayoutWarning context{report.id,report.region,L"",report.source,report.position};
                // 新增的只读版式提醒不能使已通过内容核验的区域回滚。
                try{auto shape=collectionItem(get(range.Get(),L"InlineShapes").Get(),1);auto shapeRange=get(shape.Get(),L"Range");
                    appendFixedLineSpacingWarning(shapeRange.Get(),number(shape.Get(),L"Height"),context,result.warnings);
                }catch(const std::exception& error){require(!taskCancelled()&&std::string(error.what())!="TASK_CANCELLED","TASK_CANCELLED");
                    std::printf("LAYOUT_CHECK_FAILED ID=%zu %s\n",report.id,error.what());context.code=L"LAYOUT_CHECK_UNAVAILABLE";result.warnings.push_back(std::move(context));}
            }
            }
            catch(const std::exception& error){throw failure(binding.key,L"VERIFYING",error);}}
        // 所有修改完成后固定各区域快照，保存重开时逐区域核对对象顺序和完整内容。
        std::vector<Snapshot> expected;expected.reserve(regions.size());auto bodyAssets=confirmedBodyNative(word.document.Get());
        for(const auto& region:regions){try{expected.push_back(region.main?snapshotWithState(region,bodyState,std::move(bodyAssets)):snapshot(region,{},timings));}
            catch(const std::exception& error){throw failure(region.key,L"VERIFYING",error);}}
        require(!taskCancelled(),"TASK_CANCELLED");invoke(word.document.Get(),L"Save",DISPATCH_METHOD);invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();
        word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{path.c_str(),false,true,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));
        std::vector<Region> reopened;try{reopened=collectRegions(word.document.Get(),options,headers);}
        catch(const std::exception& error){throw failure(L"*",L"REOPENING",error);}
        if(reopened.size()!=expected.size()){
            std::printf("REGION_IDENTITY_COUNT expected=%zu reopened=%zu\n",expected.size(),reopened.size());
            for(const auto& region:expected)std::printf("REGION_IDENTITY_EXPECTED %s\n",utf8(region.key).c_str());
            for(const auto& region:reopened)std::printf("REGION_IDENTITY_REOPENED %s\n",utf8(region.key).c_str());
        }
        if(reopened.size()!=expected.size())throw failure(L"*",L"REOPENING",std::runtime_error("Experimental region identities preserved"));
        auto reopenedBodyAssets=confirmedBodyNative(word.document.Get());
        for(const auto& before:expected){taskPulse();try{
            auto found=std::find_if(reopened.begin(),reopened.end(),[&](const Region& region){return region.key==before.key;});require(found!=reopened.end(),"Experimental region identity exists after save");
            const auto actual=found->main?snapshotWithState(*found,documentState(found->range.Get(),timings),std::move(reopenedBodyAssets)):snapshot(*found,{},timings);
            require(skipped.count(before.key)||!regionFault(before.key,L"VERIFY"),"Injected experimental saved verification failure");
            std::printf("EXPERIMENTAL_VERIFY %s key=%d text=%d layout=%d omml=%d fields=%d native=%d\n",utf8(before.key).c_str(),actual.key==before.key,actual.body==before.body,actual.layout==before.layout,actual.maths==before.maths,actual.fields==before.fields,actual.native==before.native);
            require(actual.key==before.key&&actual.body==before.body&&actual.layout==before.layout&&actual.maths==before.maths&&actual.fields==before.fields&&actual.native==before.native,"Saved experimental text, layout, fields, OMML and ordered native contents verified");
            const auto table=bodyTableEnds.find(before.key);
            if(table!=bodyTableEnds.end())result.bodyPositionShifts.push_back({table->second,integer(found->range.Get(),L"End")-table->second});}
            catch(const std::exception& error){throw failure(before.key,L"VERIFYING",error);}}
        // 逐区域有序原生流已覆盖逻辑对象；这里只核对整包资产都必须来自原对象、新导入或空模板。
        try{const auto saved=inspectPackage(path);requireContains(saved.equations,retained.equations,"Experimental retained native streams unchanged");requireContains(saved.media,retained.media,"Experimental retained media unchanged");requireContains(saved.otherEmbeddings,retained.otherEmbeddings,"Experimental unconverted embedded packages unchanged");
        std::set<std::vector<unsigned char>> expectedNative,actualNative,allowed=emptyNative;
        for(const auto& value:retained.equations)expectedNative.insert(nativePayload(value));expectedNative.insert(added.begin(),added.end());
        allowed.insert(expectedNative.begin(),expectedNative.end());if(reverse)for(const auto& value:bodyPackage.equations)allowed.insert(nativePayload(value));
        for(const auto& value:saved.equations){auto native=nativePayload(value);require(allowed.count(native)!=0,"Saved experimental native asset must match known original, imported or empty template data");actualNative.insert(std::move(native));}
        for(const auto& value:expectedNative)require(actualNative.count(value)!=0,"Expected experimental native data exists in saved package");
        }catch(const std::exception& error){throw failure(L"*",L"PACKAGE_VERIFY",error);}
    }catch(const std::exception&){if(word.document){invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();}throw;}
    if(selfTest)invoke(word.document.Get(),L"ExportAsFixedFormat",DISPATCH_METHOD,{(path+L".pdf").c_str(),17L});
    invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();
    std::puts("EXPERIMENTAL_CONVERSION_VERIFIED");return result;
}

ExperimentalResult convertExperimental(WordSession& word,MathTypeApi& api,const std::wstring& path,const ExperimentalOptions& options,bool reverse,bool dollar,size_t firstId,const PackageState& bodyPackage,bool selfTest,const ScanOptions& scanOptions,ConversionTimings& timings,const ExperimentalSourceLocations& sourceLocations) {
    const bool fast=scanOptions.fast;
    if(!options.enabled())return {};
    std::unique_ptr<ConversionCheckpoint> checkpoint;std::set<std::wstring> skipped;std::vector<ProcessingIssue> issues;std::vector<ExperimentalItem> preserved;
    if(fast){
        try{return convertPass(word,api,path,options,reverse,dollar,firstId,selfTest,skipped,checkpoint,bodyPackage,scanOptions,timings,sourceLocations);}
        catch(const std::exception& error){require(!taskCancelled(),"TASK_CANCELLED");const std::string reason=error.what();ExperimentalResult result;result.enabled=true;result.options=options;result.issues.push_back({L"*",L"实验区域",L"FAST_CONVERSION",std::wstring(reason.begin(),reason.end()),{-1,0,L"原文页码未知 · 实验区域"}});emitProcessingIssue(result.issues.back());return result;}
    }
    for(;;){
        try{
            auto result=convertPass(word,api,path,options,reverse,dollar,firstId,selfTest,skipped,checkpoint,bodyPackage,scanOptions,timings,sourceLocations);result.issues=issues;
            for(auto item:preserved){item.id=firstId+result.items.size();result.unattempted+=item.status==L"NOT_ATTEMPTED";result.rejected+=item.status==L"REJECTED_PRESERVED";result.items.push_back(std::move(item));}
            return result;
        }catch(const RegionError& error){
            require(!taskCancelled(),"TASK_CANCELLED");
            emitProcessingIssue(error.issue);
            // 局部恢复失败后不再信任当前文档；先丢弃全部本轮实验修改，再从正文检查点重做其余区域。
            if(word.document){invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();}
            // 扫描阶段失败时还没有建立检查点，候选文件仍是已核验的正文副本，无需恢复。
            if(checkpoint)checkpoint->restore(path);
            issues.push_back(error.issue);preserved.insert(preserved.end(),error.preserved.begin(),error.preserved.end());
            std::printf("EXPERIMENTAL_REGION_SKIPPED %s %s\n",utf8(error.issue.key).c_str(),error.what());std::puts("UI_EXPERIMENTAL_RESET");
            for(const auto& issue:issues)emitProcessingIssue(issue);
            if(error.issue.key==L"body"||error.issue.key==L"*"||!skipped.insert(error.issue.key).second){
                // 无法核验区域边界时只交付此前已经完整核验的正文副本，明确其余实验项未完成。
                ExperimentalResult result;result.enabled=true;result.options=options;result.issues=issues;
                result.issues.push_back({L"*",L"其余实验区域",L"NOT_ATTEMPTED",L"EXPERIMENTAL_CHECKPOINT_RESTORED",{-1,0,L"原文页码未知 · 其余实验区域"}});emitProcessingIssue(result.issues.back());
                for(auto item:preserved){item.id=firstId+result.items.size();result.unattempted+=item.status==L"NOT_ATTEMPTED";result.rejected+=item.status==L"REJECTED_PRESERVED";result.items.push_back(std::move(item));}
                return result;
            }
        }
    }
}
