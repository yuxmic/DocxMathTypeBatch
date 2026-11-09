// 正文 LaTeX 与 OMML 的原生包路径；一次扫描、独立资产、逆序替换。
#include "ultra_conversion.h"
#include "equation_styles.h"
#include "equation_import.h"
#include "omml_adapter.h"
#include "omml_tex.h"

namespace {
constexpr const wchar_t* namespaces=L"xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'";
ComPtr<IXMLDOMNode> select(IXMLDOMNode* node,const wchar_t* query){ComPtr<IXMLDOMNode> found;check(node->selectSingleNode(Value(query).v.bstrVal,&found),"Select ultra XML node");return found;}
std::vector<ComPtr<IXMLDOMNode>> nodes(IXMLDOMNode* node,const wchar_t* query){ComPtr<IXMLDOMNodeList> list;check(node->selectNodes(Value(query).v.bstrVal,&list),"Select ultra XML nodes");long count{};check(list->get_length(&count),"Ultra XML count");std::vector<ComPtr<IXMLDOMNode>> found;for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> value;check(list->get_item(i,&value),"Ultra XML item");found.push_back(std::move(value));}return found;}
std::wstring nodeText(IXMLDOMNode* node){Value value;value.v.vt=VT_BSTR;check(node->get_text(&value.v.bstrVal),"Ultra XML text");return value.v.bstrVal?value.v.bstrVal:L"";}
std::wstring nodeXml(IXMLDOMNode* node){Value value;value.v.vt=VT_BSTR;check(node->get_xml(&value.v.bstrVal),"Ultra XML serialization");return value.v.bstrVal;}
std::wstring attribute(IXMLDOMNode* node,const wchar_t* name,const wchar_t* fallback=L""){auto value=select(node,(L"@"+std::wstring(name)).c_str());return value?nodeText(value.Get()):fallback;}
ComPtr<IXMLDOMNode> importXml(IXMLDOMDocument2* document,const std::wstring& fragment){auto parsed=xmlDocument(L"<root "+std::wstring(namespaces)+L">"+fragment+L"</root>");auto child=select(parsed.Get(),L"/root/*[1]");ComPtr<IXMLDOMDocument3> owner;check(document->QueryInterface(IID_PPV_ARGS(&owner)),"Ultra DOM owner");ComPtr<IXMLDOMNode> copied;check(owner->importNode(child.Get(),VARIANT_TRUE,&copied),"Import ultra XML");return copied;}
void append(IXMLDOMNode* parent,IXMLDOMNode* child){check(parent->appendChild(child,nullptr),"Append ultra XML");}
void appendXml(IXMLDOMDocument2* document,IXMLDOMNode* parent,const std::wstring& fragment){
    // 同一片段可包含多个相邻节点，编号整体只解析及导入一次。
    auto group=importXml(document,L"<root>"+fragment+L"</root>");for(auto& child:nodes(group.Get(),L"*"))append(parent,child.Get());
}
ComPtr<IXMLDOMDocument2> partXml(IOpcPart* part){ComPtr<IStream> stream;check(part->GetContentStream(&stream),"Read ultra XML part");auto document=xmlDocument(L"<root/>");Value source;source.v.vt=VT_UNKNOWN;source.v.punkVal=stream.Get();stream->AddRef();VARIANT_BOOL loaded{};check(document->load(source.v,&loaded),"Load ultra XML part");require(loaded!=VARIANT_FALSE,"Valid ultra XML part");check(document->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(namespaces).v),"Ultra XML namespaces");return document;}
void writeXml(IOpcPart* part,IXMLDOMDocument2* document){ComPtr<IStream> stream;check(part->GetContentStream(&stream),"Ultra XML output stream");ULARGE_INTEGER zero{};check(stream->SetSize(zero),"Replace ultra XML bytes");LARGE_INTEGER start{};check(stream->Seek(start,STREAM_SEEK_SET,nullptr),"Rewind ultra XML output");Value target;target.v.vt=VT_UNKNOWN;target.v.punkVal=stream.Get();stream->AddRef();check(document->save(target.v),"Write ultra XML part");}
ComPtr<IOpcPartUri> partUri(IOpcFactory* factory,const std::wstring& path){ComPtr<IOpcPartUri> uri;check(factory->CreatePartUri(path.c_str(),&uri),"Ultra part URI");return uri;}
ComPtr<IOpcPart> findPart(IOpcFactory* factory,IOpcPartSet* parts,const wchar_t* path){auto uri=partUri(factory,path);BOOL exists{};check(parts->PartExists(uri.Get(),&exists),"Find ultra part");if(!exists)return {};ComPtr<IOpcPart> part;check(parts->GetPart(uri.Get(),&part),"Open ultra part");return part;}
ComPtr<IOpcPart> createPart(IOpcFactory* factory,IOpcPartSet* parts,const std::wstring& path,const wchar_t* type){auto uri=partUri(factory,path);ComPtr<IOpcPart> part;check(parts->CreatePart(uri.Get(),type,OPC_COMPRESSION_NORMAL,&part),"Create ultra part");return part;}
void relationship(IOpcFactory* factory,IOpcRelationshipSet* relations,const std::wstring& id,const wchar_t* kind,const std::wstring& path){auto uri=partUri(factory,path);ComPtr<IOpcRelationship> relation;check(relations->CreateRelationship(id.c_str(),(L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"+std::wstring(kind)).c_str(),uri.Get(),OPC_URI_TARGET_MODE_INTERNAL,&relation),"Create ultra relationship");}
void assetPart(IOpcFactory* factory,IOpcPartSet* parts,const std::wstring& path,const wchar_t* type,const std::vector<unsigned char>& bytes){auto part=createPart(factory,parts,path,type);ComPtr<IStream> stream;check(part->GetContentStream(&stream),"Ultra asset stream");ULONG written{};check(stream->Write(bytes.data(),static_cast<ULONG>(bytes.size()),&written),"Write ultra asset");require(written==bytes.size(),"Complete ultra asset");}

#include "ultra_source.h"
long numberWidth(IXMLDOMDocument2* document){require(!select(document,L"//w:instrText[contains(.,'MTEqn')] | //w:fldSimple[contains(@w:instr,'MTEqn')]"),"ULTRA_EXISTING_NUMBERS");auto sections=nodes(document,L"//w:sectPr");require(sections.size()==1,"ULTRA_SINGLE_SECTION_REQUIRED");auto columns=select(sections[0].Get(),L"w:cols");require(!columns||attribute(columns.Get(),L"w:num",L"1")==L"1","ULTRA_SINGLE_COLUMN_REQUIRED");auto page=select(sections[0].Get(),L"w:pgSz"),margin=select(sections[0].Get(),L"w:pgMar");require(page&&margin,"ULTRA_PAGE_SIZE_REQUIRED");return std::stol(attribute(page.Get(),L"w:w"))-std::stol(attribute(margin.Get(),L"w:left",L"0"))-std::stol(attribute(margin.Get(),L"w:right",L"0"))-std::stol(attribute(margin.Get(),L"w:gutter",L"0"));}
void layout(IXMLDOMDocument2* document,IXMLDOMNode* p,const std::wstring& alignment,long width,IXMLDOMNode* prepared=nullptr){auto props=select(p,L"w:pPr");if(!props){props=importXml(document,L"<w:pPr/>");ComPtr<IXMLDOMNode> first;check(p->get_firstChild(&first),"First paragraph node");check(p->insertBefore(props.Get(),Value(first.Get()).v,nullptr),"Insert paragraph properties");}
    for(auto& child:nodes(props.Get(),width?L"w:jc | w:tabs | w:ind":L"w:jc"))check(props->removeChild(child.Get(),nullptr),"Replace ultra layout property");
    if(prepared){for(auto& child:nodes(prepared,L"*"))append(props.Get(),ultraClone(child.Get()).Get());return;}
    if(width)appendXml(document,props.Get(),L"<w:tabs><w:tab w:val='center' w:pos='"+std::to_wstring(width/2)+L"'/><w:tab w:val='right' w:pos='"+std::to_wstring(width)+L"'/></w:tabs>");
    appendXml(document,props.Get(),L"<w:jc w:val='"+(width?std::wstring(L"left"):alignment)+L"'/>");
}
void number(IXMLDOMDocument2* document,IXMLDOMNode* p,size_t value){
    appendXml(document,p,L"<w:r><w:fldChar w:fldCharType='begin'/></w:r>"
        L"<w:r><w:instrText xml:space='preserve'> MACROBUTTON MTPlaceRef \\* MERGEFORMAT </w:instrText></w:r>"
        L"<w:r><w:fldChar w:fldCharType='begin'/></w:r>"
        L"<w:r><w:instrText xml:space='preserve'> SEQ MTEqn \\h \\* MERGEFORMAT </w:instrText></w:r>"
        L"<w:r><w:fldChar w:fldCharType='end'/></w:r><w:r><w:instrText>(</w:instrText></w:r>"
        L"<w:fldSimple w:instr=' SEQ MTEqn \\c \\* Arabic \\* MERGEFORMAT '><w:r><w:instrText>"+std::to_wstring(value)+L"</w:instrText></w:r></w:fldSimple>"
        L"<w:r><w:instrText>)</w:instrText></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r>");
}
void numberProperties(IOpcFactory* factory,IOpcPackage* package,IOpcPartSet* parts,const std::wstring& key){auto part=findPart(factory,parts,L"/docProps/custom.xml");const bool existed=part!=nullptr;auto properties=part?partXml(part.Get()):xmlDocument(L"<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties' xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'/>");auto root=select(properties.Get(),L"/*");long pid=1;for(auto& child:nodes(root.Get(),L"*"))pid=std::max(pid,std::stol(attribute(child.Get(),L"pid")));
    for(const auto& entry:{std::make_pair(L"MTEquationNumber2",L"(#E1)"),std::make_pair(L"MTCustomEquationNumber",L"0"),std::make_pair(L"MTEquationSection",L"1")}){auto old=select(root.Get(),(L"*[@name='"+std::wstring(entry.first)+L"']").c_str());if(old)check(root->removeChild(old.Get(),nullptr),"Replace MathType property");appendXml(properties.Get(),root.Get(),L"<property xmlns='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties' xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes' fmtid='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}' pid='"+std::to_wstring(++pid)+L"' name='"+entry.first+L"'><vt:lpwstr>"+entry.second+L"</vt:lpwstr></property>");}
    if(!existed){part=createPart(factory,parts,L"/docProps/custom.xml",L"application/vnd.openxmlformats-officedocument.custom-properties+xml");ComPtr<IOpcRelationshipSet> relations;check(package->GetRelationshipSet(&relations),"Package root relationships");relationship(factory,relations.Get(),key+L"_custom",L"custom-properties",L"/docProps/custom.xml");}writeXml(part.Get(),properties.Get());
}
std::wstring stylesheetPath(){std::wstring path(32768,L'\0');DWORD length=static_cast<DWORD>(path.size()*sizeof(wchar_t));require(RegGetValueW(HKEY_LOCAL_MACHINE,L"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration",L"InstallationPath",RRF_RT_REG_SZ,nullptr,path.data(),&length)==ERROR_SUCCESS,"OFFICE_OMML_TRANSFORM_MISSING");path.resize(wcslen(path.c_str()));return path+L"\\root\\Office16\\OMML2MML.XSL";}
void checkCancelled(const UltraOptions& options){taskPulse();require(!taskCancelled()&&(options.stopFile.empty()||GetFileAttributesW(options.stopFile.c_str())==INVALID_FILE_ATTRIBUTES),"TASK_CANCELLED");}
}

UltraResult convertUltra(MathTypeApi& api,const std::wstring& candidate,const UltraOptions& options,ConversionTimings& timings){
    check(OleInitialize(nullptr),"Initialize ultra OLE apartment");struct Apartment{~Apartment(){OleUninitialize();}} apartment;UltraResult result;
    ComPtr<IOpcFactory> factory;ComPtr<IStream> input;ComPtr<IOpcPackage> package;ComPtr<IOpcPartSet> parts;ComPtr<IOpcPart> main;ComPtr<IXMLDOMDocument2> document;std::vector<Candidate> selected;long width{};
    std::puts("UI_PHASE PACKAGE_READ");{
        ScopedConversionTiming timing(timings,ConversionMetric::PackageRead);check(CoCreateInstance(__uuidof(OpcFactory),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&factory)),"Ultra package factory");check(factory->CreateStreamOnFile(candidate.c_str(),OPC_STREAM_IO_READ,nullptr,0,&input),"Read private ultra package");check(factory->ReadPackageFromStream(input.Get(),OPC_CACHE_ON_ACCESS,&package),"Load private ultra package");check(package->GetPartSet(&parts),"Ultra package parts");
        main=findPart(factory.Get(),parts.Get(),L"/word/document.xml");require(main!=nullptr,"Ultra main document exists");document=partXml(main.Get());
        auto settingsPart=findPart(factory.Get(),parts.Get(),L"/word/settings.xml");if(settingsPart){auto settings=partXml(settingsPart.Get());require(!select(settings.Get(),L"//w:documentProtection"),"Protected document cannot be converted");require(!select(settings.Get(),L"//w:trackRevisions[not(@w:val='0' or @w:val='false' or @w:val='off')]"),"Disable revision tracking in source before batch conversion");}
        selected=candidates(document.Get(),result,options.progress);result.targets=selected.size();
        for(const auto& item:selected)if(item.tex)++result.texTargets;else ++result.ommlTargets;
        if(options.numbers&&std::any_of(selected.begin(),selected.end(),[](const Candidate& item){return item.display;}))width=numberWidth(document.Get());
    }
    if(options.progress)std::printf("UI_SCAN_COUNTS 0 %zu %zu 0\n",result.targets,static_cast<size_t>(std::count_if(selected.begin(),selected.end(),[](const Candidate& item){return !item.display;})));
    checkCancelled(options);if(selected.empty())return result;
    std::puts("UI_PHASE INITIALIZING");ComPtr<IXSLTemplate> stylesheet;if(result.ommlTargets)stylesheet=compileOmmlStylesheet(stylesheetPath(),timings);
    if(!api.initialized){require(api.init(0,10)>0,"Initialize ultra MathType API");api.initialized=true;taskMathTypeConnected();}
    const auto empty=[&](){EquationHost host(timings);return readNative(api,host.object.Get());}();const auto preferencesPath=candidate+L".ultra.eqp";bool preferencesOwned{};std::string preferences;
    try{preferences=sizedPreferences(api,options.points,preferencesPath,preferencesOwned);}catch(...){if(preferencesOwned)DeleteFileW(preferencesPath.c_str());throw;}
    if(preferencesOwned)require(DeleteFileW(preferencesPath.c_str())!=0,"Remove ultra preference file");
    ComPtr<IStorage> blanks[2];for(const auto& item:selected){const size_t index=item.display?1:0;if(!blanks[index])blanks[index]=styledStorage(api,styledNative(api,empty,preferences,item.display),timings);}
    GUID guid{};check(CoCreateGuid(&guid),"Ultra asset identity");wchar_t guidText[40]{};StringFromGUID2(guid,guidText,40);std::wstring key=L"ultra";for(wchar_t c:std::wstring(guidText))if(iswalnum(c))key+=c;
    ComPtr<IOpcRelationshipSet> relations;check(main->GetRelationshipSet(&relations),"Ultra body relationships");
    // 所有公式共用一份解析后的 XML 骨架；每个位置克隆后写入独立标识和尺寸。
    const auto skeleton=importXml(document.Get(),L"<w:r><w:rPr><w:position w:val='0'/></w:rPr><w:object w:dxaOrig='0' w:dyaOrig='0'><v:shape id='' style='' coordsize='21600,21600' o:ole='' filled='f' stroked='f'><v:imagedata r:id='' o:title=''/></v:shape><o:OLEObject Type='Embed' ProgID='Equation.DSMT4' ShapeID='' DrawAspect='Content' ObjectID='' r:id=''/></w:object></w:r>");
    auto set=[](IXMLDOMNode* node,const wchar_t* name,const std::wstring& value){ComPtr<IXMLDOMElement> element;check(node->QueryInterface(IID_PPV_ARGS(&element)),"Ultra asset element");check(element->setAttribute(Value(name).v.bstrVal,Value(value.c_str()).v),"Set ultra asset attribute");};
    std::map<std::wstring,Asset> cache;std::vector<ComPtr<IXMLDOMNode>> replacements(selected.size()),displayParagraphs(selected.size());
    // 复用已有计时差值，仅保留最慢八式的位置和输入长度，不逐式输出源码。
    struct ImportCost {size_t index{},block{},characters{};double milliseconds{};bool tex{};};std::array<ImportCost,8> slowest{};
    std::puts("UI_PHASE CONVERTING");StageProgress progress("PROCESS_FORMULAS",selected.size(),options.progress);
    for(size_t index=0;index<selected.size();++index){checkCancelled(options);const auto& item=selected[index];const Asset* asset{};
        try{
            const auto source=item.tex?normalizedTex(item.source):nodeXml(item.math.Get());
            const auto cacheKey=std::wstring(item.tex?L"T":L"M")+(item.display?L"D":L"I")+source;auto found=cache.find(cacheKey);
            if(found!=cache.end()){asset=&found->second;++result.cacheHits;}
            else{
                bool useTex=item.tex;std::wstring equation=source;
                if(!item.tex){
                    {ScopedConversionTiming timing(timings,ConversionMetric::OmmlTexPrepare);useTex=tryOmmlTex(item.math.Get(),equation);}
                    if(useTex)++result.ommlTexImports;
                    else{++result.ommlMathmlImports;equation=ommlMathML(source,stylesheet.Get(),item.display,timings,true);}
                }
                CompatibleTexInput tex;
                if(useTex){
                    tex=compatibleTexInput(std::move(equation));
                    equation=(item.display?L"\\[":L"$")+tex.tex+(item.display?L"\\]":L"$");
                    if(!tex.fixes.empty()){
                        equation=mathmlFromTex(api,equation,item.display,tex.fixes,timings);
                        useTex=false;
                    }
                }
                EquationHost host(timings,blanks[item.display?1:0].Get());
                const auto metric=useTex?ConversionMetric::TexImport:ConversionMetric::MathmlImport;auto& elapsed=timings.samples[static_cast<size_t>(metric)].elapsed;const auto before=elapsed;
                {ScopedConversionTiming timing(timings,metric);require(equation.size()<=LONG_MAX&&api.write(host.object.Get(),useTex?1:2,equation.c_str(),static_cast<long>(equation.size()))==0,"Import ultra equation");}
                const auto milliseconds=std::chrono::duration<double,std::milli>(elapsed-before).count();
                if(milliseconds>slowest.back().milliseconds){slowest.back()={index+1,item.bodyBlock,equation.size(),milliseconds,useTex};std::sort(slowest.begin(),slowest.end(),[](const ImportCost& a,const ImportCost& b){return a.milliseconds>b.milliseconds;});}
                asset=&cache.emplace(cacheKey,host.save(timings)).first->second;
            }
        }catch(const UltraServiceError&){throw;}catch(const std::exception& error){checkCancelled(options);const std::string reason=error.what();UltraProblem problem{item.bodyBlock,item.source,std::wstring(reason.begin(),reason.end())};result.problems.push_back(problem);++result.failed;
            std::printf("UI_PROCESSING_ISSUE %s %s %s %s %s\n",wireHex(utf8(L"正文")).c_str(),wireHex("CONVERTING").c_str(),wireHex(reason).c_str(),wireHex(utf8(L"正文块 "+std::to_wstring(item.bodyBlock))).c_str(),wireHex(utf8(problem.source)).c_str());}
        if(asset){ScopedConversionTiming timing(timings,ConversionMetric::PackageWrite);const auto id=key+L"_"+std::to_wstring(index+1),image=id+L"_image",ole=id+L"_ole",shape=L"_x0000_"+id;wchar_t dimensions[160]{};swprintf_s(dimensions,L"width:%.4fpt;height:%.4fpt",asset->cx*72.0/2540,asset->cy*72.0/2540);
            auto run=ultraClone(skeleton.Get());auto object=select(run.Get(),L"w:object"),picture=select(object.Get(),L"v:shape"),embedding=select(object.Get(),L"o:OLEObject");
            set(select(run.Get(),L"w:rPr/w:position").Get(),L"w:val",std::to_wstring(asset->baselineHalfPoints));set(object.Get(),L"w:dxaOrig",std::to_wstring(MulDiv(asset->cx,1440,2540)));set(object.Get(),L"w:dyaOrig",std::to_wstring(MulDiv(asset->cy,1440,2540)));
            set(picture.Get(),L"id",shape);set(picture.Get(),L"style",dimensions);set(select(picture.Get(),L"v:imagedata").Get(),L"r:id",image);set(embedding.Get(),L"ShapeID",shape);set(embedding.Get(),L"ObjectID",L"_"+std::to_wstring(GetTickCount64()+index));set(embedding.Get(),L"r:id",ole);
            // 缓存只省生成过程；每个出现位置仍有独立 OLE 包部件，编辑互不影响。
            assetPart(factory.Get(),parts.Get(),L"/word/media/"+id+L".wmf",L"image/x-wmf",asset->preview);assetPart(factory.Get(),parts.Get(),L"/word/embeddings/"+id+L".bin",L"application/vnd.openxmlformats-officedocument.oleObject",asset->ole);relationship(factory.Get(),relations.Get(),image,L"image",L"/word/media/"+id+L".wmf");relationship(factory.Get(),relations.Get(),ole,L"oleObject",L"/word/embeddings/"+id+L".bin");
            replacements[index]=std::move(run);++result.succeeded;if(!item.tex)--result.remaining;
        }progress.advance();
    }
    for(const auto& cost:slowest)if(cost.index)std::printf("ULTRA_IMPORT_DETAIL INDEX=%zu BODY_BLOCK=%zu KIND=%s INPUT_CHARS=%zu MS=%.3f\n",cost.index,cost.block,cost.tex?"TEX":"MATHML",cost.characters,cost.milliseconds);
    checkCancelled(options);std::puts("UI_PHASE PACKAGE_WRITE");{
        ScopedConversionTiming timing(timings,ConversionMetric::PackageWrite);
        for(size_t i=selected.size();i-->0;)if(replacements[i]){checkCancelled(options);auto paragraph=replaceCandidate(document.Get(),selected[i],replacements[i].Get());if(selected[i].display)displayParagraphs[i]=std::move(paragraph);}
        auto numberTemplate=importXml(document.Get(),L"<root/>");if(width)number(document.Get(),numberTemplate.Get(),1);
        const auto tabTemplate=importXml(document.Get(),L"<w:r><w:tab/></w:r>");
        std::map<std::wstring,ComPtr<IXMLDOMNode>> layouts;
        for(size_t i=0;i<selected.size();++i)if(displayParagraphs[i]){
            const auto alignment=width?std::wstring(L"left"):selected[i].alignment;auto& prepared=layouts[alignment];
            if(!prepared){auto sample=importXml(document.Get(),L"<w:p/>");layout(document.Get(),sample.Get(),alignment,width);prepared=select(sample.Get(),L"w:pPr");}
            auto& paragraph=displayParagraphs[i];layout(document.Get(),paragraph.Get(),alignment,width,prepared.Get());
            if(width){auto tab=ultraClone(tabTemplate.Get());check(paragraph->insertBefore(tab.Get(),Value(replacements[i].Get()).v,nullptr),"Center ultra equation");append(paragraph.Get(),ultraClone(tabTemplate.Get()).Get());
                auto group=ultraClone(numberTemplate.Get());auto value=select(group.Get(),L"w:fldSimple/w:r/w:instrText");const auto text=std::to_wstring(++result.numbersAdded);check(value->put_text(Value(text.c_str()).v.bstrVal),"Set cached ultra number");for(auto& child:nodes(group.Get(),L"*"))append(paragraph.Get(),child.Get());}
        }
        if(result.numbersAdded)numberProperties(factory.Get(),package.Get(),parts.Get(),key);
        writeXml(main.Get(),document.Get());ComPtr<IStream> output;check(CreateStreamOnHGlobal(nullptr,TRUE,&output),"Ultra package output buffer");check(factory->WritePackageToStream(package.Get(),OPC_WRITE_DEFAULT,output.Get()),"Serialize ultra package");
        // 释放所有引用后才覆盖同一个私有副本，原件从不参与写入。
        relations.Reset();main.Reset();parts.Reset();package.Reset();input.Reset();LARGE_INTEGER zero{};check(output->Seek(zero,STREAM_SEEK_SET,nullptr),"Rewind ultra package");ComPtr<IStream> destination;check(factory->CreateStreamOnFile(candidate.c_str(),OPC_STREAM_IO_WRITE,nullptr,0,&destination),"Write private ultra package");ULARGE_INTEGER all{};all.QuadPart=~0ULL;check(output->CopyTo(destination.Get(),all,nullptr,nullptr),"Complete ultra package write");
    }if(taskShared)InterlockedExchange(&taskShared->mathReady,0);std::puts("UI_PHASE TERMINATING_MATHTYPE");require(api.term()==0,"Terminate ultra MathType API");api.initialized=false;return result;
}
void ultraXmlSelfTest(){
    auto document=xmlDocument(L"<w:document "+std::wstring(namespaces)+L"><w:body><w:p><w:pPr><w:jc w:val='left'/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val='center'/></m:oMathParaPr><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p><w:sectPr><w:pgSz w:w='10000'/><w:pgMar w:left='1000' w:right='1000'/></w:sectPr></w:body></w:document>");
    check(document->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(namespaces).v),"Test ultra namespaces");UltraResult scan;auto chosen=candidates(document.Get(),scan);
    require(chosen.size()==1&&chosen[0].alignment==L"center"&&scan.remaining==1,"Capture original OMML alignment and count");
    layout(document.Get(),chosen[0].paragraph.Get(),chosen[0].alignment,0);require(nodeText(select(chosen[0].paragraph.Get(),L"w:pPr/w:jc/@w:val").Get())==L"center","Unnumbered equation retains center alignment");
    require(numberWidth(document.Get())==8000,"Number width prepared before generating assets");appendXml(document.Get(),select(document.Get(),L"/w:document/w:body").Get(),L"<w:sectPr/>");bool rejected{};try{numberWidth(document.Get());}catch(const std::exception&){rejected=true;}require(rejected,"Reject multi-section numbering before conversion");
    // 编号合并导入后仍产生原来的九个相邻节点，计数与位置直接复用候选扫描。
    const auto before=nodes(chosen[0].paragraph.Get(),L"*").size();number(document.Get(),chosen[0].paragraph.Get(),7);
    require(nodes(chosen[0].paragraph.Get(),L"*").size()==before+9&&nodeText(select(chosen[0].paragraph.Get(),L"w:fldSimple/w:r/w:instrText").Get())==L"7","Append complete equation number once");
    ultraSourceSelfTest();
    std::puts("ULTRA_XML_SELF_TEST_PASSED");
}
