// 共用公式对象、包资产和安全快照；正文与实验分支使用相同的内容校验。
#pragma once
#include "docx_scan.h"
#include "mathtype_api.h"
#include <xmllite.h>
#include <wincrypt.h>
#include <set>
// 检查点仅复制本任务的私有副本；恢复前由调用方关闭 Word 文档并释放范围。
struct ConversionCheckpoint {
    std::wstring path;
    ConversionCheckpoint(const std::wstring& source,const wchar_t* suffix):path(source+suffix){taskCopy(source,path);}
    ~ConversionCheckpoint(){if(!DeleteFileW(path.c_str())&&GetLastError()!=ERROR_FILE_NOT_FOUND)std::puts("OWNED_CHECKPOINT_CLEANUP_FAILED");}
    void restore(const std::wstring& destination) const {
        require(!taskCancelled(),"TASK_CANCELLED");
        require(CopyFileW(path.c_str(),destination.c_str(),FALSE)!=0,"Restore owned conversion checkpoint");
    }
};
inline bool conversionFault(const wchar_t* variable,const std::wstring& key,const wchar_t* stage) {
    wchar_t value[512]{};return GetEnvironmentVariableW(variable,value,512)&&std::wstring(value)==key+L"|"+stage;
}
// 正文和实验区共用日志摘要，截断时保留完整 UTF-16 字符，不能让诊断导致转换失败。
inline std::string diagnosticExcerpt(const std::wstring& source) {
    auto excerpt=source.substr(0,160);
    if(!excerpt.empty()&&excerpt.back()>=0xd800&&excerpt.back()<=0xdbff)excerpt.pop_back();
    if(excerpt.size()<source.size())excerpt+=L"…";
    return wireHex(utf8(excerpt));
}
// 只读包检查使用系统 Packaging API/结构化存储，不激活任何既有 MathType。
// 依据：https://learn.microsoft.com/en-us/windows/win32/api/msopc/nf-msopc-iopcfactory-readpackagefromstream
inline std::vector<unsigned char> streamBytes(IStream* stream) {
    STATSTG stat{};
    check(stream->Stat(&stat, STATFLAG_NONAME), "Package stream size");
    require(stat.cbSize.QuadPart <= 128 * 1024 * 1024, "Package part exceeds 128 MiB safety limit");
    LARGE_INTEGER zero{};
    check(stream->Seek(zero, STREAM_SEEK_SET, nullptr), "Package stream seek");
    std::vector<unsigned char> bytes(static_cast<size_t>(stat.cbSize.QuadPart));
    ULONG read{};
    if (!bytes.empty()) {
        check(stream->Read(bytes.data(), static_cast<ULONG>(bytes.size()), &read), "Package stream read");
        require(read == bytes.size(), "Complete package part read");
    }
    return bytes;
}
struct PackageState {
    std::vector<std::vector<unsigned char>> equations, media;
    bool displayStyle{};
};
// WordOpenXML 的单对象包和 DOCX 的 OLE part 共用系统结构化存储读取。
inline std::vector<unsigned char> nativeFromOle(const std::vector<unsigned char>& bytes) {
    ComPtr<ILockBytes> locked;
    check(CreateILockBytesOnHGlobal(nullptr, TRUE, &locked), "OLE byte storage");
    ULARGE_INTEGER offset{}; ULONG written{};
    require(bytes.size() <= MAXDWORD, "OLE storage size bounds");
    check(locked->WriteAt(offset, bytes.data(), static_cast<ULONG>(bytes.size()), &written), "Copy OLE bytes");
    require(written == bytes.size(), "Complete OLE byte copy");
    ComPtr<IStorage> storage;
    check(StgOpenStorageOnILockBytes(locked.Get(), nullptr, STGM_READ | STGM_SHARE_EXCLUSIVE, nullptr, 0, &storage), "Read OLE storage");
    ComPtr<IStream> native;
    const auto status = storage->OpenStream(L"Equation Native", nullptr, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &native);
    if (status == STG_E_FILENOTFOUND) return {};
    check(status, "Read Equation Native stream");
    auto equation = streamBytes(native.Get());
    require(!equation.empty(), "Nonempty Equation Native stream");
    return equation;
}
struct NativeAsset { std::vector<unsigned char> equation; std::vector<std::vector<unsigned char>> previews; };
// 从指定对象 Range 的 Word 原生 Flat OPC 包取资产，扫描不激活 MathType 服务。
inline NativeAsset rangeNativeAsset(IDispatch* range) {
    const auto xml = utf8(text(range, L"WordOpenXML"));
    require(xml.size() < 128 * 1024 * 1024, "Object XML package bounds");
    ComPtr<IStream> stream; check(CreateStreamOnHGlobal(nullptr, TRUE, &stream), "Object XML stream");
    ULONG written{}; check(stream->Write(xml.data(), static_cast<ULONG>(xml.size()), &written), "Write object XML stream");
    require(written == xml.size(), "Complete object XML stream"); LARGE_INTEGER zero{};
    check(stream->Seek(zero, STREAM_SEEK_SET, nullptr), "Rewind object XML");
    ComPtr<IXmlReader> reader; check(CreateXmlReader(__uuidof(IXmlReader), reinterpret_cast<void**>(reader.GetAddressOf()), nullptr), "Object XML reader");
    check(reader->SetProperty(XmlReaderProperty_DtdProcessing, DtdProcessing_Prohibit), "Reject XML DTD");
    check(reader->SetInput(stream.Get()), "Read object Flat OPC");
    NativeAsset asset; std::wstring part, data; bool binary{}; XmlNodeType node{}; HRESULT status{};
    while ((status = reader->Read(&node)) == S_OK) {
        const wchar_t* local{}; check(reader->GetLocalName(&local, nullptr), "Object XML name");
        if (node == XmlNodeType_Element && wcscmp(local, L"part") == 0) {
            part.clear(); data.clear();
            for (auto attribute = reader->MoveToFirstAttribute(); attribute == S_OK; attribute = reader->MoveToNextAttribute()) {
                const wchar_t* name{}; const wchar_t* value{};
                check(reader->GetLocalName(&name, nullptr), "Object part attribute");
                check(reader->GetValue(&value, nullptr), "Object part attribute value");
                if (wcscmp(name, L"name") == 0) part = value;
            }
            check(reader->MoveToElement(), "Return to object part");
        } else if (node == XmlNodeType_Element && wcscmp(local, L"binaryData") == 0) binary = true;
        else if ((node == XmlNodeType_Text || node == XmlNodeType_Whitespace || node == XmlNodeType_CDATA) && binary) {
            const wchar_t* value{}; check(reader->GetValue(&value, nullptr), "Object base64 data"); data += value;
        } else if (node == XmlNodeType_EndElement && wcscmp(local, L"binaryData") == 0) binary = false;
        else if (node == XmlNodeType_EndElement && wcscmp(local, L"part") == 0 &&
                 (part.find(L"/embeddings/") != std::wstring::npos || part.find(L"/media/") != std::wstring::npos)) {
            require(!data.empty() && data.size() <= MAXDWORD, "Object binary data bounds"); DWORD length{};
            require(CryptStringToBinaryW(data.c_str(), static_cast<DWORD>(data.size()), CRYPT_STRING_BASE64, nullptr, &length, nullptr, nullptr) != 0,
                    "Query object base64 bytes");
            std::vector<unsigned char> bytes(length);
            require(CryptStringToBinaryW(data.c_str(), static_cast<DWORD>(data.size()), CRYPT_STRING_BASE64, bytes.data(), &length, nullptr, nullptr) != 0,
                    "Decode object base64 bytes"); bytes.resize(length);
            if (part.find(L"/embeddings/") != std::wstring::npos) {
                require(asset.equation.empty(), "Single equation embedding in target range"); asset.equation = nativeFromOle(bytes);
            } else asset.previews.push_back(std::move(bytes));
        }
    }
    require(status == S_FALSE && !asset.equation.empty(), "Complete single-object native package");
    return asset;
}
inline std::vector<unsigned char> nativePayload(const std::vector<unsigned char>& stream) {
    require(stream.size() >= 12, "Equation Native header bounds"); DWORD header{}, length{};
    memcpy(&header, stream.data(), 4); memcpy(&length, stream.data() + 8, 4);
    require(header >= 12 && header <= stream.size() && length > 0 && length == stream.size() - header && length < 1000000,
            "Equation Native header and payload length");
    return {stream.begin() + header, stream.end()};
}
// 一次导出正文 Flat OPC，按正文对象顺序关联关系和原生流，省去逐式导出整个样式包。
inline std::vector<std::vector<unsigned char>> bodyNativeAssets(IDispatch* document) {
    auto content=get(document,L"Content");const auto xml=text(content.Get(),L"WordOpenXML");
    require(xml.size()<128*1024*1024,"Body native package bounds");
    ComPtr<IXMLDOMDocument2> dom;check(CoCreateInstance(__uuidof(DOMDocument60),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&dom)),"Body native XML reader");
    check(dom->put_async(VARIANT_FALSE),"Read body XML synchronously");check(dom->put_resolveExternals(VARIANT_FALSE),"Disable body XML external resolution");
    check(dom->setProperty(Value(L"ProhibitDTD").v.bstrVal,Value(true).v),"Disable body XML DTD");
    VARIANT_BOOL loaded{};check(dom->loadXML(Value(xml.c_str()).v.bstrVal,&loaded),"Read body Flat OPC");require(loaded!=VARIANT_FALSE,"Valid body Flat OPC");
    auto attribute=[](IXMLDOMNode* node,const wchar_t* name) {
        ComPtr<IXMLDOMNamedNodeMap> attributes;check(node->get_attributes(&attributes),"Body XML attributes");
        ComPtr<IXMLDOMNode> value;check(attributes->getNamedItem(Value(name).v.bstrVal,&value),"Body XML attribute");
        require(value!=nullptr,"Required body XML attribute");Value textValue;textValue.v.vt=VT_BSTR;check(value->get_text(&textValue.v.bstrVal),"Body XML attribute text");return std::wstring(textValue.v.bstrVal);
    };
    std::map<std::wstring,std::wstring> relationships,binary;
    ComPtr<IXMLDOMNodeList> nodes;long count{};
    check(dom->selectNodes(Value(L"//*[local-name()='part' and @*[local-name()='name']='/word/_rels/document.xml.rels']//*[local-name()='Relationship']").v.bstrVal,&nodes),"Body relationships");
    check(nodes->get_length(&count),"Body relationship count");
    for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> node;check(nodes->get_item(i,&node),"Body relationship");relationships.emplace(attribute(node.Get(),L"Id"),attribute(node.Get(),L"Target"));}
    check(dom->selectNodes(Value(L"//*[local-name()='part'][*[local-name()='binaryData']]").v.bstrVal,&nodes),"Body binary parts");check(nodes->get_length(&count),"Body binary count");
    for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> node,data,name;check(nodes->get_item(i,&node),"Body binary part");
        check(node->selectSingleNode(Value(L"@*[local-name()='name']").v.bstrVal,&name),"Body binary name");
        Value key,value;key.v.vt=value.v.vt=VT_BSTR;check(name->get_text(&key.v.bstrVal),"Body binary URI");
        check(node->selectSingleNode(Value(L"*[local-name()='binaryData']").v.bstrVal,&data),"Body binary data");check(data->get_text(&value.v.bstrVal),"Body binary text");binary.emplace(key.v.bstrVal,value.v.bstrVal);
    }
    check(dom->selectNodes(Value(L"//*[local-name()='part' and @*[local-name()='name']='/word/document.xml']//*[local-name()='OLEObject' and starts-with(@ProgID,'Equation.DSMT')]").v.bstrVal,&nodes),"Body MathType objects");check(nodes->get_length(&count),"Body MathType count");
    std::vector<std::vector<unsigned char>> result;
    for(long i=0;i<count;++i){ComPtr<IXMLDOMNode> node,id;check(nodes->get_item(i,&node),"Body MathType object");
        check(node->selectSingleNode(Value(L"@*[local-name()='id' and namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/relationships']").v.bstrVal,&id),"Body object relation");
        if(!id)return {}; // 链接或不常见容器继续使用已验证的逐范围核验。
        Value key;key.v.vt=VT_BSTR;check(id->get_text(&key.v.bstrVal),"Body object relation ID");
        const auto relation=relationships.find(key.v.bstrVal);if(relation==relationships.end())return {};
        const auto target=relation->second.empty()||relation->second[0]!=L'/'?L"/word/"+relation->second:relation->second;
        const auto part=binary.find(target);if(part==binary.end())return {};
        require(part->second.size()<=MAXDWORD,"Body binary length");DWORD length{};
        require(CryptStringToBinaryW(part->second.c_str(),static_cast<DWORD>(part->second.size()),CRYPT_STRING_BASE64,nullptr,&length,nullptr,nullptr)!=0,"Query body OLE bytes");
        std::vector<unsigned char> bytes(length);require(CryptStringToBinaryW(part->second.c_str(),static_cast<DWORD>(part->second.size()),CRYPT_STRING_BASE64,bytes.data(),&length,nullptr,nullptr)!=0,"Decode body OLE bytes");bytes.resize(length);
        result.push_back(nativePayload(nativeFromOle(bytes)));
    }
    return result;
}
inline PackageState inspectPackage(const std::wstring& path,const char* progressStage=nullptr) {
    ComPtr<IOpcFactory> factory;
    check(CoCreateInstance(__uuidof(OpcFactory), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory)), "Packaging factory");
    ComPtr<IStream> stream;
    check(factory->CreateStreamOnFile(path.c_str(), OPC_STREAM_IO_READ, nullptr, 0, &stream), "Open OPC stream");
    ComPtr<IOpcPackage> package;
    check(factory->ReadPackageFromStream(stream.Get(), OPC_VALIDATE_ON_LOAD, &package), "Read OPC package");
    ComPtr<IOpcPartSet> parts;
    check(package->GetPartSet(&parts), "Package parts");
    ComPtr<IOpcPartEnumerator> enumerator;
    check(parts->GetEnumerator(&enumerator), "Package part enumerator");
    PackageState state;
    auto relationships = [&](IOpcRelationshipSet* relations) {
        ComPtr<IOpcRelationshipEnumerator> entries;
        check(relations->GetEnumerator(&entries), "Relationship enumerator");
        for (;;) {
            BOOL more{}; check(entries->MoveNext(&more), "Next relationship"); if (!more) break;
            ComPtr<IOpcRelationship> relation; check(entries->GetCurrent(&relation), "Current relationship");
            OPC_URI_TARGET_MODE mode{}; check(relation->GetTargetMode(&mode), "Relationship target mode");
            if (mode == OPC_URI_TARGET_MODE_EXTERNAL) continue; // 不联网或打开外部关系。
            ComPtr<IOpcUri> source; ComPtr<IUri> target; ComPtr<IOpcPartUri> combined;
            check(relation->GetSourceUri(&source), "Relationship source");
            check(relation->GetTargetUri(&target), "Relationship target");
            check(source->CombinePartUri(target.Get(), &combined), "Resolve internal relationship");
            BOOL exists{}; check(parts->PartExists(combined.Get(), &exists), "Relationship part exists");
            require(exists != FALSE, "Missing internal relationship part");
        }
    };
    ComPtr<IOpcRelationshipSet> relations;
    check(package->GetRelationshipSet(&relations), "Package relationships"); relationships(relations.Get());
    // 只枚举部件引用获得总数，内容仍各读一次，不为显示进度重复解析文档。
    std::vector<ComPtr<IOpcPart>> allParts;
    for (;;) {
        BOOL more{}; check(enumerator->MoveNext(&more), "Next package part"); if (!more) break;
        ComPtr<IOpcPart> part; check(enumerator->GetCurrent(&part), "Current package part");
        allParts.push_back(std::move(part));
    }
    StageProgress partProgress(progressStage,allParts.size());
    for(const auto& part:allParts) {
        ComPtr<IOpcPartUri> name; check(part->GetName(&name), "Package part URI");
        BSTR raw{}; check(name->GetAbsoluteUri(&raw), "Package part name");
        const std::wstring uri(raw, SysStringLen(raw)); SysFreeString(raw);
        ComPtr<IStream> content; check(part->GetContentStream(&content), "Package content stream");
        auto bytes = streamBytes(content.Get()); // 读取全部 part，不能只读目录就宣称容器可用。
        if(uri==L"/word/styles.xml"){
            LARGE_INTEGER zero{};check(content->Seek(zero,STREAM_SEEK_SET,nullptr),"Rewind style definitions");ComPtr<IXmlReader> reader;
            check(CreateXmlReader(__uuidof(IXmlReader),reinterpret_cast<void**>(reader.GetAddressOf()),nullptr),"Read native display style definition");check(reader->SetInput(content.Get()),"Style definition input");
            XmlNodeType node{};HRESULT status{};
            while((status=reader->Read(&node))==S_OK){if(node!=XmlNodeType_Element)continue;const wchar_t* local{};check(reader->GetLocalName(&local,nullptr),"Style element name");
                if(!wcscmp(local,L"name")&&reader->MoveToAttributeByName(L"val",L"http://schemas.openxmlformats.org/wordprocessingml/2006/main")==S_OK){const wchar_t* value{};check(reader->GetValue(&value,nullptr),"Style name");if(!wcscmp(value,L"MTDisplayEquation"))state.displayStyle=true;}}
            check(status,"Complete style definitions");
        }
        check(part->GetRelationshipSet(&relations), "Part relationships"); relationships(relations.Get());
        if (uri.find(L"/word/media/") == 0) state.media.push_back(bytes);
        if (uri.find(L"/word/embeddings/") == 0) {
            auto equation = nativeFromOle(bytes);
            if(!equation.empty())state.equations.push_back(std::move(equation));
        }
        partProgress.advance();
    }
    return state;
}
inline void requireContains(std::vector<std::vector<unsigned char>> after,
                            const std::vector<std::vector<unsigned char>>& before, const char* label) {
    // 有序多重集保留重复资产计数，避免逐项从头查找；仍按完整字节核对。
    std::multiset<std::vector<unsigned char>> remaining;
    for (auto& value : after) remaining.insert(std::move(value));
    for (const auto& value : before) {
        const auto found = remaining.find(value);
        require(found != remaining.end(), label); remaining.erase(found);
    }
}
// 厂商书签含 QUOTE 域，解除域后只留下一个 OLE，避免把模板域带进用户文档。
inline ComPtr<IDispatch> insertBlankEquation(IDispatch* document,IDispatch* range,const std::wstring& install,IDispatch* storyRoot=nullptr) {
    auto content=storyRoot?get(storyRoot,L"Duplicate"):get(document,L"Content");const ScanRangeContext context{document,storyRoot};
    const long start=integer(range,L"Start"),beforeEnd=integer(content.Get(),L"End");
    const auto templatePath=install+L"\\Office Support\\BlankEqn.doc";
    invoke(range,L"InsertFile",DISPATCH_METHOD,{templatePath.c_str(),L"MTBlankEqn",false,false,false});
    auto inserted=docRange(context,start,start+integer(content.Get(),L"End")-beforeEnd);
    auto fields=get(inserted.Get(),L"Fields");
    for(long i=integer(fields.Get(),L"Count");i>0;--i){auto field=collectionItem(fields.Get(),i);const long type=integer(field.Get(),L"Type");
        require(type==35||type==58,"Vendor blank equation field type");if(type==35)invoke(field.Get(),L"Unlink",DISPATCH_METHOD);}
    inserted=docRange(context,start,start+integer(content.Get(),L"End")-beforeEnd);
    auto shapes=get(inserted.Get(),L"InlineShapes");require(integer(shapes.Get(),L"Count")==1,"Blank template inserts one object only");
    auto shape=collectionItem(shapes.Get(),1);
    require(text(inserted.Get(),L"Text")==text(get(shape.Get(),L"Range").Get(),L"Text"),"Blank template has no extra text");return shape;
}

inline std::wstring withoutWordRevisionIds(std::wstring xml) {
    // ponytail: 仅规范化 Word 生产的 rsid 元数据；通用 XML 比较再用 XmlLite。
    // Undo 可以换 rsid，不能因此把已恢复的正文/样式误判为损坏。
    for (size_t start = xml.find(L" w:rsid"); start != std::wstring::npos; start = xml.find(L" w:rsid", start)) {
        const auto equals = xml.find(L'=', start);
        require(equals != std::wstring::npos, "Word rsid attribute");
        const auto valueStart = xml.find(L'\"', equals), end = xml.find(L'\"', valueStart + 1);
        require(valueStart != std::wstring::npos && end != std::wstring::npos, "Word rsid value");
        xml.erase(start, end + 1 - start);
    }
    return xml;
}
inline std::wstring sourceSignature(IDispatch* range, bool normalizeObjectId = false) {
    auto xml = text(range, L"WordOpenXML");
    // Range.WordOpenXML 还带全局 settings/styles；只比较本范围正文，避免全局 rsids 列表变化。
    const auto bodyStart = xml.find(L"<w:body>"), bodyEnd = xml.find(L"</w:body>");
    require(bodyStart != std::wstring::npos && bodyEnd != std::wstring::npos && bodyEnd > bodyStart, "Word range XML body");
    xml = withoutWordRevisionIds(xml.substr(bodyStart, bodyEnd + 9 - bodyStart));
    if (normalizeObjectId) {
        // Word Undo 会重新分配 OLE ObjectID；原生流和预览另作完整字节比较，不能仅比ID。
        for (size_t start = xml.find(L" ObjectID=\""); start != std::wstring::npos; start = xml.find(L" ObjectID=\"", start)) {
            const auto end = xml.find(L'\"', start + 11); require(end != std::wstring::npos, "OLE ObjectID attribute");
            xml.erase(start, end + 1 - start);
        }
    }
    return xml;
}
inline std::wstring exportedInner(const std::string& value, bool display) {
    require(std::all_of(value.begin(), value.end(), [](unsigned char c) {return c < 128;}), "UNSUPPORTED_NON_ASCII_EXPORT");
    std::wstring tex(value.begin(), value.end());
    const auto first = tex.find_first_not_of(L" \t\r\n"), last = tex.find_last_not_of(L" \t\r\n");
    require(first != std::wstring::npos, "EMPTY_NATIVE_EXPORT"); tex = tex.substr(first, last - first + 1);
    const std::wstring open = display ? L"\\[" : L"$", close = display ? L"\\]" : L"$";
    require(tex.size() >= open.size() + close.size() && tex.substr(0, open.size()) == open &&
            tex.substr(tex.size() - close.size()) == close, "Native export delimiter matches true equation type");
    auto inner = tex.substr(open.size(), tex.size() - open.size() - close.size());
    require(lexicalReason(inner).empty() && inner.find(L'%') == std::wstring::npos, "UNSUPPORTED_OR_INVALID_NATIVE_EXPORT");
    return normalizedTex(inner);
}
