// DOCX 包、外链设置和原生资产检查，不扫描或修改 Word 范围。
#pragma once
#include "word_automation.h"
#include "worker_protocol.h"
#include <msopc.h>
#include <msxml6.h>
#include <xmllite.h>
#include <wincrypt.h>
#include <cstring>
#include <map>
#include <set>

// 在同一把原件只读锁内绑定扫描和转换，计算后将句柄归零供后续读取。
inline std::string lockedSourceSha256(HANDLE source) {
    struct Digest {
        HCRYPTPROV provider{};
        HCRYPTHASH hash{};
        ~Digest(){if(hash)CryptDestroyHash(hash);if(provider)CryptReleaseContext(provider,0);}
    } digest;
    require(CryptAcquireContextW(&digest.provider,nullptr,nullptr,PROV_RSA_AES,CRYPT_VERIFYCONTEXT|CRYPT_SILENT)!=0,
            "Open SHA-256 provider");
    require(CryptCreateHash(digest.provider,CALG_SHA_256,0,0,&digest.hash)!=0,"Create source SHA-256");
    LARGE_INTEGER zero{};
    require(SetFilePointerEx(source,zero,nullptr,FILE_BEGIN)!=0,"Rewind locked source for SHA-256");
    BYTE block[64*1024];
    for(;;){
        require(!taskCancelled(),"TASK_CANCELLED");taskPulse();
        DWORD read{};
        require(ReadFile(source,block,static_cast<DWORD>(sizeof(block)),&read,nullptr)!=0,"Read locked source for SHA-256");
        if(!read)break;
        require(CryptHashData(digest.hash,block,read,0)!=0,"Update source SHA-256");
    }
    BYTE bytes[32];DWORD size=static_cast<DWORD>(sizeof(bytes));
    require(CryptGetHashParam(digest.hash,HP_HASHVAL,bytes,&size,0)!=0&&size==sizeof(bytes),"Read source SHA-256");
    require(SetFilePointerEx(source,zero,nullptr,FILE_BEGIN)!=0,"Rewind locked source after SHA-256");
    constexpr char digits[]="0123456789ABCDEF";
    std::string result;result.reserve(64);
    for(const auto value:bytes){result+=digits[value>>4];result+=digits[value&15];}
    return result;
}

inline ComPtr<IXMLDOMDocument2> xmlDocument(const std::wstring& xml) {
    ComPtr<IXMLDOMDocument2> dom;check(CoCreateInstance(__uuidof(DOMDocument60),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&dom)),"Experimental XML reader");
    check(dom->put_async(VARIANT_FALSE),"Synchronous experimental XML");check(dom->put_preserveWhiteSpace(VARIANT_TRUE),"Keep experimental XML whitespace");
    check(dom->put_resolveExternals(VARIANT_FALSE),"Disable experimental external XML");check(dom->setProperty(Value(L"ProhibitDTD").v.bstrVal,Value(true).v),"Disable experimental XML DTD");
    VARIANT_BOOL loaded{};check(dom->loadXML(Value(xml.c_str()).v.bstrVal,&loaded),"Read experimental XML");require(loaded!=VARIANT_FALSE,"Valid experimental XML");return dom;
}
using DocxLinkUpdates = std::map<std::wstring, std::vector<std::wstring>>;
// 只修改私有副本的 OLE 外链更新方式，正常保存后恢复，不修改 Word 全局选项。
inline DocxLinkUpdates setDocxLinkUpdates(const std::wstring& path, const DocxLinkUpdates* restore = nullptr, bool* formulaTabs = nullptr) {
    if (restore && restore->empty()) return {};
    const bool progress=taskShared!=nullptr;
    if(progress)std::puts(restore?"UI_PHASE RESTORE_LINKS":"UI_PHASE LINK_SETTINGS");
    ComPtr<IOpcFactory> factory; check(CoCreateInstance(__uuidof(OpcFactory),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&factory)),"Link settings package factory");
    ComPtr<IStream> input; check(factory->CreateStreamOnFile(path.c_str(),OPC_STREAM_IO_READ,nullptr,0,&input),"Read private link settings");
    ComPtr<IOpcPackage> package; check(factory->ReadPackageFromStream(input.Get(),OPC_VALIDATE_ON_LOAD,&package),"Read private package for links");
    DocxLinkUpdates updates; size_t restored{};
    {
        ComPtr<IOpcPartSet> parts;check(package->GetPartSet(&parts),"Link settings parts");
        ComPtr<IOpcPartEnumerator> entries;check(parts->GetEnumerator(&entries),"Link settings part enumeration");BOOL more{};HRESULT advanced{};
        std::vector<ComPtr<IOpcPart>> allParts;
        while((advanced=entries->MoveNext(&more))==S_OK&&more){
            ComPtr<IOpcPart> part;check(entries->GetCurrent(&part),"Link settings current part");allParts.push_back(std::move(part));
        }
        check(advanced,"Complete linked OLE part enumeration");
        StageProgress partProgress(restore?"RESTORE_LINK_PARTS":"LINK_PARTS",allParts.size(),progress);
        for(const auto& part:allParts){
            taskPulse();
            ComPtr<IOpcPartUri> uri;check(part->GetName(&uri),"Link settings part name");Value name;name.v.vt=VT_BSTR;check(uri->GetAbsoluteUri(&name.v.bstrVal),"Link settings part URI");
            const std::wstring partName=name.v.bstrVal;
            if(partName.rfind(L"/word/",0)!=0||(partName.size()<4)||(partName.substr(partName.size()-4)!=L".xml"&&partName.substr(partName.size()-4)!=L".vml")){partProgress.advance();continue;}
            if(restore&&restore->find(partName)==restore->end()){partProgress.advance();continue;}
            ComPtr<IStream> content;check(part->GetContentStream(&content),"Read XML link settings");
            ComPtr<IXMLDOMDocument2> document;check(CoCreateInstance(__uuidof(DOMDocument60),nullptr,CLSCTX_INPROC_SERVER,IID_PPV_ARGS(&document)),"System XML link settings reader");
            check(document->put_async(VARIANT_FALSE),"Synchronous link settings XML");check(document->put_preserveWhiteSpace(VARIANT_TRUE),"Keep XML whitespace");
            check(document->put_resolveExternals(VARIANT_FALSE),"Disable XML external resolution");check(document->setProperty(Value(L"ProhibitDTD").v.bstrVal,Value(true).v),"Disable XML DTD");
            Value source;source.v.vt=VT_UNKNOWN;source.v.punkVal=content.Get();source.v.punkVal->AddRef();VARIANT_BOOL loaded{};
            check(document->load(source.v,&loaded),"Load private XML part");require(loaded!=VARIANT_FALSE,"Valid private XML part");
            if(formulaTabs&&!*formulaTabs){
                ComPtr<IXMLDOMNode> tab;
                check(document->selectSingleNode(Value(L"//*[local-name()='t' and namespace-uri()='http://schemas.openxmlformats.org/officeDocument/2006/math' and contains(.,'\t')]").v.bstrVal,&tab),"Find literal OMML tabs");
                *formulaTabs=static_cast<bool>(tab);
            }
            ComPtr<IXMLDOMNodeList> links;check(document->selectNodes(Value(L"//*[local-name()='OLEObject' and namespace-uri()='urn:schemas-microsoft-com:office:office' and @Type='Link']").v.bstrVal,&links),"Find linked OLE objects");
            long count{};check(links->get_length(&count),"Count linked OLE objects");if(!count&&!restore){partProgress.advance();continue;}
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
            partProgress.advance();
        }
    }
    if(restore)require(restored==restore->size(),"All original linked OLE parts preserved");
    if(!restore&&updates.empty()){if(progress)std::puts("UI_PHASE PACKAGE_RELEASE");return updates;}
    ComPtr<IStream> output;check(CreateStreamOnHGlobal(nullptr,TRUE,&output),"Private package rewrite buffer");
    if(progress)std::puts("UI_PHASE LINK_SETTINGS_SERIALIZE");
    check(factory->WritePackageToStream(package.Get(),OPC_WRITE_DEFAULT,output.Get()),"Write private link settings package");package.Reset();input.Reset();
    LARGE_INTEGER zero{};check(output->Seek(zero,STREAM_SEEK_SET,nullptr),"Rewind private package buffer");
    if(progress)std::puts("UI_PHASE LINK_SETTINGS_WRITE");
    ComPtr<IStream> destination;check(factory->CreateStreamOnFile(path.c_str(),OPC_STREAM_IO_WRITE,nullptr,0,&destination),"Replace owned private package");
    ULARGE_INTEGER all{};all.QuadPart=~0ULL;check(output->CopyTo(destination.Get(),all,nullptr,nullptr),"Complete private package rewrite");
    if(progress)std::puts("UI_PHASE PACKAGE_RELEASE");return updates;
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
    std::vector<std::vector<unsigned char>> equations, media, otherEmbeddings;
    bool displayStyle{};
};
// WordOpenXML 的单对象包和 DOCX 的 OLE part 共用系统结构化存储读取。
inline std::vector<unsigned char> nativeFromOle(const std::vector<unsigned char>& bytes) {
    // XLSX 等嵌入包不是复合存储，原样保留，不送入 OLE 读取器。
    constexpr unsigned char signature[]{0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1};
    if(bytes.size()<sizeof(signature)||memcmp(bytes.data(),signature,sizeof(signature))!=0)return {};
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
inline std::vector<unsigned char> nativePayload(const std::vector<unsigned char>& stream) {
    require(stream.size() >= 12, "Equation Native header bounds"); DWORD header{}, length{};
    memcpy(&header, stream.data(), 4); memcpy(&length, stream.data() + 8, 4);
    require(header >= 12 && header <= stream.size() && length > 0 && length == stream.size() - header && length < 1000000,
            "Equation Native header and payload length");
    return {stream.begin() + header, stream.end()};
}
inline PackageState inspectPackage(const std::wstring& path,const char* progressStage=nullptr,bool styleOnly=false,bool progress=false) {
    if(progress)std::puts("UI_PHASE OPENING_PACKAGE");
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
    if(!styleOnly){if(progress)std::puts("UI_PHASE PACKAGE_RELATIONSHIPS");check(package->GetRelationshipSet(&relations), "Package relationships"); relationships(relations.Get());}
    // 只枚举部件引用获得总数，内容仍各读一次，不为显示进度重复解析文档。
    std::vector<ComPtr<IOpcPart>> allParts;
    if(progress)std::puts("UI_PHASE PACKAGE_ENUMERATION");
    for (;;) {
        BOOL more{}; check(enumerator->MoveNext(&more), "Next package part"); if (!more) break;
        ComPtr<IOpcPart> part; check(enumerator->GetCurrent(&part), "Current package part");
        allParts.push_back(std::move(part));
    }
    StageProgress partProgress(styleOnly?"INPUT_STYLES":progressStage,styleOnly?1:allParts.size(),progress);
    for(const auto& part:allParts) {
        ComPtr<IOpcPartUri> name; check(part->GetName(&name), "Package part URI");
        BSTR raw{}; check(name->GetAbsoluteUri(&raw), "Package part name");
        const std::wstring uri(raw, SysStringLen(raw)); SysFreeString(raw);
        if(styleOnly&&uri!=L"/word/styles.xml")continue;
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
        if(styleOnly){partProgress.advance();if(progress)std::puts("UI_PHASE PACKAGE_RELEASE");return state;}
        check(part->GetRelationshipSet(&relations), "Part relationships"); relationships(relations.Get());
        if (uri.find(L"/word/media/") == 0) state.media.push_back(std::move(bytes));
        if (uri.find(L"/word/embeddings/") == 0) {
            auto equation = nativeFromOle(bytes);
            if(!equation.empty())state.equations.push_back(std::move(equation));
            else state.otherEmbeddings.push_back(std::move(bytes));
        }
        partProgress.advance();
    }
    if(styleOnly)partProgress.advance();
    if(progress)std::puts("UI_PHASE PACKAGE_RELEASE");return state;
}
inline void requireContains(const std::vector<std::vector<unsigned char>>& after,
                            const std::vector<std::vector<unsigned char>>& before, const char* label) {
    // 只索引现有字节，不复制媒体；比较及重复数量仍按完整内容核对。
    auto compare=[](const auto* left,const auto* right){return *left<*right;};
    std::multiset<const std::vector<unsigned char>*,decltype(compare)> remaining(compare);
    for (const auto& value : after) remaining.insert(&value);
    for (const auto& value : before) {
        const auto found = remaining.find(&value);
        require(found != remaining.end(), label); remaining.erase(found);
    }
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
inline std::wstring sourceSignature(const std::wstring& package, bool normalizeObjectId = false) {
    // Range.WordOpenXML 还带全局 settings/styles；只比较本范围正文，避免全局 rsids 列表变化。
    const auto bodyStart = package.find(L"<w:body>"), bodyEnd = package.find(L"</w:body>");
    require(bodyStart != std::wstring::npos && bodyEnd != std::wstring::npos && bodyEnd > bodyStart, "Word range XML body");
    auto xml = withoutWordRevisionIds(package.substr(bodyStart, bodyEnd + 9 - bodyStart));
    if (normalizeObjectId) {
        // Word Undo 会重新分配 OLE ObjectID；原生流和预览另作完整字节比较，不能仅比ID。
        for (size_t start = xml.find(L" ObjectID=\""); start != std::wstring::npos; start = xml.find(L" ObjectID=\"", start)) {
            const auto end = xml.find(L'\"', start + 11); require(end != std::wstring::npos, "OLE ObjectID attribute");
            xml.erase(start, end + 1 - start);
        }
    }
    return xml;
}
