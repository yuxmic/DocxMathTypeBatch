// Word 公式对象的插入、原生资产读取和范围快照。
#pragma once
#include "docx_scan.h"
#include "docx_package.h"
#include <xmllite.h>
#include <wincrypt.h>
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
// 厂商书签含 QUOTE 域，解除域后只留下一个 OLE，避免把模板域带进用户文档。
inline ComPtr<IDispatch> insertBlankEquation(IDispatch* document,IDispatch* range,const std::wstring& install,IDispatch* storyRoot=nullptr,bool fast=false) {
    auto content=storyRoot?get(storyRoot,L"Duplicate"):get(document,L"Content");const ScanRangeContext context{document,storyRoot};
    const long start=integer(range,L"Start"),beforeEnd=integer(content.Get(),L"End");
    const auto templatePath=install+L"\\Office Support\\BlankEqn.doc";
    invoke(range,L"InsertFile",DISPATCH_METHOD,{templatePath.c_str(),L"MTBlankEqn",false,false,false});
    auto inserted=docRange(context,start,start+integer(content.Get(),L"End")-beforeEnd);
    auto fields=get(inserted.Get(),L"Fields");
    for(long i=integer(fields.Get(),L"Count");i>0;--i){auto field=collectionItem(fields.Get(),i);const long type=integer(field.Get(),L"Type");
        require(fast||type==35||type==58,"Vendor blank equation field type");if(type==35)invoke(field.Get(),L"Unlink",DISPATCH_METHOD);}
    inserted=docRange(context,start,start+integer(content.Get(),L"End")-beforeEnd);
    auto shapes=get(inserted.Get(),L"InlineShapes");if(!fast)require(integer(shapes.Get(),L"Count")==1,"Blank template inserts one object only");
    auto shape=collectionItem(shapes.Get(),1);
    if(!fast)require(text(inserted.Get(),L"Text")==text(get(shape.Get(),L"Range").Get(),L"Text"),"Blank template has no extra text");return shape;
}

inline std::wstring sourceSignature(IDispatch* range, bool normalizeObjectId = false) {
    return sourceSignature(text(range,L"WordOpenXML"),normalizeObjectId);
}
