// 独立 MathType 对象工厂：只复用已保存的空模板，不复用活动对象。
#pragma once
#include "mathtype_api.h"
#include "conversion_timing.h"
class EquationSite final : public IOleClientSite {
    ULONG references{1};
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,void** result) override{if(!result)return E_POINTER;*result=nullptr;if(iid==IID_IUnknown||iid==IID_IOleClientSite){*result=static_cast<IOleClientSite*>(this);AddRef();return S_OK;}return E_NOINTERFACE;}
    ULONG STDMETHODCALLTYPE AddRef() override{return ++references;}
    ULONG STDMETHODCALLTYPE Release() override{const auto left=--references;if(!left)delete this;return left;}
    HRESULT STDMETHODCALLTYPE SaveObject() override{return S_OK;}
    HRESULT STDMETHODCALLTYPE GetMoniker(DWORD,DWORD,IMoniker** result) override{*result=nullptr;return E_NOTIMPL;}
    HRESULT STDMETHODCALLTYPE GetContainer(IOleContainer** result) override{*result=nullptr;return E_NOINTERFACE;}
    HRESULT STDMETHODCALLTYPE ShowObject() override{return S_OK;}
    HRESULT STDMETHODCALLTYPE OnShowWindow(BOOL) override{return S_OK;}
    HRESULT STDMETHODCALLTYPE RequestNewObjectLayout() override{return E_NOTIMPL;}
};
struct Asset {std::vector<unsigned char> ole,preview;long cx{},cy{},baselineHalfPoints{};};
inline long previewBaseline(const std::vector<unsigned char>& bytes){
    // 厂商 WMF 的 MathType/0 注释以 1/16 pt 给出底边至基线距离；Word 使用带符号的半磅。
    auto word=[&](size_t at){unsigned short value{};memcpy(&value,bytes.data()+at,2);return value;};
    for(size_t at=40;at+6<=bytes.size();){
        unsigned long count{};memcpy(&count,bytes.data()+at,4);const auto length=static_cast<size_t>(count)*2;
        require(length>=6&&length<=bytes.size()-at,"Valid vendor WMF record");
        if(word(at+4)==META_ESCAPE&&length>=22&&word(at+6)==MFCOMMENT&&word(at+8)>=12&&
           memcmp(bytes.data()+at+10,"MathType",8)==0&&word(at+18)==0)return -static_cast<long>((word(at+20)+4)/8);
        at+=length;
    }
    throw std::runtime_error("Vendor preview has no equation baseline");
}
struct UltraServiceError : std::runtime_error {UltraServiceError():std::runtime_error("MATHTYPE_SERVICE_DISCONNECTED") {}};
inline void checkOle(HRESULT status,const char* label){
    if(status==RPC_E_DISCONNECTED||status==RPC_E_SERVER_DIED||status==RPC_E_SERVER_DIED_DNE||status==CO_E_OBJNOTCONNECTED||status==HRESULT_FROM_WIN32(RPC_S_SERVER_UNAVAILABLE))throw UltraServiceError();
    check(status,label);
}
struct EquationHost {
    ComPtr<IOleClientSite> site;ComPtr<ILockBytes> bytes;ComPtr<IStorage> storage;ComPtr<IOleObject> ole;ComPtr<IDispatch> object;ComPtr<IPersistStorage> persistent;
    ConversionTimings& closeTimings;
    explicit EquationHost(ConversionTimings& timings,IStorage* blank=nullptr):closeTimings(timings){
        site.Attach(new EquationSite());checkOle(CreateILockBytesOnHGlobal(nullptr,TRUE,&bytes),"Allocate ultra OLE bytes");checkOle(StgCreateDocfileOnILockBytes(bytes.Get(),STGM_CREATE|STGM_READWRITE|STGM_SHARE_EXCLUSIVE,0,&storage),"Create ultra OLE storage");
        if(blank){ScopedConversionTiming timing(timings,ConversionMetric::OleLoad);checkOle(blank->CopyTo(0,nullptr,nullptr,storage.Get()),"Clone styled ultra storage");checkOle(OleLoad(storage.Get(),IID_IOleObject,site.Get(),reinterpret_cast<void**>(ole.GetAddressOf())),"Load independent ultra template");}
        else {ScopedConversionTiming timing(timings,ConversionMetric::OleCreate);CLSID cls{};checkOle(CLSIDFromProgID(L"Equation.DSMT4",&cls),"MathType OLE class");checkOle(OleCreate(cls,IID_IOleObject,OLERENDER_DRAW,nullptr,site.Get(),storage.Get(),reinterpret_cast<void**>(ole.GetAddressOf())),"Create standalone MathType");}
        ScopedConversionTiming timing(timings,ConversionMetric::ObjectActivate);
        checkOle(ole->SetHostNames(L"DocxMathTypeBatch Ultra",L"Formula factory"),"Ultra host name");RECT rect{0,0,100,100};checkOle(ole->DoVerb(2,nullptr,site.Get(),0,nullptr,&rect),"Activate ultra formula");checkOle(ole.As(&object),"Ultra MathType dispatch");checkOle(ole.As(&persistent),"Ultra persistent object");}
    ~EquationHost(){if(ole){ScopedConversionTiming timing(closeTimings,ConversionMetric::OleClose);ole->Close(OLECLOSE_NOSAVE);}}
    Asset save(ConversionTimings& timings){Asset asset;
        {ScopedConversionTiming timing(timings,ConversionMetric::OleStore);checkOle(OleSave(persistent.Get(),storage.Get(),TRUE),"Save ultra OLE");checkOle(persistent->SaveCompleted(nullptr),"Complete ultra OLE save");STATSTG state{};checkOle(bytes->Stat(&state,STATFLAG_NONAME),"Ultra OLE size");require(state.cbSize.HighPart==0,"Ultra OLE size bounds");asset.ole.resize(state.cbSize.LowPart);ULARGE_INTEGER zero{};ULONG read{};checkOle(bytes->ReadAt(zero,asset.ole.data(),static_cast<ULONG>(asset.ole.size()),&read),"Read ultra OLE");require(read==asset.ole.size(),"Complete ultra OLE read");}
        {ScopedConversionTiming timing(timings,ConversionMetric::PreviewExport);ComPtr<IDataObject> source;checkOle(ole.As(&source),"Ultra preview source");FORMATETC format{CF_METAFILEPICT,nullptr,DVASPECT_CONTENT,-1,TYMED_MFPICT};
        struct Medium {STGMEDIUM value{};~Medium(){ReleaseStgMedium(&value);}} medium;checkOle(source->GetData(&format,&medium.value),"Read ultra preview");auto picture=static_cast<METAFILEPICT*>(GlobalLock(medium.value.hMetaFilePict));require(picture!=nullptr,"Lock ultra preview");
        struct Unlock {HGLOBAL value;~Unlock(){GlobalUnlock(value);}} unlock{medium.value.hMetaFilePict};const auto count=GetMetaFileBitsEx(picture->hMF,0,nullptr);asset.preview.resize(22+count);require(count&&GetMetaFileBitsEx(picture->hMF,count,asset.preview.data()+22)==count,"Read ultra metafile");asset.cx=picture->xExt;asset.cy=picture->yExt;
        // WMF 预览来自厂商，只补标准 placeable 文件头。
        const unsigned short header[]{0xcdd7,0x9ac6,0,0,0,static_cast<unsigned short>(MulDiv(asset.cx,1440,2540)),static_cast<unsigned short>(MulDiv(asset.cy,1440,2540)),1440,0,0,0};memcpy(asset.preview.data(),header,sizeof(header));unsigned short checksum{};for(size_t i=0;i<10;++i)checksum^=header[i];memcpy(asset.preview.data()+20,&checksum,2);}
        asset.baselineHalfPoints=previewBaseline(asset.preview);
        return asset;
    }
};
inline ComPtr<IStorage> styledStorage(MathTypeApi& api,const std::vector<unsigned char>& native,ConversionTimings& timings){
    EquationHost host(timings);{
        ScopedConversionTiming timing(timings,ConversionMetric::StyleWrite);
        require(api.writeBytes(host.object.Get(),3,native.data(),static_cast<long>(native.size()))==0,"Prepare styled ultra template");
    }
    {ScopedConversionTiming timing(timings,ConversionMetric::OleStore);checkOle(OleSave(host.persistent.Get(),host.storage.Get(),TRUE),"Save styled ultra template");checkOle(host.persistent->SaveCompleted(nullptr),"Complete styled ultra template");}
    return host.storage;
}
