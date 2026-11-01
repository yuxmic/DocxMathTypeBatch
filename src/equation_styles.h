// 新公式样式准备与偏好恢复共用；不保存任务配置。
#pragma once
#include "equation_objects.h"
#include "mathtype_api.h"
#include <cmath>

inline double number(IDispatch* object, const wchar_t* property) {
    auto value = invoke(object, property, DISPATCH_PROPERTYGET);
    if (value.v.vt == VT_R4) return value.v.fltVal;
    if (value.v.vt == VT_R8) return value.v.dblVal;
    if (value.v.vt == VT_I4) return value.v.lVal;
    throw std::runtime_error("Automation numeric type");
}
inline bool boolProperty(IDispatch* object, const wchar_t* property) {
    auto value = invoke(object, property, DISPATCH_PROPERTYGET);
    require(value.v.vt == VT_BOOL, "Automation Boolean type"); return value.v.boolVal != VARIANT_FALSE;
}
inline void requireStyle(const EquationStyle& actual, const EquationStyle& expected) {
    require(actual.inlineEquation == expected.inlineEquation && actual.sizes == expected.sizes && actual.styles == expected.styles,
            "New equation native style matches task EQP");
}

// 使用厂商转换器在内存中应用样式，不设置共享服务的 NEXT_EQN 偏好。
inline std::vector<unsigned char> styledNative(MathTypeApi& api,std::vector<unsigned char> source,const std::string& prefs,bool display) {
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
inline std::string sizedPreferences(MathTypeApi& api,double points,const std::wstring& path,bool& owned) {
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
inline std::vector<unsigned char> emptyEquation(MathTypeApi& api, WordSession& word,
                                               const std::string* prefs = nullptr, bool display = true,bool fast=false) {
    require(word.owned, "Own scratch Word session");
    auto scratch = asObject(invoke(get(word.application.Get(), L"Documents").Get(), L"Add", DISPATCH_METHOD, {Value(), false, Value(), false}));
    ComPtr<IDispatch> object;
    bool active = false;
    try {
        // 从厂商空公式书签插入 OLE，避免 Word.AddOLEObject 自动执行交互式 Open。
        auto blankRange=docRange(scratch.Get(),0,0);
        auto shape=insertBlankEquation(scratch.Get(),blankRange.Get(),api.install,nullptr,fast);
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
