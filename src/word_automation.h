// 原生小样、只读扫描与副本批处理共用的 Word Automation；本头文件不依赖 MathType。
#pragma once
#define NOMINMAX
#include <windows.h>
#include <oleauto.h>
#include <oleidl.h>
#include <tlhelp32.h>
#include <wrl/client.h>
#include <algorithm>
#include <cstdio>
#include <initializer_list>
#include <memory>
#include <stdexcept>
#include <string>
#include <vector>
#include <cwctype>

using Microsoft::WRL::ComPtr;

// VARIANT 的所有权局限在一次 Automation 调用，不借用悬空字符串或对象。
struct Value {
    VARIANT v{};
    Value() { v.vt = VT_ERROR; v.scode = DISP_E_PARAMNOTFOUND; }
    Value(long n) { v.vt = VT_I4; v.lVal = n; }
    Value(double n) { v.vt = VT_R8; v.dblVal = n; }
    Value(bool b) { v.vt = VT_BOOL; v.boolVal = b ? VARIANT_TRUE : VARIANT_FALSE; }
    Value(const wchar_t* s) {
        v.vt = VT_BSTR; v.bstrVal = SysAllocString(s);
        if (!v.bstrVal) throw std::bad_alloc();
    }
    Value(IDispatch* p) { v.vt = VT_DISPATCH; v.pdispVal = p; if (p) p->AddRef(); }
    Value(const Value& other) {
        if (FAILED(VariantCopy(&v, &other.v))) throw std::bad_alloc();
    }
    Value(Value&& other) noexcept : v(other.v) { VariantInit(&other.v); }
    Value& operator=(Value&& other) noexcept {
        if (this != &other) { VariantClear(&v); v = other.v; VariantInit(&other.v); }
        return *this;
    }
    ~Value() { VariantClear(&v); }
};

static void require(bool ok, const char* label) {
    if (!ok) throw std::runtime_error(label);
}

static void check(HRESULT status, const char* label) {
    if (FAILED(status)) {
        std::printf("HRESULT %s 0x%08lX\n", label, static_cast<unsigned long>(status));
        throw std::runtime_error(label);
    }
}

#include "task_guard.h"

// 忙碌重试只覆盖明确被 COM 拒绝的调用，不重试可能已经产生修改的普通错误。
static void pump(DWORD milliseconds) {
    const ULONGLONG end = taskActiveMilliseconds() + milliseconds;
    do {
        MSG message{};
        while (PeekMessageW(&message, nullptr, 0, 0, PM_REMOVE)) {
            TranslateMessage(&message); DispatchMessageW(&message);
        }
        MsgWaitForMultipleObjects(0, nullptr, FALSE, 25, QS_ALLINPUT);
    } while (taskActiveMilliseconds() < end);
}

static Value invoke(IDispatch* object, const wchar_t* name, WORD flags,
                    std::initializer_list<Value> arguments = {}) {
    DISPID id{};
    taskPulse();
    LPOLESTR member = const_cast<wchar_t*>(name);
    // Word 正在加载加载项或处理其他调用时会拒绝 GetIDsOfNames；与 Invoke 一样等待重试。
    HRESULT resolved{};
    for (int attempt = 0; attempt != 20; ++attempt) {
        resolved = object->GetIDsOfNames(IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &id);
        if (resolved != RPC_E_CALL_REJECTED && resolved != RPC_E_SERVERCALL_RETRYLATER) break;
        pump(100);
    }
    if (FAILED(resolved)) std::fwprintf(stderr, L"UNKNOWN_MEMBER %ls\n", name);
    check(resolved, "GetIDsOfNames");
    std::vector<Value> values(arguments);
    std::reverse(values.begin(), values.end());
    std::vector<VARIANT> raw;
    for (const auto& value : values) raw.push_back(value.v);
    DISPID put = DISPID_PROPERTYPUT;
    DISPPARAMS parameters{raw.data(), flags == DISPATCH_PROPERTYPUT ? &put : nullptr,
                         static_cast<UINT>(raw.size()), flags == DISPATCH_PROPERTYPUT ? 1u : 0u};
    Value result;
    VariantClear(&result.v);
    HRESULT status{};
    for (int attempt = 0; attempt != 20; ++attempt) {
        taskPulse();
        EXCEPINFO exception{};
        UINT badArgument{};
        status = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &parameters,
                                &result.v, &exception, &badArgument);
        if (exception.pfnDeferredFillIn) exception.pfnDeferredFillIn(&exception);
        if (FAILED(status) && status != RPC_E_CALL_REJECTED && status != RPC_E_SERVERCALL_RETRYLATER)
            std::wprintf(L"COM_MEMBER %ls: %ls\n", name, exception.bstrDescription ? exception.bstrDescription : L"");
        SysFreeString(exception.bstrSource); SysFreeString(exception.bstrDescription);
        SysFreeString(exception.bstrHelpFile);
        if (status != RPC_E_CALL_REJECTED && status != RPC_E_SERVERCALL_RETRYLATER) break;
        VariantClear(&result.v);
        pump(100);
    }
    check(status, "IDispatch::Invoke");
    taskPulse();
    return result;
}

static ComPtr<IDispatch> asObject(const Value& value) {
    ComPtr<IDispatch> result;
    if (value.v.vt == VT_DISPATCH && value.v.pdispVal)
        check(value.v.pdispVal->QueryInterface(IID_PPV_ARGS(&result)), "QueryInterface IDispatch");
    else if (value.v.vt == VT_UNKNOWN && value.v.punkVal)
        check(value.v.punkVal->QueryInterface(IID_PPV_ARGS(&result)), "QueryInterface IDispatch");
    else throw std::runtime_error("Automation did not return an object");
    return result;
}

static ComPtr<IDispatch> get(IDispatch* object, const wchar_t* name) {
    return asObject(invoke(object, name, DISPATCH_PROPERTYGET));
}

static long integer(IDispatch* object, const wchar_t* name) {
    auto value = invoke(object, name, DISPATCH_PROPERTYGET);
    require(value.v.vt == VT_I4, "Automation integer type");
    return value.v.lVal;
}

static std::wstring text(IDispatch* object, const wchar_t* name) {
    auto value = invoke(object, name, DISPATCH_PROPERTYGET);
    require(value.v.vt == VT_BSTR, "Automation string type");
    if (SysStringLen(value.v.bstrVal) == 0) return {};
    return std::wstring(value.v.bstrVal, SysStringLen(value.v.bstrVal));
}

// 判断 Word 实例是否为本任务新建，避免对已有用户会话执行 Quit。
static std::vector<DWORD> wordProcesses() {
    std::vector<DWORD> ids;
    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    require(snapshot != INVALID_HANDLE_VALUE, "Process snapshot");
    PROCESSENTRY32W entry{sizeof(entry)};
    for (BOOL more = Process32FirstW(snapshot, &entry); more; more = Process32NextW(snapshot, &entry))
        if (_wcsicmp(entry.szExeFile, L"WINWORD.EXE") == 0) ids.push_back(entry.th32ProcessID);
    CloseHandle(snapshot);
    return ids;
}

struct WordSession {
    ComPtr<IDispatch> application;
    ComPtr<IDispatch> document;
    std::vector<DWORD> before;
    bool owned{};
    ComPtr<IDispatch> linkOptions;
    bool originalUpdateLinks{};
    WordSession() : before(wordProcesses()) {
        taskPulse();
        CLSID id{};
        check(CLSIDFromProgID(L"Word.Application", &id), "Word registration");
        if (!taskShared) { check(CoCreateInstance(id, nullptr, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&application)), "Start Word"); return; }
        // 明确创建 Word 后才取得 COM 对象，关闭启动阶段尚无 HWND 时的进程归属空隙。
        wchar_t clsid[40]{}; StringFromGUID2(id, clsid, 40);
        const auto key = L"CLSID\\" + std::wstring(clsid) + L"\\LocalServer32";
        wchar_t registered[32768]{}; DWORD bytes = sizeof(registered);
        require(RegGetValueW(HKEY_CLASSES_ROOT, key.c_str(), nullptr, RRF_RT_REG_SZ, nullptr, registered, &bytes) == ERROR_SUCCESS, "Registered Word executable path");
        const std::wstring registration = registered;
        auto folded = registration; std::transform(folded.begin(), folded.end(), folded.begin(), [](wchar_t character) { return static_cast<wchar_t>(towlower(character)); });
        const auto start = registration.find(L'\"'), end = start == std::wstring::npos ? folded.find(L".exe") : registration.find(L'\"', start + 1);
        require(end != std::wstring::npos, "Parse registered Word executable");
        const auto program = start == std::wstring::npos ? registration.substr(0, end + 4) : registration.substr(start + 1, end - start - 1);
        const auto slash = program.find_last_of(L"\\/");
        require(slash != std::wstring::npos && !_wcsicmp(program.c_str() + slash + 1, L"WINWORD.EXE"), "Word registration names WINWORD.EXE");
        // /w 提供独立空白文档的原生对象，避免命令行打开 RTF 时 Word 再次显示窗口。
        // /a 禁止加载全局模板和加载项，/q 禁用启动画面；后续文档显式以不可见方式打开。
        auto command = taskQuote(program) + L" /Automation /q /a /w";STARTUPINFOEXW startup{};
        startup.StartupInfo.cb=sizeof(startup);startup.StartupInfo.dwFlags=STARTF_USESHOWWINDOW;startup.StartupInfo.wShowWindow=SW_HIDE;
        TaskHandle creationJob(OpenJobObjectW(JOB_OBJECT_ASSIGN_PROCESS,FALSE,taskShared->job));require(creationJob.value!=nullptr,"Open Word creation job");
        SIZE_T attributeBytes{};InitializeProcThreadAttributeList(nullptr,1,0,&attributeBytes);std::vector<unsigned char> attributeStorage(attributeBytes);
        startup.lpAttributeList=reinterpret_cast<LPPROC_THREAD_ATTRIBUTE_LIST>(attributeStorage.data());
        require(InitializeProcThreadAttributeList(startup.lpAttributeList,1,0,&attributeBytes)!=0,"Word creation attributes");
        const bool assigned=UpdateProcThreadAttribute(startup.lpAttributeList,0,PROC_THREAD_ATTRIBUTE_JOB_LIST,&creationJob.value,sizeof(HANDLE),nullptr,nullptr)!=0;
        PROCESS_INFORMATION information{};
        // 创建时加入 Job，消除取消恰好发生在创建后、登记前而遗留 Word 的窗口。
        const bool launched=assigned&&CreateProcessW(program.c_str(),command.data(),nullptr,nullptr,FALSE,CREATE_SUSPENDED|EXTENDED_STARTUPINFO_PRESENT,nullptr,nullptr,&startup.StartupInfo,&information)!=0;
        DeleteProcThreadAttributeList(startup.lpAttributeList);require(launched,"Create Word atomically inside task job");
        TaskHandle process(information.hProcess), thread(information.hThread); FILETIME created{}, exited{}, kernel{}, user{};
        require(GetProcessTimes(process.value, &created, &exited, &kernel, &user), "Read explicit Word creation time");
        ULARGE_INTEGER identity{}; identity.LowPart = created.dwLowDateTime; identity.HighPart = created.dwHighDateTime;
        taskOwnWord(information.dwProcessId, identity.QuadPart);
        std::printf("OWNED_WORD_PID=%lu\nOWNED_WORD_STARTED=%llu\n", information.dwProcessId, identity.QuadPart);
        // 使用系统启动隐藏和 Word 自身可见性设置，不向 Word 注入代码。
        require(ResumeThread(thread.value)!=MAXDWORD,"Start explicit task Word");
        taskTestBoundary(L"WORD_STARTUP");
        // 启动时可能先出现一个尚不提供原生对象的隐藏 _WwG，必须检查全部同属窗口并有限重试。
        struct WindowSearch { DWORD pid; std::vector<HWND> natives; } search{information.dwProcessId, {}};
        HMODULE accessibility = LoadLibraryExW(L"oleacc.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
        require(accessibility != nullptr, "Load system Office native-object bridge");
        using NativeObject = HRESULT (WINAPI*)(HWND, DWORD, REFIID, void**);
        const auto nativeObject = reinterpret_cast<NativeObject>(GetProcAddress(accessibility, "AccessibleObjectFromWindow"));
        require(nativeObject != nullptr, "Office native-object bridge export");
        const auto deadline=taskActiveMilliseconds()+TaskDefaultTimeoutSeconds*1000ULL;
        while (!application && taskActiveMilliseconds() < deadline && WaitForSingleObject(process.value, 0) == WAIT_TIMEOUT) {
            search.natives.clear();
            EnumWindows([](HWND window, LPARAM parameter) -> BOOL {
                auto& context = *reinterpret_cast<WindowSearch*>(parameter); DWORD pid{}; GetWindowThreadProcessId(window, &pid);
                if (pid != context.pid) return TRUE;
                wchar_t kind[80]{}; GetClassNameW(window, kind, 80);
                if (wcscmp(kind, L"OpusApp")) return TRUE; // 原生文档窗口可能尚未填入标题。
                EnumChildWindows(window, [](HWND child, LPARAM childParameter) -> BOOL {
                    auto& context = *reinterpret_cast<WindowSearch*>(childParameter); wchar_t name[80]{}; GetClassNameW(child, name, 80);
                    if (!wcscmp(name, L"_WwG")) context.natives.push_back(child); return TRUE;
                }, parameter);
                if (!context.natives.empty()) ShowWindowAsync(window, SW_HIDE); // 不向初始化中的辅助线程同步发隐藏消息。
                return TRUE;
            }, reinterpret_cast<LPARAM>(&search));
            for (const auto handle : search.natives) {
                ComPtr<IDispatch> window;
                if (SUCCEEDED(nativeObject(handle, 0xFFFFFFF0, IID_PPV_ARGS(&window))) && window) { application = get(window.Get(), L"Application"); break; }
            }
            if (!application) pump(50);
        }
        FreeLibrary(accessibility);
        require(application!=nullptr,"Task Word native object starts within deadline");
        invoke(application.Get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
        invoke(application.Get(), L"ScreenUpdating", DISPATCH_PROPERTYPUT, {false});
        // 微软定义 _WwG + OBJID_NATIVEOM 返回 Word Window，只连接已创建 PID 的对象。
        std::puts("OWNED_WORD_NATIVE_BRIDGE_READY");
    }
    void verifyOwnership() {
        // Word 的 Hwnd 属于 Window 对象，不属于 Application；用自建文档定位宿主进程。
        auto window = get(document.Get(), L"ActiveWindow");
        DWORD pid{};
        GetWindowThreadProcessId(reinterpret_cast<HWND>(static_cast<LONG_PTR>(integer(window.Get(), L"Hwnd"))), &pid);
        owned = pid != 0 && std::find(before.begin(), before.end(), pid) == before.end();
        if (taskShared) require(pid == taskShared->wordPid, "Word HWND matches explicitly created process");
        require(owned, "Refuse existing Word session");
        std::printf("OWNED_WORD_PID=%lu\n", pid);
        // 启动时间用于外部故障验收核对进程身份，不能只凭可复用的 PID 操作 Office。
        HANDLE process = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
        require(process != nullptr, "Query owned Word process identity");
        FILETIME created{}, exited{}, kernel{}, user{};
        const bool timed = GetProcessTimes(process, &created, &exited, &kernel, &user) != 0;
        CloseHandle(process); require(timed, "Read owned Word creation time");
        ULARGE_INTEGER started{}; started.LowPart = created.dwLowDateTime; started.HighPart = created.dwHighDateTime;
        taskOwnWord(pid, started.QuadPart);
        std::printf("OWNED_WORD_STARTED=%llu\n", started.QuadPart);
    }
    void disableLinkUpdates() {
        require(owned, "Own Word before changing link option");
        auto options = get(application.Get(), L"Options");
        const auto original = invoke(options.Get(), L"UpdateLinksAtOpen", DISPATCH_PROPERTYGET);
        require(original.v.vt == VT_BOOL, "Word link option type");
        originalUpdateLinks = original.v.boolVal != VARIANT_FALSE;
        taskLinksBeforeChange(originalUpdateLinks);
        linkOptions = options; // 先记录恢复责任，设置失败也尝试恢复原值。
        invoke(linkOptions.Get(), L"UpdateLinksAtOpen", DISPATCH_PROPERTYPUT, {false});
    }
    void shutdown() {
        if(!linkOptions)std::puts("OWNED_WORD_OPTIONS_UNCHANGED_PASSED");
        // Options 可能持久化到用户设置，关闭会话前恢复，不能把全局偏好留在关闭状态。
        if (linkOptions) {
            invoke(linkOptions.Get(), L"UpdateLinksAtOpen", DISPATCH_PROPERTYPUT, {originalUpdateLinks});
            taskLinksRestored();
            linkOptions.Reset();
            std::puts("OWNED_WORD_OPTIONS_RESTORED_PASSED");
        }
        if (document) {
            invoke(document.Get(), L"Close", DISPATCH_METHOD, {0L});
            document.Reset();
        }
        if (owned) {
            const auto started=GetTickCount64();
            invoke(application.Get(), L"Quit", DISPATCH_METHOD, {0L});
            owned = false;
            std::printf("WORD_QUIT_MS=%llu\n",GetTickCount64()-started);
            std::puts("OWNED_WORD_QUIT_PASSED");
        }
    }
    ~WordSession() {
        try { shutdown(); }
        catch (...) { std::puts("OWNED_WORD_QUIT_FAILED"); }
        // 退出后的代理释放可能等待 Word；单独计时，不改变原有释放责任。
        const auto started=GetTickCount64();
        linkOptions.Reset();document.Reset();application.Reset();
        std::printf("WORD_COM_RELEASE_MS=%llu\n",GetTickCount64()-started);
    }
};

static ComPtr<IDispatch> cellRange(IDispatch* table, long row, long column) {
    return get(asObject(invoke(table, L"Cell", DISPATCH_METHOD, {row, column})).Get(), L"Range");
}

// 只用于自建小样的黑色三线表；转换现有文档时不能整表重设格式。
static void blackBorder(IDispatch* owner, long side, long width) {
    auto border = asObject(invoke(get(owner, L"Borders").Get(), L"Item", DISPATCH_METHOD, {side}));
    invoke(border.Get(), L"LineStyle", DISPATCH_PROPERTYPUT, {1L});
    invoke(border.Get(), L"Color", DISPATCH_PROPERTYPUT, {0L});
    invoke(border.Get(), L"LineWidth", DISPATCH_PROPERTYPUT, {width});
}
