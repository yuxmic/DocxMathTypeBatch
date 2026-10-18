// 只挂接本工具明确创建的 Word 主线程，在窗口显示之前阻止显示和激活；组件按确定性版本缓存。
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>

static LRESULT CALLBACK HiddenWindow(HWND hwnd,UINT message,WPARAM wparam,LPARAM lparam,UINT_PTR id,DWORD_PTR data) {
    LRESULT result;
    UNREFERENCED_PARAMETER(data);
    result=DefSubclassProc(hwnd,message,wparam,lparam);
    if(message==WM_WINDOWPOSCHANGING){
        WINDOWPOS* position=(WINDOWPOS*)lparam;
        position->flags=(position->flags&~SWP_SHOWWINDOW)|SWP_HIDEWINDOW|SWP_NOACTIVATE;
    }
    if(message==WM_NCDESTROY)RemoveWindowSubclass(hwnd,HiddenWindow,id);
    return result;
}

__declspec(dllexport) LRESULT CALLBACK GuardCBT(int code,WPARAM wparam,LPARAM lparam) {
    if(code==HCBT_CREATEWND){
        CBT_CREATEWNDW* created=(CBT_CREATEWNDW*)lparam;
        if(!(created->lpcs->style&WS_CHILD)){
            created->lpcs->style&=~WS_VISIBLE;
            created->lpcs->dwExStyle|=WS_EX_NOACTIVATE|WS_EX_TOOLWINDOW;
            SetWindowSubclass((HWND)wparam,HiddenWindow,0x444D5442,0);
        }
    }else if(code==HCBT_ACTIVATE){
        if(!(GetWindowLongPtrW((HWND)wparam,GWL_STYLE)&WS_CHILD))return 1;
    }
    return CallNextHookEx(NULL,code,wparam,lparam);
}

// 在所属 Word 内安装本进程钩子，避免跨进程解析模块路径。
__declspec(dllexport) volatile DWORD GuardReadyHook=0;
__declspec(dllexport) DWORD WINAPI InstallGuard(LPVOID parameter) {
    DWORD thread=(DWORD)(UINT_PTR)parameter;
    ULONGLONG deadline=GetTickCount64()+5000;
    HHOOK hook;
    do {hook=SetWindowsHookExW(WH_CBT,GuardCBT,NULL,thread);if(!hook)Sleep(1);}while(!hook&&GetTickCount64()<deadline);
    GuardReadyHook=(DWORD)(UINT_PTR)hook;
    if(hook){MSG message;while(GetMessageW(&message,NULL,0,0)>0){TranslateMessage(&message);DispatchMessageW(&message);}UnhookWindowsHookEx(hook);}
    return 0;
}

BOOL WINAPI DllMain(HINSTANCE instance,DWORD reason,LPVOID reserved) {
    UNREFERENCED_PARAMETER(reserved);
    if(reason==DLL_PROCESS_ATTACH)DisableThreadLibraryCalls(instance);
    return TRUE;
}
