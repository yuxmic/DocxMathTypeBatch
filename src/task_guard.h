// 产品任务监管：父进程不调用 COM，限时结束所属工作进程/Word，并回收登记的副本。
#pragma once
#include <cwchar>
#include <cstring>
#include <limits>
#include <new>

constexpr DWORD TaskTimeoutExit = 124, TaskServiceLostExit = 125, TaskRecoveryExit = 126, TaskCancelExit = 127;
constexpr DWORD TaskDefaultTimeoutSeconds = 120, TaskRecoverySeconds = 30;
constexpr size_t TaskMaxFiles = 16, TaskNativeCapacity = 65536;
// 只累计系统运行时间，主动睡眠和休眠不会误触发无进展超时。
inline ULONGLONG taskActiveMilliseconds() {
    ULONGLONG ticks{}; require(QueryUnbiasedInterruptTime(&ticks) != 0, "Read active task clock");
    return ticks / 10000;
}
// 跨进程心跳可能在父进程采样 now 后更新；未来心跳按0等待计算，避免无符号减法溢出。
constexpr ULONGLONG taskElapsed(ULONGLONG now, ULONGLONG before) { return now >= before ? now - before : 0; }
static_assert(taskElapsed(100, 101) == 0 && taskElapsed(100, 100) == 0 && taskElapsed(100, 80) == 20, "Cross-process heartbeat elapsed bounds");
struct TaskRecoveryRecord {
    DWORD magic{0x44584D54}, version{1}, nativeLength{};
    wchar_t install[32768]{};
    char defaults[32768]{};
    unsigned char native[TaskNativeCapacity]{};
};
struct TaskTrackedFile {
    wchar_t path[32768]{};
    DWORD volume{}, indexHigh{}, indexLow{};
    bool finalFile{};
};
struct TaskSharedState {
    alignas(8) volatile LONG64 heartbeat{};
    volatile LONG files{}, wordReady{}, mathReady{}, pendingPrefs{}, recordSaved{}, linksChanged{}, committed{}, cancelRequested{};
    DWORD wordPid{}, mathPid{}; ULONGLONG wordStarted{}, mathStarted{}; bool originalLinks{}, scanTask{};
    wchar_t job[160]{}, recoveryPath[32768]{};
    TaskRecoveryRecord recovery;
    DWORD expectedPendingLength{}; unsigned char expectedPendingNative[TaskNativeCapacity]{};
    TaskTrackedFile tracked[TaskMaxFiles];
};
inline TaskSharedState* taskShared{};
inline volatile LONG taskConsoleStop{};

struct TaskHandle {
    HANDLE value{};
    explicit TaskHandle(HANDLE handle = nullptr) : value(handle) {}
    ~TaskHandle() { if (value && value != INVALID_HANDLE_VALUE) CloseHandle(value); }
    TaskHandle(const TaskHandle&) = delete;
    TaskHandle& operator=(const TaskHandle&) = delete;
};
struct TaskMapping {
    TaskHandle handle; TaskSharedState* state{};
    TaskMapping(const std::wstring& name, bool create) : handle(create ?
        CreateFileMappingW(INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, sizeof(TaskSharedState), name.c_str()) :
        OpenFileMappingW(FILE_MAP_ALL_ACCESS, FALSE, name.c_str())) {
        require(handle.value != nullptr, "Open task supervision mapping");
        if (create) require(GetLastError() != ERROR_ALREADY_EXISTS, "Unique task supervision mapping");
        state = static_cast<TaskSharedState*>(MapViewOfFile(handle.value, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(TaskSharedState)));
        require(state != nullptr, "Map task supervision state");
        if (create) { new (state) TaskSharedState(); InterlockedExchange64(&state->heartbeat, static_cast<LONG64>(taskActiveMilliseconds())); }
    }
    ~TaskMapping() { if (state) UnmapViewOfFile(state); }
};
inline void taskPulse() {
    if (!taskShared) return;
    InterlockedExchange64(&taskShared->heartbeat, static_cast<LONG64>(taskActiveMilliseconds()));
    require(InterlockedCompareExchange(&taskShared->cancelRequested, 0, 0) == 0, "TASK_CANCELLED");
}
inline bool taskCancelled() { return taskShared && InterlockedCompareExchange(&taskShared->cancelRequested, 0, 0) != 0; }
inline BOOL WINAPI taskConsoleHandler(DWORD event) {
    if (event != CTRL_C_EVENT && event != CTRL_BREAK_EVENT) return FALSE;
    InterlockedExchange(&taskConsoleStop, 1); return TRUE;
}
inline std::wstring taskQuote(const std::wstring& value) {
    std::wstring result = L"\""; size_t slashes{};
    for (const auto character : value) {
        if (character == L'\\') ++slashes;
        else { result.append(slashes * (character == L'\"' ? 2 : 1) + (character == L'\"' ? 1 : 0), L'\\'); result += character; slashes = 0; }
    }
    result.append(slashes * 2, L'\\'); return result + L"\"";
}
inline std::string taskHex(const std::string& value) {
    constexpr char digits[] = "0123456789ABCDEF"; std::string result;
    for (const unsigned char character : value) { result += digits[character >> 4]; result += digits[character & 15]; }
    return result.empty() ? "-" : result;
}
inline void taskError(const char* value) { std::printf("UI_ERROR %s\n", taskHex(value).c_str()); }
inline void taskOwnWord(DWORD pid, ULONGLONG started) {
    if (!taskShared) return;
    // 明确 CreateProcess 创建或经 HWND/启动前快照核实的 Word，以固定身份加入独立 Job。
    TaskHandle process(OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SET_QUOTA | PROCESS_TERMINATE | SYNCHRONIZE, FALSE, pid));
    require(process.value != nullptr, "Hold owned Word process");
    FILETIME created{}, exited{}, kernel{}, user{};
    require(GetProcessTimes(process.value, &created, &exited, &kernel, &user) != 0, "Read guarded Word identity");
    ULARGE_INTEGER identity{}; identity.LowPart = created.dwLowDateTime; identity.HighPart = created.dwHighDateTime;
    require(identity.QuadPart == started, "Guarded Word creation time matches");
    taskShared->wordPid = pid; taskShared->wordStarted = started;
    InterlockedExchange(&taskShared->wordReady, 1);
    TaskHandle job(OpenJobObjectW(JOB_OBJECT_ASSIGN_PROCESS | JOB_OBJECT_QUERY, FALSE, taskShared->job)); BOOL assigned{};
    require(job.value && IsProcessInJob(process.value, job.value, &assigned) && (assigned || AssignProcessToJobObject(job.value, process.value)), "Assign only owned Word to task job");
}
inline void taskLinksBeforeChange(bool original) {
    if (taskShared) { taskShared->originalLinks = original; InterlockedExchange(&taskShared->linksChanged, 1); }
}
inline void taskLinksRestored() { if (taskShared) InterlockedExchange(&taskShared->linksChanged, 0); }
inline void taskMathTypeConnected() {
    if (!taskShared) return;
    // 服务可能共享，只记录唯一 MathType 的身份用于失联检测，绝不将其加入 Job 或结束它。
    TaskHandle snapshot(CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)); require(snapshot.value != INVALID_HANDLE_VALUE, "Snapshot MathType service");
    PROCESSENTRY32W entry{sizeof(entry)}; std::vector<DWORD> ids;
    for (BOOL more = Process32FirstW(snapshot.value, &entry); more; more = Process32NextW(snapshot.value, &entry))
        if (!_wcsicmp(entry.szExeFile, L"MathType.exe")) ids.push_back(entry.th32ProcessID);
    if (ids.size() != 1) return;
    TaskHandle process(OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, FALSE, ids[0]));
    FILETIME created{}, exited{}, kernel{}, user{}; if (!process.value || !GetProcessTimes(process.value, &created, &exited, &kernel, &user)) return;
    ULARGE_INTEGER identity{}; identity.LowPart = created.dwLowDateTime; identity.HighPart = created.dwHighDateTime;
    taskShared->mathPid = ids[0]; taskShared->mathStarted = identity.QuadPart; InterlockedExchange(&taskShared->mathReady, 1);
    std::printf("SHARED_MATHTYPE_PID=%lu STARTED=%llu\n", ids[0], identity.QuadPart);
}
inline void taskTrackFile(const std::wstring& path, HANDLE file, bool finalFile) {
    if (!taskShared) return;
    const auto index = InterlockedCompareExchange(&taskShared->files, 0, 0);
    require(index >= 0 && static_cast<size_t>(index) < TaskMaxFiles, "Task owned-file journal capacity");
    BY_HANDLE_FILE_INFORMATION identity{}; require(GetFileInformationByHandle(file, &identity) != 0, "Track owned-file identity");
    auto& entry = taskShared->tracked[index];
    require(wcscpy_s(entry.path, path.c_str()) == 0, "Track complete owned-file path");
    entry.volume = identity.dwVolumeSerialNumber; entry.indexHigh = identity.nFileIndexHigh; entry.indexLow = identity.nFileIndexLow; entry.finalFile = finalFile;
    InterlockedExchange(&taskShared->files, index + 1);
}
// 私有路径先登记再创建，强制取消落在创建/登记之间也能清理。
inline void taskReservePrivatePath(const std::wstring& path) {
    if(!taskShared)return;require(GetFileAttributesW(path.c_str())==INVALID_FILE_ATTRIBUTES,"Private task path is unoccupied");
    const auto index=InterlockedCompareExchange(&taskShared->files,0,0);require(index>=0&&static_cast<size_t>(index)<TaskMaxFiles,"Private file journal capacity");
    wcscpy_s(taskShared->tracked[index].path,path.c_str());taskShared->tracked[index].finalFile=false;InterlockedExchange(&taskShared->files,index+1);
}
inline void taskCopy(const std::wstring& input, const std::wstring& output) {
    if (!taskShared) { require(CopyFileW(input.c_str(), output.c_str(), TRUE) != 0, "Create independent task copy"); return; }
    // 先独占创建再登记，后续只填充本任务拥有的文件，避免覆盖别人创建的路径。
    taskReservePrivatePath(output);
    TaskHandle file(CreateFileW(output.c_str(), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                               nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
    require(file.value != INVALID_HANDLE_VALUE, "Reserve independent task copy");
    require(CopyFileW(input.c_str(), output.c_str(), FALSE) != 0, "Fill reserved task copy");
}
inline void taskTrackWordOwner(const std::wstring& document) {
    if (!taskShared) return;
    const auto count = InterlockedCompareExchange(&taskShared->files, 0, 0);
    require(count >= 0 && static_cast<size_t>(count) < TaskMaxFiles, "Word owner-file journal capacity");
    bool registered{}; for (LONG index = 0; index < count; ++index)
        if (!taskShared->tracked[index].finalFile && document == taskShared->tracked[index].path) registered = true;
    require(registered, "Owner file belongs to reserved task copy");
    const auto slash = document.find_last_of(L"\\/"); require(slash != std::wstring::npos && document.size() > slash + 3, "Task copy basename");
    const auto owner = document.substr(0, slash + 1) + L"~$" + document.substr(slash + 3);
    require(GetFileAttributesW(owner.c_str()) == INVALID_FILE_ATTRIBUTES, "Refuse pre-existing Word owner file");
    // Word 在 Open 返回前就可能创建 ~$ 锁文件，提前登记它与私有副本的唯一对应关系。
    auto& entry = taskShared->tracked[count]; wcscpy_s(entry.path, owner.c_str()); entry.finalFile = false;
    InterlockedExchange(&taskShared->files, count + 1);
}
inline void taskMove(const std::wstring& source, const std::wstring& destination) {
    TaskHandle file(CreateFileW(source.c_str(), FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                               nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    require(file.value != INVALID_HANDLE_VALUE, "Hold staged publication identity"); taskTrackFile(destination, file.value, true);
    require(MoveFileExW(source.c_str(), destination.c_str(), MOVEFILE_WRITE_THROUGH) != 0, "Publish verified file without overwrite");
}
inline void taskCommit() { taskPulse();if (taskShared) InterlockedExchange(&taskShared->committed, 1); }
inline void taskBaseline(const std::wstring& install, const std::string& defaults, const std::vector<unsigned char>& native) {
    if (!taskShared) return;
    auto& record = taskShared->recovery;
    require(!native.empty() && native.size() <= TaskNativeCapacity && defaults.size() < sizeof(record.defaults), "Preference recovery baseline bounds");
    require(wcscpy_s(record.install, install.c_str()) == 0 && strcpy_s(record.defaults, defaults.c_str()) == 0, "Store preference recovery baseline");
    record.nativeLength = static_cast<DWORD>(native.size()); std::copy(native.begin(), native.end(), record.native);
}
inline void taskPending(bool pending) {
    if (!taskShared) return;
    if (pending && !InterlockedCompareExchange(&taskShared->recordSaved, 0, 0)) {
        require(taskShared->recovery.nativeLength > 0, "Recovery baseline precedes NEXT_EQN");
        // 在设置共享 NEXT_EQN 之前持久化；父进程异常退出后，下个任务仍会先恢复。
        TaskHandle file(CreateFileW(taskShared->recoveryPath, GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
        require(file.value != INVALID_HANDLE_VALUE, "Create pending-preference recovery marker");
        InterlockedExchange(&taskShared->recordSaved, 1); // 短写发生在设置偏好之前，父进程仍可清理自有标记。
        DWORD written{}; require(WriteFile(file.value, &taskShared->recovery, sizeof(TaskRecoveryRecord), &written, nullptr) &&
            written == sizeof(TaskRecoveryRecord) && FlushFileBuffers(file.value), "Persist pending-preference recovery marker");
        InterlockedExchange(&taskShared->recordSaved, 1);
    }
    InterlockedExchange(&taskShared->pendingPrefs, pending ? 1 : 0); taskPulse();
}
inline void taskExpectedPending(const std::vector<unsigned char>& native) {
    if (!taskShared) return;
    require(native.size() <= TaskNativeCapacity, "Pending preference style bounds");
    std::copy(native.begin(), native.end(), taskShared->expectedPendingNative); taskShared->expectedPendingLength = static_cast<DWORD>(native.size());
}
inline void taskTestBoundary(const wchar_t* name) {
    // 仅开发故障入口在已完成真实 API 操作后等待外部介入；生产任务不设置该环境变量。
    wchar_t boundary[80]{}; if (!GetEnvironmentVariableW(L"DOCX_MATH_FAULT_BOUNDARY", boundary, 80) || wcscmp(boundary, name)) return;
    std::wprintf(L"FAULT_BOUNDARY=%ls\n", name); fflush(stdout);
    wchar_t gate[32768]{};
    require(GetEnvironmentVariableW(L"DOCX_MATH_FAULT_GATE", gate, 32768) > 0, "Development fault gate path");
    while (GetFileAttributesW(gate) == INVALID_FILE_ATTRIBUTES) {require(!taskCancelled(),"TASK_CANCELLED");Sleep(20);}
    taskPulse();
}
inline std::wstring taskRecoveryPath() {
    wchar_t directory[32768]{};
    if (!GetEnvironmentVariableW(L"DOCX_MATH_RECOVERY_DIRECTORY", directory, 32768)) {
        require(GetEnvironmentVariableW(L"LOCALAPPDATA", directory, 32768) > 0, "Local application data directory");
        const auto root = std::wstring(directory) + L"\\DocxMathTypeBatch";
        require(CreateDirectoryW(root.c_str(), nullptr) || GetLastError() == ERROR_ALREADY_EXISTS, "Create task recovery directory");
        return root + L"\\pending-preferences.bin";
    }
    require(GetFileAttributesW(directory) != INVALID_FILE_ATTRIBUTES, "Explicit recovery directory exists");
    return std::wstring(directory) + L"\\pending-preferences.bin";
}
inline DWORD taskTimeoutSeconds() {
    wchar_t value[24]{};
    if (!GetEnvironmentVariableW(L"DOCX_MATH_TIMEOUT_SECONDS", value, 24)) return TaskDefaultTimeoutSeconds;
    wchar_t* end{}; const auto seconds = wcstoul(value, &end, 10);
    require(*value && !*end && seconds >= 5 && seconds <= 3600, "Task timeout must be 5..3600 seconds");
    return seconds;
}
inline TaskHandle* taskVerifiedWord(TaskSharedState& state, bool terminating) {
    if (!InterlockedCompareExchange(&state.wordReady, 0, 0)) return nullptr;
    auto process = std::make_unique<TaskHandle>(OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_TERMINATE | SYNCHRONIZE, FALSE, state.wordPid));
    if (!process->value) return nullptr;
    if (WaitForSingleObject(process->value, 0) == WAIT_OBJECT_0) return nullptr; // 已退出进程的映像查询可能失败，无需再操作。
    FILETIME created{}, exited{}, kernel{}, user{}; wchar_t image[32768]{}; DWORD size = 32768;
    require(GetProcessTimes(process->value, &created, &exited, &kernel, &user), "Query supervisor Word creation time");
    ULARGE_INTEGER identity{}; identity.LowPart = created.dwLowDateTime; identity.HighPart = created.dwHighDateTime;
    require(identity.QuadPart == state.wordStarted, "Refuse mismatched task Word creation time");
    if (!terminating) {
        if (!QueryFullProcessImageNameW(process->value, 0, image, &size) && WaitForSingleObject(process->value, 1000) == WAIT_OBJECT_0) return nullptr;
        const auto slash = wcsrchr(image, L'\\');
        require(slash && !_wcsicmp(slash + 1, L"WINWORD.EXE"), "Refuse mismatched task Word image");
    } // Job 已按登记身份请求退出时只等待，不重复查询正在卸载的映像。
    return process.release();
}
inline void taskEndBlockedWorker(TaskSharedState& state, HANDLE job, HANDLE worker, DWORD code) {
    // 先让任务 Word 断连，给后端一次有界的异常展开/MTTermAPI 机会，避免硬杀积累 SDK 会话。
    std::unique_ptr<TaskHandle> word(taskVerifiedWord(state, false));
    if (word && WaitForSingleObject(word->value, 0) == WAIT_TIMEOUT) {
        const bool requested = TerminateProcess(word->value, code) != 0;
        const auto exited = WaitForSingleObject(word->value, 5000);
        require((requested || exited == WAIT_OBJECT_0) && exited == WAIT_OBJECT_0, "Blocked owned Word exits");
    }
    if (WaitForSingleObject(worker, 3000) == WAIT_OBJECT_0) std::puts("BLOCKED_WORKER_UNWOUND_PASSED");
    else { require(TerminateJobObject(job, code), "End only isolated blocked task job"); std::puts("BLOCKED_WORKER_HARD_EXIT_PASSED"); }
}
inline bool taskCleanupFiles(TaskSharedState& state) {
    bool complete = true; const auto count = InterlockedCompareExchange(&state.files, 0, 0);
    require(count >= 0 && static_cast<size_t>(count) <= TaskMaxFiles, "Valid owned-file journal count");
    for (LONG index = 0; index < count; ++index) {
        const auto& entry = state.tracked[index]; if (entry.finalFile && state.committed) continue;
        if (GetFileAttributesW(entry.path) == INVALID_FILE_ATTRIBUTES) continue;
        if (entry.finalFile) {
            // 发布过程中遇到重名失败时，绝不能回滚原来属于用户的目标文件。
            TaskHandle file(CreateFileW(entry.path, FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
            BY_HANDLE_FILE_INFORMATION identity{};
            if (file.value == INVALID_HANDLE_VALUE || !GetFileInformationByHandle(file.value, &identity) || identity.dwVolumeSerialNumber != entry.volume ||
                identity.nFileIndexHigh != entry.indexHigh || identity.nFileIndexLow != entry.indexLow) continue;
        }
        bool removed{}; for (int attempt = 0; attempt != 40; ++attempt) {
            if (DeleteFileW(entry.path) || GetLastError() == ERROR_FILE_NOT_FOUND) { removed = true; break; }
            Sleep(50);
        }
        if (!removed) { complete = false; std::printf("TASK_COPY_CLEANUP_FAILED WIN32=%lu\n", GetLastError()); }
    }
    if (complete) std::puts("TASK_COPY_CLEANUP_PASSED"); return complete;
}

// 只纠正已核验的任务 Word 抢前台，目标始终跟随用户最近选择的可见窗口。
struct TaskForegroundGuard {
    TaskSharedState& state;std::unique_ptr<TaskHandle> word;HWND userWindow{};DWORD userPid{},userThread{};HWINEVENTHOOK hook{};
    TaskForegroundGuard* previous{};inline static thread_local TaskForegroundGuard* current{};
    unsigned restored{},failed{};
    void observe(HWND window) {
        if(!window)return;DWORD pid{};const auto thread=GetWindowThreadProcessId(window,&pid);
        if(state.wordReady&&pid==state.wordPid){
            DWORD targetPid{};const auto targetThread=GetWindowThreadProcessId(userWindow,&targetPid);
            if(word&&WaitForSingleObject(word->value,0)==WAIT_TIMEOUT&&GetForegroundWindow()==window&&
               targetPid==userPid&&targetThread==userThread&&IsWindow(userWindow)&&IsWindowVisible(userWindow)&&IsWindowEnabled(userWindow)&&!IsIconic(userWindow)){
                if(SetForegroundWindow(userWindow))++restored;else ++failed;
            }
        }else if(IsWindowVisible(window)&&IsWindowEnabled(window)&&!IsIconic(window)){userWindow=window;userPid=pid;userThread=thread;}
    }
    static void CALLBACK changed(HWINEVENTHOOK source,DWORD,HWND window,LONG,LONG,DWORD,DWORD) {
        for(auto guard=current;guard;guard=guard->previous)if(guard->hook==source){guard->observe(window);return;}
    }
    explicit TaskForegroundGuard(TaskSharedState& shared):state(shared) {
        userWindow=GetForegroundWindow();userThread=GetWindowThreadProcessId(userWindow,&userPid);previous=current;
        hook=SetWinEventHook(EVENT_SYSTEM_FOREGROUND,EVENT_SYSTEM_FOREGROUND,nullptr,changed,0,0,WINEVENT_OUTOFCONTEXT);
        require(hook!=nullptr,"Start task foreground protection");current=this;
        std::puts("TASK_FOREGROUND_GUARD_READY");
    }
    void poll() {
        if(!word&&InterlockedCompareExchange(&state.wordReady,0,0))word.reset(taskVerifiedWord(state,false));
        MSG message{};while(PeekMessageW(&message,nullptr,0,0,PM_REMOVE)){TranslateMessage(&message);DispatchMessageW(&message);}
        observe(GetForegroundWindow());
    }
    ~TaskForegroundGuard(){observe(GetForegroundWindow());UnhookWinEvent(hook);current=previous;
        std::printf("TASK_FOREGROUND_RESTORED=%u FAILED_REQUESTS=%u\n",restored,failed);}
};
struct TaskAwake {
    // 允许熄屏和锁屏，仅阻止任务期间自动睡眠；嵌套恢复结束后保留外层请求。
    EXECUTION_STATE previous{SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)};
    TaskAwake() { require(previous != 0, "Keep active task awake"); }
    ~TaskAwake() { SetThreadExecutionState(previous); }
};
inline int superviseTaskArguments(const std::vector<std::wstring>& arguments, bool mathTypeTask = true);
inline bool taskRecover(TaskMapping& original, const std::wstring& originalName, bool preferences) {
    if (!preferences && !original.state->linksChanged) return true;
    std::puts(preferences ? "PREFERENCE_RECOVERY_BEGIN" : "WORD_OPTIONS_RECOVERY_BEGIN");
    const int code = superviseTaskArguments({L"--worker-recover", originalName});
    return code == 0 && (!preferences || !original.state->pendingPrefs) && !original.state->linksChanged;
}
inline int superviseTaskArguments(const std::vector<std::wstring>& arguments, bool mathTypeTask) {
    setvbuf(stdout, nullptr, _IONBF, 0); DWORD code = 1;
    require(!arguments.empty(), "Task worker mode");
    TaskAwake awake;
    const bool recoveryOnly = arguments[0] == L"--worker-recover";
    const DWORD timeout = recoveryOnly ? (std::min)(TaskRecoverySeconds, taskTimeoutSeconds()) : taskTimeoutSeconds();
    const bool usesMathType = mathTypeTask && !recoveryOnly && arguments[0] != L"--scan" && arguments[0] != L"--worker-scan";
    TaskHandle mutex(usesMathType ? CreateMutexW(nullptr, FALSE, L"Local\\DocxMathTypeBatch.MathTypePreferences.v1") : nullptr);
    bool mutexOwned{};
    struct Unlock { TaskHandle& mutex; bool& owned; ~Unlock() { if (owned) ReleaseMutex(mutex.value); } } unlock{mutex, mutexOwned};
    if (usesMathType) {
        require(mutex.value != nullptr, "Open MathType task mutex");
        const auto acquired = WaitForSingleObject(mutex.value, 0);
        if (acquired != WAIT_OBJECT_0 && acquired != WAIT_ABANDONED) { taskError("MATHTYPE_TASK_BUSY"); return 1; } mutexOwned = true;
    }
    GUID id{}; check(CoCreateGuid(&id), "Create unique supervision identity"); wchar_t guid[40]{}; StringFromGUID2(id, guid, 40);
    const auto mappingName = L"Local\\DocxMathTypeBatch-" + std::wstring(guid);
    TaskMapping mapping(mappingName, true); auto& state = *mapping.state;
    TaskForegroundGuard foreground(state);
    const auto jobName = mappingName + L"-job"; wcscpy_s(state.job, jobName.c_str());
    state.scanTask = arguments[0] == L"--scan" || arguments[0] == L"--worker-scan";
    const auto recoveryPath = taskRecoveryPath(); wcscpy_s(state.recoveryPath, recoveryPath.c_str());
    if (usesMathType && GetFileAttributesW(recoveryPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        TaskHandle file(CreateFileW(recoveryPath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        DWORD read{}; LARGE_INTEGER size{};
        require(file.value != INVALID_HANDLE_VALUE && GetFileSizeEx(file.value, &size) && size.QuadPart == sizeof(TaskRecoveryRecord) &&
            ReadFile(file.value, &state.recovery, sizeof(TaskRecoveryRecord), &read, nullptr) && read == sizeof(TaskRecoveryRecord), "Read pending-preference recovery marker");
        require(state.recovery.magic == 0x44584D54 && state.recovery.version == 1 && state.recovery.nativeLength && state.recovery.nativeLength <= TaskNativeCapacity &&
            wmemchr(state.recovery.install, 0, 32768) && memchr(state.recovery.defaults, 0, 32768), "Validate pending-preference recovery marker");
        CloseHandle(file.value); file.value = nullptr; state.pendingPrefs = 1;
        if (!taskRecover(mapping, mappingName, true)) { taskError("PREFERENCE_RECOVERY_REQUIRED"); return TaskRecoveryExit; }
        require(DeleteFileW(recoveryPath.c_str()) != 0, "Remove verified recovered preference marker");
    }
    TaskHandle job(CreateJobObjectW(nullptr, jobName.c_str())); require(job.value != nullptr, "Create isolated task job");
    JOBOBJECT_EXTENDED_LIMIT_INFORMATION limits{};
    // 禁止自动把 Word/工作进程启动的共享 MathType 加入 Job，只显式加入已核实的 Word。
    limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE | JOB_OBJECT_LIMIT_SILENT_BREAKAWAY_OK;
    require(SetInformationJobObject(job.value, JobObjectExtendedLimitInformation, &limits, sizeof(limits)), "Configure isolated task job");
    HANDLE rawRead{}, rawWrite{}; SECURITY_ATTRIBUTES security{sizeof(security), nullptr, TRUE};
    require(CreatePipe(&rawRead, &rawWrite, &security, 0), "Create supervised worker pipe");
    TaskHandle read(rawRead), write(rawWrite), nul(CreateFileW(L"NUL", GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, &security, OPEN_EXISTING, 0, nullptr));
    require(SetHandleInformation(read.value, HANDLE_FLAG_INHERIT, 0) && nul.value != INVALID_HANDLE_VALUE, "Supervised worker pipe inheritance");
    SIZE_T bytes{}; InitializeProcThreadAttributeList(nullptr, 1, 0, &bytes); std::vector<unsigned char> storage(bytes);
    auto attributes = reinterpret_cast<LPPROC_THREAD_ATTRIBUTE_LIST>(storage.data());
    require(InitializeProcThreadAttributeList(attributes, 1, 0, &bytes), "Create supervisor handle whitelist");
    HANDLE handles[]{write.value, nul.value};
    const bool configured = UpdateProcThreadAttribute(attributes, 0, PROC_THREAD_ATTRIBUTE_HANDLE_LIST, handles, sizeof(handles), nullptr, nullptr) != 0;
    STARTUPINFOEXW start{}; start.StartupInfo.cb = sizeof(start); start.StartupInfo.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    start.StartupInfo.wShowWindow = SW_HIDE; start.StartupInfo.hStdInput = nul.value; start.StartupInfo.hStdOutput = write.value; start.StartupInfo.hStdError = write.value; start.lpAttributeList = attributes;
    wchar_t program[32768]{}; require(GetModuleFileNameW(nullptr, program, 32768) > 0, "Supervisor executable path");
    auto command = taskQuote(program) + L" --task-run " + taskQuote(mappingName);
    for (const auto& argument : arguments) command += L" " + taskQuote(argument);
    PROCESS_INFORMATION information{};
    const bool started = configured && CreateProcessW(program, command.data(), nullptr, nullptr, TRUE,
        CREATE_SUSPENDED | CREATE_NO_WINDOW | EXTENDED_STARTUPINFO_PRESENT, nullptr, nullptr, &start.StartupInfo, &information);
    DeleteProcThreadAttributeList(attributes); require(started, "Create supervised native worker");
    TaskHandle process(information.hProcess), thread(information.hThread);
    if (!AssignProcessToJobObject(job.value, process.value)) { TerminateProcess(process.value, 1); throw std::runtime_error("Assign native worker to task job"); }
    require(ResumeThread(thread.value) != MAXDWORD, "Start supervised native worker");
    CloseHandle(write.value); write.value = nullptr; CloseHandle(nul.value); nul.value = nullptr;
    std::wstring stopPath;
    if ((arguments[0] == L"--worker-convert" || arguments[0] == L"--worker-reverse") && arguments.size() >= 6) stopPath = arguments[5];
    if (arguments[0] == L"--worker-scan" && arguments.size() >= 5) stopPath = arguments[4];
    wchar_t guiCommit[32768]{};
    const bool guiPublication=!stopPath.empty()&&(arguments[0]==L"--worker-convert"||arguments[0]==L"--worker-reverse")&&
        GetEnvironmentVariableW(L"DOCX_MATH_GUI_COMMIT",guiCommit,32768)>0;
    std::printf("TASK_SUPERVISOR_PID=%lu WORKER_PID=%lu TIMEOUT_SECONDS=%lu\n", GetCurrentProcessId(), information.dwProcessId, timeout);
    if (!recoveryOnly) { InterlockedExchange(&taskConsoleStop, 0); SetConsoleCtrlHandler(taskConsoleHandler, TRUE); }
    bool terminated{},cancelled{},commitRequested{};ULONGLONG commitWaiting{}; TaskHandle sharedMath;
    try {
        for (;;) {
            // 取消优先于进度转发与提交；已移动但未交付的文件仍按本次身份回滚。
            if(InterlockedCompareExchange(&taskConsoleStop,0,0)||InterlockedCompareExchange(&state.cancelRequested,0,0)||
               (!stopPath.empty()&&GetFileAttributesW(stopPath.c_str())!=INVALID_FILE_ATTRIBUTES)){
                InterlockedExchange(&state.cancelRequested,1);InterlockedExchange(&state.committed,0);
                // Word 立即结束；最多 200 ms 仅让后端释放 MathType 接口，随后结束整个 Job。
                std::unique_ptr<TaskHandle> cancelWord(taskVerifiedWord(state,false));
                if(cancelWord&&WaitForSingleObject(cancelWord->value,0)==WAIT_TIMEOUT)TerminateProcess(cancelWord->value,TaskCancelExit);
                WaitForSingleObject(process.value,200);
                require(TerminateJobObject(job.value,TaskCancelExit)!=0,"Cancel only owned task job");
                code=TaskCancelExit;terminated=cancelled=true;std::puts("TASK_CANCEL_REQUESTED");break;
            }
            foreground.poll();
            // 本监管进程只有这一线程读取管道；仅按可用字节数读取，绝不等待工作进程输出。
            DWORD available{}; if (PeekNamedPipe(read.value, nullptr, 0, nullptr, &available, nullptr) && available) {
                char buffer[4096]{}; DWORD count{};
                if (ReadFile(read.value, buffer, (std::min)(available, static_cast<DWORD>(sizeof(buffer))), &count, nullptr) && count) fwrite(buffer, 1, count, stdout);
            }
            if (WaitForSingleObject(process.value, 0) == WAIT_OBJECT_0) {
                DWORD remaining{}; if (PeekNamedPipe(read.value, nullptr, 0, nullptr, &remaining, nullptr) && remaining) continue;
                require(GetExitCodeProcess(process.value, &code), "Read supervised worker outcome");
                if(guiPublication&&code==0&&state.committed){
                    if(!commitRequested){std::puts("UI_COMMIT_READY");commitRequested=true;commitWaiting=taskActiveMilliseconds();}
                    if(GetFileAttributesW(guiCommit)==INVALID_FILE_ATTRIBUTES){
                        if(taskElapsed(taskActiveMilliseconds(),commitWaiting)>=timeout*1000ULL){InterlockedExchange(&state.committed,0);taskError("GUI_COMMIT_TIMEOUT");code=TaskTimeoutExit;break;}
                        Sleep(20);continue;
                    }
                }
                break;
            }
            const auto now = taskActiveMilliseconds();
            if (!sharedMath.value && InterlockedCompareExchange(&state.mathReady, 0, 0)) {
                sharedMath.value = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, FALSE, state.mathPid);
                if (sharedMath.value) {
                    FILETIME created{}, exited{}, kernel{}, user{}; ULARGE_INTEGER identity{};
                    require(GetProcessTimes(sharedMath.value, &created, &exited, &kernel, &user), "Read shared MathType identity");
                    identity.LowPart = created.dwLowDateTime; identity.HighPart = created.dwHighDateTime;
                    require(identity.QuadPart == state.mathStarted, "Shared MathType creation time matches");
                }
                if (!sharedMath.value) { taskError("MATHTYPE_SERVICE_DISCONNECTED"); code = TaskServiceLostExit; terminated = true; taskEndBlockedWorker(state, job.value, process.value, code); break; }
            }
            if (sharedMath.value && InterlockedCompareExchange(&state.mathReady, 0, 0) && WaitForSingleObject(sharedMath.value, 0) == WAIT_OBJECT_0) {
                taskError("MATHTYPE_SERVICE_DISCONNECTED"); std::puts("SHARED_MATHTYPE_DISCONNECT_DETECTED");
                code = TaskServiceLostExit; terminated = true; taskEndBlockedWorker(state, job.value, process.value, code); break;
            }
            const auto idle = taskElapsed(now, static_cast<ULONGLONG>(InterlockedCompareExchange64(&state.heartbeat, 0, 0)));
            if (idle >= timeout * 1000ULL) {
                taskError("TASK_TIMEOUT"); std::printf("TASK_HARD_TIMEOUT_PASSED TIMEOUT_SECONDS=%lu\n", timeout);
                code = TaskTimeoutExit; terminated = true; taskEndBlockedWorker(state, job.value, process.value, code); break;
            }
            MsgWaitForMultipleObjects(0,nullptr,FALSE,20,QS_ALLINPUT);
        }
    } catch (...) { TerminateJobObject(job.value, 1); WaitForSingleObject(process.value, 5000); throw; }
    if (!recoveryOnly) SetConsoleCtrlHandler(taskConsoleHandler, FALSE);
    require(WaitForSingleObject(process.value, 5000) == WAIT_OBJECT_0, "Supervised worker exits after bounded termination");
    std::unique_ptr<TaskHandle> word(taskVerifiedWord(state, terminated));
    if (word && WaitForSingleObject(word->value, 0) == WAIT_TIMEOUT) {
        // Job 终止是异步请求；不要对已经进入退出流程的 Word 再次终止并把 ACCESS_DENIED 当失败。
        const bool requested = TerminateProcess(word->value, code ? code : 1) != 0;
        const auto exited = WaitForSingleObject(word->value, 5000);
        require((requested || exited == WAIT_OBJECT_0) && exited == WAIT_OBJECT_0, "Verified task Word exits");
        std::puts("OWNED_WORD_HARD_EXIT_PASSED");
    }
    if (!taskCleanupFiles(state)) {taskError("TASK_COPY_CLEANUP_FAILED");code = 1;}
    if (!recoveryOnly&&!cancelled) {
        const bool pending = InterlockedCompareExchange(&state.pendingPrefs, 0, 0) != 0;
        if (!taskRecover(mapping, mappingName, pending)) { taskError("PREFERENCE_RECOVERY_REQUIRED"); if (!code) code = TaskRecoveryExit; }
        else if (state.recordSaved && GetFileAttributesW(recoveryPath.c_str()) != INVALID_FILE_ATTRIBUTES) require(DeleteFileW(recoveryPath.c_str()), "Clear resolved task preference marker");
    }
    if (terminated) std::puts("TASK_FAILURE_CLEANUP_FINISHED");
    if(cancelled&&code==TaskCancelExit)std::puts("UI_CANCELLED");
    return static_cast<int>(code);
}
inline int superviseTask(int argc, wchar_t** argv, bool mathTypeTask = true) {
    try { std::vector<std::wstring> arguments; for (int index = 1; index < argc; ++index) arguments.emplace_back(argv[index]); return superviseTaskArguments(arguments, mathTypeTask); }
    catch (const std::exception& error) { taskError(error.what()); std::printf("TASK_SUPERVISOR_FAILED: %s\n", error.what()); return 1; }
}
