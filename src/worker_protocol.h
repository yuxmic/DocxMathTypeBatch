// 工作进程的文本编码、行协议与阶段进度；调用方显式决定是否报告进度。
#pragma once
#include "word_automation.h"
#include <climits>

inline std::string utf8(const std::wstring& value) {
    if (value.empty()) return {};
    require(value.size() <= INT_MAX, "UTF-8 input bounds");
    const int count = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, value.data(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    require(count > 0, "UTF-8 length");
    std::string output(static_cast<size_t>(count), '\0');
    require(WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, value.data(), static_cast<int>(value.size()), output.data(), count, nullptr, nullptr) == count, "UTF-8 encoding");
    return output;
}
inline std::string quote(const std::wstring& value) {
    std::string output = "\"";
    for (const unsigned char c : utf8(value)) {
        if (c == '"' || c == '\\') { output += '\\'; output += static_cast<char>(c); }
        else if (c < 32) {
            char escaped[7]{}; sprintf_s(escaped, "\\u%04X", static_cast<unsigned int>(c)); output += escaped;
        } else output += static_cast<char>(c);
    }
    return output + '"';
}
// 工作进程只发送 UTF-8 文本；十六进制字段避免源式中的控制符破坏行协议。
inline std::string wireHex(const std::string& bytes) {
    if (bytes.empty()) return "-";
    constexpr char digits[] = "0123456789ABCDEF";
    std::string result;
    for (const unsigned char value : bytes) { result += digits[value >> 4]; result += digits[value & 15]; }
    return result;
}

// 只在 GUI 工作进程输出计数；阶段首尾立即发送，中途最多每 150 ms 更新一次。
struct StageProgress {
    const char* stage{};size_t total{},completed{};ULONGLONG lastUpdate{};bool enabled{};
    StageProgress(const char* name,size_t count,bool reportProgress=false):stage(name),total(count),enabled(reportProgress) {
        if(enabled&&stage&&total){std::printf("UI_STAGE_PROGRESS %s 0 %zu\n",stage,total);lastUpdate=GetTickCount64();}
    }
    void advance(){++completed;if(completed==1&&stage){
        if(!strcmp(stage,"SCAN_TEXT"))taskTestBoundary(L"SCAN_PROGRESS");
        if(!strcmp(stage,"RESCAN_TEXT"))taskTestBoundary(L"RESCAN_PROGRESS");
    }if(!enabled||!stage||!total)return;const auto now=GetTickCount64();
        if(completed==total||now-lastUpdate>=150){std::printf("UI_STAGE_PROGRESS %s %zu %zu\n",stage,completed,total);lastUpdate=now;}}
};

