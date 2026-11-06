// 三档转换模式在界面、队列和工作进程状态中共用，避免把极速误记为快速。
#pragma once
enum class ConversionMode { Standard, Fast, Ultra };
inline const wchar_t* conversionModeLabel(ConversionMode mode) {
    return mode==ConversionMode::Ultra?L"极速":mode==ConversionMode::Fast?L"快速":L"标准";
}
inline bool unverifiedMode(ConversionMode mode) { return mode!=ConversionMode::Standard; }
