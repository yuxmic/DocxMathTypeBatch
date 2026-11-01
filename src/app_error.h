// 所有模块共用的失败断言，不引入 Word 或任务监管依赖。
#pragma once
#include <stdexcept>

inline void require(bool ok, const char* label) {
    if (!ok) throw std::runtime_error(label);
}