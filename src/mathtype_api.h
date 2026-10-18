// 原生小样与批处理实际共用的厂商 API、偏好与只读 MTEF 样式检查。
#pragma once
#include "word_automation.h"

inline std::wstring mathTypeInstallProblem(const std::wstring& directory) {
    for (const auto* component : {L"MathType.exe", L"MathPage\\64\\MathPage.wll", L"Office Support\\BlankEqn.doc",
                                  L"Preferences\\Times+Symbol 12.eqp", L"Translators\\AMS LaTeX.tdl"}) {
        const auto attributes = GetFileAttributesW((directory + L"\\" + component).c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY))
            return L"MathType 缺少组件：" + std::wstring(component) + L"。请修复 MathType 安装。";
    }
    return {};
}

// 只读实际 OLE 服务的注册位置，避免接口与 Word 创建的对象来自不同安装。
inline std::wstring findMathTypeInstall(std::wstring& problem) {
    problem = L"未检测到 MathType：未安装或 Equation.DSMT4 公式组件未注册，请安装或修复 MathType。";
    auto registryText = [](const std::wstring& key, const wchar_t* name = nullptr) {
        std::wstring value(32768, L'\0'); DWORD bytes = static_cast<DWORD>(value.size() * sizeof(wchar_t));
        if (RegGetValueW(HKEY_CLASSES_ROOT, key.c_str(), name, RRF_RT_REG_SZ | RRF_RT_REG_EXPAND_SZ,
                         nullptr, value.data(), &bytes) != ERROR_SUCCESS) return std::wstring();
        value.resize(wcslen(value.c_str())); return value;
    };
    const auto registered = registryText(L"Equation.DSMT4\\CLSID"); CLSID cls{};
    if (registered.empty()) return {};
    problem = L"MathType 注册信息无效，请修复 MathType 安装。";
    if (FAILED(CLSIDFromString(registered.c_str(), &cls))) return {};
    const auto key = L"CLSID\\" + registered + L"\\LocalServer32";
    auto program = registryText(key, L"ServerExecutable");
    if (program.empty()) {
        const auto command = registryText(key); const auto start = command.find_first_not_of(L" \t");
        if (start == std::wstring::npos) return {};
        if (command[start] == L'"') {
            const auto end = command.find(L'"', start + 1); if (end == std::wstring::npos) return {};
            program = command.substr(start + 1, end - start - 1);
        } else {
            auto lower = command; std::transform(lower.begin(), lower.end(), lower.begin(), towlower);
            const auto executable = lower.find(L"\\mathtype.exe", start); if (executable == std::wstring::npos) return {};
            const auto end = executable + std::wstring(L"\\mathtype.exe").size();
            if (end != command.size() && !iswspace(command[end])) return {};
            program = command.substr(start, end - start);
        }
    }
    const auto slash = program.find_last_of(L"\\/");
    if (slash == std::wstring::npos || _wcsicmp(program.c_str() + slash + 1, L"MathType.exe") ||
        !((program.size() > 2 && program[1] == L':' && (program[2] == L'\\' || program[2] == L'/')) || program.rfind(L"\\\\", 0) == 0)) return {};
    const auto directory = program.substr(0, slash);
    problem = mathTypeInstallProblem(directory); return problem.empty() ? directory : std::wstring();
}

// 按厂商 VBA ABI 定义：Integer=16 位，Long=32 位，ByVal Object=IDispatch*。
struct MathTypeApi {
    HMODULE library{};
    const std::wstring install;
    using Init = long (WINAPI*)(short, short);
    using Term = long (WINAPI*)();
    using Write = long (WINAPI*)(IDispatch*, short, const wchar_t*, long);
    using WriteBytes = long (WINAPI*)(IDispatch*, short, const unsigned char*, long);
    using Read = long (WINAPI*)(IDispatch*, short, char*, long*);
    using ReadBytes = long (WINAPI*)(IDispatch*, short, unsigned char*, long*);
    using Close = long (WINAPI*)(long, IDispatch*);
    using Prefs = long (WINAPI*)(const char*, char*, short);
    using SetPrefs = long (WINAPI*)(short, const char*, short);
    using DefaultPrefs = long (WINAPI*)(char*, short);
    Init init{}; Term term{}; Write write{}; Read read{}; Close close{};
    WriteBytes writeBytes{};
    Prefs prefs{}; SetPrefs setPrefs{};
    ReadBytes readBytes{};
    DefaultPrefs defaultPrefs{};
    bool initialized{};

    explicit MathTypeApi(const std::wstring& directory) : install(directory) {
        library = LoadLibraryExW((install + L"\\MathPage\\64\\MathPage.wll").c_str(), nullptr,
                                LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
        require(library != nullptr, "Load x64 MathPage.wll");
        init = reinterpret_cast<Init>(GetProcAddress(library, "MTInitAPI"));
        term = reinterpret_cast<Term>(GetProcAddress(library, "MTTermAPI"));
        write = reinterpret_cast<Write>(GetProcAddress(library, "MTSetEqnFromLangStr"));
        writeBytes = reinterpret_cast<WriteBytes>(GetProcAddress(library,"MTSetEqnFromLangBytes"));
        read = reinterpret_cast<Read>(GetProcAddress(library, "MTGetLangStrFromEqn"));
        close = reinterpret_cast<Close>(GetProcAddress(library, "MTCloseOleObject"));
        prefs = reinterpret_cast<Prefs>(GetProcAddress(library, "MTGetPrefsFromFile"));
        setPrefs = reinterpret_cast<SetPrefs>(GetProcAddress(library, "MTSetMTPrefs"));
        readBytes = reinterpret_cast<ReadBytes>(GetProcAddress(library, "MTGetLangBytesFromEqn"));
        defaultPrefs = reinterpret_cast<DefaultPrefs>(GetProcAddress(library, "MTGetPrefsMTDefault"));
        if (!init || !term || !write || !writeBytes || !read || !close || !prefs || !setPrefs || !readBytes || !defaultPrefs) {
            FreeLibrary(library); library = nullptr;
            throw std::runtime_error("Missing MathType export");
        }
    }
    ~MathTypeApi() {
        if (initialized) std::printf("MTTermAPI=%ld\n", term());
        if (library) FreeLibrary(library);
    }
};

inline std::string readDefaultPrefs(MathTypeApi& api) {
    const long length = api.defaultPrefs(nullptr, 0);
    std::printf("DEFAULT_PREFERENCES_LENGTH=%ld\n", length);
    require(length > 0 && length < 32767, "Query default preference length");
    std::vector<char> buffer(static_cast<size_t>(length) + 1, '\0');
    require(api.defaultPrefs(buffer.data(), static_cast<short>(length)) == 0, "Read default preferences");
    return std::string(buffer.data());
}

// 读取任务 EQP 的偏好串；厂商接口使用 ANSI 路径，先转成本机短路径避免中文丢失。
inline std::string loadPrefs(MathTypeApi& api, const std::wstring& path) {
    std::vector<wchar_t> shortPath(32768);
    const DWORD count = GetShortPathNameW(path.c_str(), shortPath.data(), static_cast<DWORD>(shortPath.size()));
    require(count > 0 && count < shortPath.size(), "EQP short path");
    BOOL replaced = FALSE;
    const int bytes = WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, shortPath.data(), -1, nullptr, 0, nullptr, &replaced);
    require(bytes > 0 && !replaced, "EQP ANSI path");
    std::vector<char> ansi(static_cast<size_t>(bytes));
    require(WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, shortPath.data(), -1, ansi.data(), bytes, nullptr, &replaced) > 0 && !replaced,
            "EQP ANSI conversion");
    const long length = api.prefs(ansi.data(), "", 0);
    std::printf("PREFERENCE_QUERY_LENGTH=%ld\n",length);
    require(length > 0 && length < 32767, "Query preference length");
    std::vector<char> buffer(static_cast<size_t>(length) + 1, '\0');
    require(api.prefs(ansi.data(), buffer.data(), static_cast<short>(length)) == 0, "Read preferences");
    return std::string(buffer.data());
}

// lang=1 有 TeX Input 缓存时优先回读，无缓存时也可生成 TeX；不能据此保证实际结构完整。
inline std::string readTex(MathTypeApi& api, IDispatch* object) {
    long length{};
    const long query = api.read(object, 1, nullptr, &length);
    std::printf("MTGetLangStrFromEqn query=%ld length=%ld\n", query, length);
    require(query == 0 && length > 0 && length < 1000000, "Query TeX length");
    const long capacity = length;
    std::vector<char> buffer(static_cast<size_t>(capacity) + 1, '\0');
    const long result = api.read(object, 1, buffer.data(), &length);
    require(result == 0 && length >= 0 && length <= capacity, "Read TeX length");
    std::printf("MTGetLangStrFromEqn=%ld TEX=%s\n", result, buffer.data());
    return std::string(buffer.data());
}

// 原生 MTEF 字节只读回查，绝不自行生成、改写或回送 MTEF。
inline std::vector<unsigned char> readNative(MathTypeApi& api, IDispatch* object) {
    long length{};
    const long queried = api.readBytes(object, 3, nullptr, &length);
    std::printf("MTEF_QUERY=%ld LENGTH=%ld\n", queried, length);
    require(queried == 0 && length > 0 && length < 1000000, "Query native bytes");
    const long capacity = length;
    std::vector<unsigned char> buffer(static_cast<size_t>(capacity));
    require(api.readBytes(object, 3, buffer.data(), &length) == 0 && length > 0 && length <= capacity,
            "Read native bytes");
    buffer.resize(static_cast<size_t>(length));
    return buffer;
}

// 厂商翻译器读取实际 MTEF；不读 TeX Input 缓存、不用剪贴板、不回写对象。
// ABI/常量以本机 Declarations64.bas:114–119、Declarations.bas:583–605 为准。
inline std::string exportNativeTex(MathTypeApi& api, const std::vector<unsigned char>& native,
                                   const char* translator = "AMS LaTeX.tdl") {
    using Reset = long (WINAPI*)();
    using Translator = long (WINAPI*)(short, const char*);
    using Transform = long (WINAPI*)(short, short, const unsigned char*, long,
                                   short, short, char*, long, const char*, void*);
    using Status = long (WINAPI*)(short);
    const auto reset = reinterpret_cast<Reset>(GetProcAddress(api.library, "MTXFormReset"));
    const auto setTranslator = reinterpret_cast<Translator>(GetProcAddress(api.library, "MTXFormSetTranslator"));
    const auto transform = reinterpret_cast<Transform>(GetProcAddress(api.library, "MTXFormEqn"));
    const auto status = reinterpret_cast<Status>(GetProcAddress(api.library, "MTXFormGetStatus"));
    require(reset && setTranslator && transform && status, "Missing MathType translation exports");
    require(!native.empty() && native.size() < 1000000, "Native translation input bounds");
    require(reset() == 0, "Reset equation translation options");
    try {
        // INC_NONE=0：不附加翻译器名或私有 MTEF 注释；仅使用本机厂商翻译器。
        require(setTranslator(0, translator) == 0, "Select native TeX translator");
        // LOCAL=-3，MTEF=4，TEXT=7；先以小缓冲获得实际长度，再重读完整结果。
        char first[1]{};
        // 沿用厂商宏的有效 dims 和空路径字符串，不猜测本机 WLL 的可选指针行为。
        struct Dimensions { short baseline; RECT bounds; } dimensions{};
        const auto queried = transform(-3, 4, native.data(), static_cast<long>(native.size()),
                                       -3, 7, first, 1, "", &dimensions);
        const auto length = status(-1);
        require((queried == -9 || queried == 0) && length > 0 && length < 1000000,
                "Query translated TeX length");
        std::vector<char> buffer(static_cast<size_t>(length) + 1, '\0');
        require(transform(-3, 4, native.data(), static_cast<long>(native.size()), -3, 7,
                          buffer.data(), static_cast<long>(buffer.size()), "", &dimensions) == 0,
                "Translate actual native equation");
        const auto actual = status(-1);
        std::printf("TRANSLATION_LENGTH QUERY=%ld REQUIRED=%ld ACTUAL=%ld TEX=%s\n",
                    queried, length, actual, buffer.data());
        require(actual > 0 && actual <= length, "Native translator result length");
        require(buffer.at(static_cast<size_t>(actual) - 1) == '\0', "Terminated TeX translation");
        const std::string result(buffer.data());
        require(!result.empty(), "Nonempty native TeX translation");
        require(reset() == 0, "Restore equation translation options");
        return result;
    } catch (...) {
        reset();
        throw;
    }
}

struct EquationStyle {
    bool inlineEquation{};
    std::vector<std::pair<unsigned int, std::string>> sizes;
    std::vector<std::pair<std::string, unsigned int>> styles;
};

// ponytail: 只检查厂商 MTEF v5 的头部偏好区；完整数学结构审计时再增加正文解析。
// 依据：https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later
inline EquationStyle inspectStyle(const std::vector<unsigned char>& data) {
    size_t position{};
    auto byte = [&]() { return static_cast<unsigned int>(data.at(position++)); };
    auto unsignedInteger = [&]() {
        const auto first = byte();
        if (first != 255) return first;
        const auto low = byte(); return low + (byte() << 8);
    };
    auto string = [&]() {
        std::string value;
        for (auto next = byte(); next != 0; next = byte()) value += static_cast<char>(next);
        return value;
    };
    require(byte() == 5, "MTEF v5 header");
    position = 5;
    string();
    EquationStyle result;
    const auto options = byte();
    require((options & ~1u) == 0, "MTEF equation options");
    result.inlineEquation = (options & 1u) != 0;
    std::vector<std::string> fonts(1);
    auto dimensions = [&]() {
        std::vector<std::pair<unsigned int, std::string>> values;
        const auto count = byte();
        size_t nibble = position * 2;
        auto nextNibble = [&]() {
            const auto value = data.at(nibble / 2);
            return static_cast<unsigned int>((nibble++ % 2 == 0) ? value >> 4 : value & 15);
        };
        for (unsigned int index = 0; index < count; ++index) {
            const auto unit = nextNibble();
            require(unit <= 4, "Dimension unit");
            std::string value;
            for (auto digit = nextNibble(); digit != 15; digit = nextNibble()) {
                require(digit <= 11, "Dimension digit");
                value += digit < 10 ? static_cast<char>('0' + digit) : digit == 10 ? '.' : '-';
            }
            require(!value.empty(), "Empty dimension");
            values.emplace_back(unit, value);
        }
        position = (nibble + 1) / 2;
        return values;
    };
    while (position < data.size()) {
        const auto record = byte();
        if (record == 19) string();
        else if (record == 17) { unsignedInteger(); fonts.push_back(string()); }
        else if (record == 8) { unsignedInteger(); byte(); }
        else if (record == 15) unsignedInteger();
        else if (record == 16) {
            const auto colorOptions = byte();
            require((colorOptions & ~7u) == 0, "Color definition options");
            for (unsigned int index = 0; index < ((colorOptions & 1u) ? 8u : 6u); ++index) byte();
            if (colorOptions & 4u) string();
        }
        else if (record >= 100) {
            // 厂商定义的扩展记录带无符号长度，按长度跳过，不能在任意字节中搜索标签。
            const auto length = unsignedInteger();
            require(length <= data.size() - position, "Future record length");
            position += length;
        }
        else if (record == 18) {
            require(byte() == 0, "MTEF preferences options");
            result.sizes = dimensions();
            dimensions();
            const auto count = byte();
            for (unsigned int index = 0; index < count; ++index) {
                const auto font = unsignedInteger();
                const auto style = font == 0 ? 0 : byte();
                require(style <= 3, "Character style bits");
                result.styles.emplace_back(fonts.at(font), style);
            }
            return result;
        } else {
            std::printf("MTEF_UNSUPPORTED_RECORD=%u OFFSET=%zu\n", record, position - 1);
            throw std::runtime_error("Unsupported MTEF preference-prefix record");
        }
    }
    throw std::runtime_error("Missing MTEF preferences");
}

inline void verifyStyle(const EquationStyle& style, const std::string& font, const std::string& points, bool inlineEquation) {
    std::printf("NATIVE_STYLE FONT=%s FULL=%s INLINE=%d\n", style.styles.at(2).first.c_str(), style.sizes.at(0).second.c_str(), style.inlineEquation);
    require(style.inlineEquation == inlineEquation, "Internal inline/display type");
    require(style.sizes.at(0) == std::make_pair(2u, points), "Internal Full point size");
    require(style.sizes.at(1) == std::make_pair(4u, std::string("58")) &&
            style.sizes.at(2) == std::make_pair(4u, std::string("42")), "Subscript size ratios");
    for (size_t index : {0u, 1u, 2u, 6u, 7u, 9u}) require(style.styles.at(index).first == font, "Internal main style font");
    require(style.styles.at(2).second == 2 && style.styles.at(6).second == 1, "Variable italic/vector bold");
    require(style.styles.at(5).first == "Symbol", "Preserved Symbol font");
}
