// 极速正文转换：调用方负责私有副本、任务互斥及最终发布。
#pragma once
#include <string>
#include <vector>
struct MathTypeApi;
struct ConversionTimings;
struct UltraOptions {double points{12};bool numbers{},progress{};std::wstring stopFile;};
struct UltraProblem {size_t bodyBlock{};std::wstring source,reason;bool rejected{};};
struct UltraResult {size_t targets{},succeeded{},failed{},remaining{},numbersAdded{},texTargets{},ommlTargets{},cacheHits{},ommlTexImports{},ommlMathmlImports{};std::vector<UltraProblem> problems;};
UltraResult convertUltra(MathTypeApi& api,const std::wstring& candidate,const UltraOptions& options,ConversionTimings& timings);
// 只解析 XML 的开发断言，不启动 Office 或 MathType。
void ultraXmlSelfTest();
