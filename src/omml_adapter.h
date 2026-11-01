// 正文和额外区域共用 OMML 导入适配；不处理编号或段落布局。
#pragma once
#include "word_automation.h"
#include <msxml6.h>

struct ConversionTimings;
std::wstring ommlSource(IDispatch* range);
// 只抽取明确主故事并核对 Word 实际数量；调用方按实际 OMaths 顺序核对完整 XML 锚点。
bool ommlSourcesFromBodyXml(const std::wstring& package,size_t wordCount,std::vector<std::wstring>& sources);
ComPtr<IXSLTemplate> compileOmmlStylesheet(const std::wstring& path,ConversionTimings& timings);
std::wstring finishMathML(std::wstring fragment,bool display,ConversionTimings& timings);
std::wstring ommlMathML(const std::wstring& sourcePackage,IXSLTemplate* stylesheet,bool display,ConversionTimings& timings);
