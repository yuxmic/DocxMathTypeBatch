// 每个任务只累计关键阶段耗时，结束时一次输出；不增加逐式日志或后台线程。
#pragma once
#include <array>
#include <chrono>
#include <cstddef>
#include <cstdio>

enum class ConversionMetric {SourceXml,OmmlSnapshot,XslCompile,XmlPrepare,XslTransform,MathmlAdapt,TexImport,MathmlImport,SourceLocations,ExperimentalRegionState,ExperimentalRecovery,Count};
struct ConversionTimings {
    struct Sample {std::chrono::nanoseconds elapsed{};size_t count{};};
    std::array<Sample,static_cast<size_t>(ConversionMetric::Count)> samples{};
    void report() const {
        constexpr const char* names[]{"SOURCE_XML","OMML_SNAPSHOT","XSL_COMPILE","XML_PREPARE","XSL_TRANSFORM","MATHML_ADAPT","TEX_IMPORT","MATHML_IMPORT","SOURCE_LOCATIONS","EXPERIMENTAL_REGION_STATE","EXPERIMENTAL_RECOVERY"};
        static_assert(sizeof(names)/sizeof(*names)==static_cast<size_t>(ConversionMetric::Count),"Every timing stage has a name");
        for(size_t i=0;i<samples.size();++i)if(samples[i].count)
            std::printf("TIMING_DETAIL %s COUNT=%zu MS=%.3f\n",names[i],samples[i].count,std::chrono::duration<double,std::milli>(samples[i].elapsed).count());
    }
};
struct ScopedConversionTiming {
    ConversionTimings::Sample& sample;
    std::chrono::steady_clock::time_point started{std::chrono::steady_clock::now()};
    ScopedConversionTiming(ConversionTimings& timings,ConversionMetric metric):sample(timings.samples[static_cast<size_t>(metric)]){}
    ~ScopedConversionTiming(){sample.elapsed+=std::chrono::duration_cast<std::chrono::nanoseconds>(std::chrono::steady_clock::now()-started);++sample.count;}
    ScopedConversionTiming(const ScopedConversionTiming&)=delete;
    ScopedConversionTiming& operator=(const ScopedConversionTiming&)=delete;
};
