// 正文、额外位置和极速包转换共用的厂商公式导入。
#pragma once
#include "tex_mathml.h"
#include "ultra_assets.h"

inline std::wstring mathmlFromTex(MathTypeApi& api,const std::wstring& input,bool display,
                                 const std::vector<TexMathMLFix>& fixes,ConversionTimings& timings) {
    // 厂商导入会追加内容，标记只能写入独立的临时对象，不能写进成品。
    EquationHost scratch(timings);
    require(input.size()<=LONG_MAX&&api.write(scratch.object.Get(),1,input.c_str(),static_cast<long>(input.size()))==0,
            "UNSUPPORTED_TEX_MATHML_COMPATIBILITY");
    const auto exported=exportNativeTex(api,readNative(api,scratch.object.Get()),"MathML2 (namespace attr).tdl");
    require(std::all_of(exported.begin(),exported.end(),[](unsigned char c){return c<128;}),"UNSUPPORTED_TEX_MATHML_COMPATIBILITY");
    return compatibleTexMathML(std::wstring(exported.begin(),exported.end()),fixes,display);
}
