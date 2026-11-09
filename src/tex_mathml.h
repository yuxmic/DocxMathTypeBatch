// 从厂商导出的 MathML 恢复兼容标记，保留 TeX 导入器已识别的公式主体。
#pragma once
#include "docx_package.h"
#include "latex_rules.h"

inline std::wstring compatibleTexMathML(const std::wstring& xml,const std::vector<TexMathMLFix>& fixes,bool display) {
    constexpr const wchar_t* ns=L"http://www.w3.org/1998/Math/MathML";
    auto document=xmlDocument(xml);
    ComPtr<IXMLDOMElement> root;
    check(document->get_documentElement(&root),"Read TeX compatibility MathML root");
    check(root->setAttribute(Value(L"display").v.bstrVal,Value(display?L"block":L"inline").v),"Set TeX compatibility display style");
    const auto select=[](IXMLDOMNode* node,const wchar_t* path){
        ComPtr<IXMLDOMNode> selected;
        check(node->selectSingleNode(Value(path).v.bstrVal,&selected),"Find TeX compatibility placeholder");
        require(static_cast<bool>(selected),"UNSUPPORTED_TEX_MATHML_COMPATIBILITY");
        return selected;
    };
    const auto element=[&](const wchar_t* name){
        ComPtr<IXMLDOMNode> result;
        check(document->createNode(Value(1L).v,Value(name).v.bstrVal,Value(ns).v.bstrVal,&result),"Create TeX compatibility MathML node");
        return result;
    };
    const auto space=[&](const std::wstring& width){
        auto node=element(L"mspace");
        ComPtr<IXMLDOMElement> attributes;
        check(node.As(&attributes),"Read TeX MathML spacing element");
        check(attributes->setAttribute(Value(L"width").v.bstrVal,Value(width.c_str()).v),"Set TeX MathML spacing width");
        return node;
    };
    for(const auto& fix:fixes){
        const auto path=L"//*[(local-name()='mtext' or local-name()='mi' or local-name()='mo') and contains(.,'"+fix.marker+L"')]";
        ComPtr<IXMLDOMNodeList> matches;
        check(document->selectNodes(Value(path.c_str()).v.bstrVal,&matches),"Count TeX compatibility placeholders");
        long count{};
        check(matches->get_length(&count),"Read TeX compatibility placeholder count");
        require(count==1,"UNSUPPORTED_TEX_MATHML_COMPATIBILITY");
        ComPtr<IXMLDOMNode> token;
        check(matches->get_item(0,&token),"Read TeX compatibility placeholder");
        if(fix.kind==TexMathMLFixKind::TextLetter){
            Value content;
            content.v.vt=VT_BSTR;
            check(token->get_text(&content.v.bstrVal),"Read TeX text token");
            std::wstring text(content.v.bstrVal);
            const auto start=text.find(fix.marker);
            require(start!=std::wstring::npos&&text.find(fix.marker,start+fix.marker.size())==std::wstring::npos,
                    "UNSUPPORTED_TEX_MATHML_COMPATIBILITY");
            text.replace(start,fix.marker.size(),fix.value);
            check(token->put_text(Value(text.c_str()).v.bstrVal),"Restore double acute text letter");
            continue;
        }

        auto mover=select(token.Get(),L"ancestor::*[local-name()='mover'][1]");
        auto upper=select(mover.Get(),(L"*[2][normalize-space(.)='"+fix.marker+L"']").c_str());
        if(fix.kind==TexMathMLFixKind::Accent){
            if(fix.value==L"\u23b4"||fix.value==L"\u21bc"){
                // 厂商只能导入固定宽度的方括号和左鱼叉，多原子参数保持原 TeX。
                ComPtr<IXMLDOMNode> atom;
                check(mover->selectSingleNode(Value(L"*[1][(local-name()='mi' or local-name()='mn' or local-name()='mo') and string-length(.)=1]").v.bstrVal,&atom),
                      "Check fixed-width accent base");
                require(static_cast<bool>(atom),fix.value==L"\u23b4"?"UNSUPPORTED_TEX_BRACKET_WIDTH":"UNSUPPORTED_TEX_HARPOON_WIDTH");
            }
            auto mark=element(L"mo");
            check(mark->put_text(Value(fix.value.c_str()).v.bstrVal),"Set TeX accent glyph");
            check(mover->replaceChild(mark.Get(),upper.Get(),nullptr),"Restore TeX accent");
            continue;
        }

        ComPtr<IXMLDOMNode> replacement,parent;
        if(fix.kind==TexMathMLFixKind::MathLetter){
            auto base=select(mover.Get(),L"*[1]");
            auto letter=select(base.Get(),L"descendant-or-self::*[local-name()='mi' and .='x']");
            check(letter->put_text(Value(fix.value.c_str()).v.bstrVal),"Restore double acute math letter");
            check(base->cloneNode(VARIANT_TRUE,&replacement),"Retain double acute math letter style");
        }else replacement=space(fix.value);
        check(mover->get_parentNode(&parent),"Read TeX compatibility parent");
        check(parent->replaceChild(replacement.Get(),mover.Get(),nullptr),"Restore TeX MathML token");
    }
    if(std::any_of(fixes.begin(),fixes.end(),[](const TexMathMLFix& fix){return fix.kind==TexMathMLFixKind::TextLetter;})){
        // 带双锐音的文字经 MathML 写回时，裸 NBSP 可能在 MathType 预览中画成问号。
        ComPtr<IXMLDOMNodeList> textRuns;
        check(document->selectNodes(Value(L"//*[local-name()='mtext' and contains(.,'\u00a0')]").v.bstrVal,&textRuns),"Find TeX compatibility text spaces");
        long count{};check(textRuns->get_length(&count),"Count TeX compatibility text spaces");
        for(long i=0;i<count;++i){
            ComPtr<IXMLDOMNode> text,parent;check(textRuns->get_item(i,&text),"Read TeX compatibility text");
            Value content;content.v.vt=VT_BSTR;check(text->get_text(&content.v.bstrVal),"Read TeX compatibility text value");
            const std::wstring value=content.v.bstrVal;
            auto row=element(L"mrow");
            for(size_t start=0;start<value.size();){
                const auto end=std::min(value.find(L'\u00a0',start),value.size());
                if(end>start){
                    ComPtr<IXMLDOMNode> part;check(text->cloneNode(VARIANT_FALSE,&part),"Retain TeX text style");
                    check(part->put_text(Value(value.substr(start,end-start).c_str()).v.bstrVal),"Write TeX text segment");
                    check(row->appendChild(part.Get(),nullptr),"Append TeX text segment");
                }
                if(end<value.size())check(row->appendChild(space(L"0.17em").Get(),nullptr),"Append TeX text space");
                start=end+1;
            }
            check(text->get_parentNode(&parent),"Read TeX text parent");
            check(parent->replaceChild(row.Get(),text.Get(),nullptr),"Replace NBSP with explicit TeX text spaces");
        }
    }
    // 独立鱼叉字形能导入，厂商左鱼叉伸缩模板返回错误。零宽空白保留原有居中字形。
    ComPtr<IXMLDOMNodeList> leftHarpoons;
    check(document->selectNodes(Value(L"//*[local-name()='mover']/*[2][local-name()='mo' and .='\u21bc']").v.bstrVal,&leftHarpoons),
          "Find left harpoon accents in TeX compatibility MathML");
    long count{};
    check(leftHarpoons->get_length(&count),"Count left harpoon accents");
    for(long index=count;index>0;--index){
        ComPtr<IXMLDOMNode> harpoon,parent;
        check(leftHarpoons->get_item(index-1,&harpoon),"Read left harpoon glyph");
        check(harpoon->get_parentNode(&parent),"Read left harpoon parent");
        auto row=element(L"mrow"),zero=space(L"0em");
        check(parent->replaceChild(row.Get(),harpoon.Get(),nullptr),"Place left harpoon in a slot");
        check(row->appendChild(harpoon.Get(),nullptr),"Retain left harpoon glyph");
        check(row->appendChild(zero.Get(),nullptr),"Keep the left harpoon slot explicit");
    }
    Value output;
    output.v.vt=VT_BSTR;
    check(document->get_xml(&output.v.bstrVal),"Serialize TeX compatibility MathML");
    return output.v.bstrVal;
}
