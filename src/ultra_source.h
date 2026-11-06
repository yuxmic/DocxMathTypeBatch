// 极速正文扫描与定点替换；依赖 ultra_conversion.cpp 中的 DOM 小工具。
#pragma once

struct UltraTextSlice {ComPtr<IXMLDOMNode> text;size_t offset{},length{};};
struct Candidate {
    ComPtr<IXMLDOMNode> paragraph,container,math;
    size_t bodyBlock{},textStart{},textLength{};
    std::wstring alignment{L"center"},source;
    bool display{},tex{},manualNumber{};
    std::vector<UltraTextSlice> slices;
};
bool ultraBlank(const std::wstring& value){return std::all_of(value.begin(),value.end(),[](wchar_t c){return iswspace(c)||c==0xa0||c==0x3000;});}
ComPtr<IXMLDOMNode> ultraClone(IXMLDOMNode* node,bool deep=true){ComPtr<IXMLDOMNode> copy;check(node->cloneNode(deep?VARIANT_TRUE:VARIANT_FALSE,&copy),"Clone ultra source node");return copy;}
ComPtr<IXMLDOMNode> ultraParent(IXMLDOMNode* node){ComPtr<IXMLDOMNode> parent;check(node->get_parentNode(&parent),"Ultra source parent");return parent;}
bool ultraSame(IXMLDOMNode* left,IXMLDOMNode* right){ComPtr<IUnknown> a,b;check(left->QueryInterface(IID_PPV_ARGS(&a)),"Ultra node identity");check(right->QueryInterface(IID_PPV_ARGS(&b)),"Ultra node identity");return a.Get()==b.Get();}
void ultraAfter(IXMLDOMNode* previous,IXMLDOMNode* node){auto parent=ultraParent(previous);ComPtr<IXMLDOMNode> next;check(previous->get_nextSibling(&next),"Next ultra source node");if(next)check(parent->insertBefore(node,Value(next.Get()).v,nullptr),"Insert ultra source sibling");else append(parent.Get(),node);}
void ultraText(IXMLDOMNode* node,const std::wstring& value){
    check(node->put_text(Value(value.c_str()).v.bstrVal),"Replace ultra text fragment");ComPtr<IXMLDOMElement> element;check(node->QueryInterface(IID_PPV_ARGS(&element)),"Ultra text element");
    // 原来的段中空格在切开 run 后可能成为首尾空格，必须显式保留。
    check(element->setAttribute(Value(L"xml:space").v.bstrVal,Value(L"preserve").v),"Preserve surrounding spaces");
}

struct UltraTextPiece {ComPtr<IXMLDOMNode> node;size_t start{},length{};};
void ultraTexCandidates(IXMLDOMNode* paragraph,size_t block,const std::wstring& text,const std::vector<UltraTextPiece>& pieces,size_t paragraphOffset,std::vector<Candidate>& selected){
    struct Marker{size_t at{},length{};int kind{};};
    Marker opening;bool haveOpening{};size_t pieceIndex{};
    // 一次线性扫描识别定界符，转义符及混用的定界符均不跨过去拼接。
    for(size_t at=0;at<text.size();){
        Marker token{at,0,0};
        if(text[at]==L'\\'&&at+1<text.size()){
            const wchar_t next=text[at+1];token.length=2;
            if(next==L'[')token.kind=3;else if(next==L']')token.kind=-3;else if(next==L'(')token.kind=4;else if(next==L')')token.kind=-4;
            if(!token.kind){at+=2;continue;}
        }else if(text[at]==L'$'){
            size_t end=at+1;while(end<text.size()&&text[end]==L'$')++end;
            token.length=end-at;token.kind=token.length==1?1:token.length==2?2:5;
        }else{++at;continue;}
        at+=token.length;
        if(token.kind==5){haveOpening=false;continue;}
        if(!haveOpening){if(token.kind>0){opening=token;haveOpening=true;}continue;}
        const bool closes=token.kind==(opening.kind<=2?opening.kind:-opening.kind);
        if(closes&&token.at>opening.at+opening.length){
            Candidate item;item.paragraph=paragraph;item.bodyBlock=block;item.tex=true;item.display=opening.kind==2||opening.kind==3;
            item.source=text.substr(opening.at+opening.length,token.at-opening.at-opening.length);
            const size_t end=token.at+token.length;
            item.textStart=paragraphOffset+opening.at;item.textLength=end-opening.at;
            while(pieceIndex<pieces.size()&&pieces[pieceIndex].start+pieces[pieceIndex].length<=opening.at)++pieceIndex;
            for(size_t index=pieceIndex;index<pieces.size();++index){const auto& piece=pieces[index];
                if(piece.start>=end)break;
                const size_t begin=std::max(opening.at,piece.start),finish=std::min(end,piece.start+piece.length);
                if(begin<finish)item.slices.push_back({piece.node,begin-piece.start,finish-begin});
            }
            if(!ultraBlank(item.source))selected.push_back(std::move(item));
        }
        haveOpening=false;
    }
}

std::vector<Candidate> candidates(IXMLDOMDocument2* document,UltraResult& result,bool=false){
    std::vector<Candidate> selected;size_t block{};long fieldDepth{};
    for(auto& paragraph:nodes(document,L"/w:document/w:body/*")){
        ++block;if(!select(paragraph.Get(),L"self::w:p"))continue;
        const auto first=selected.size();std::wstring text,paragraphText;std::vector<UltraTextPiece> pieces;bool ordinary=true,hasTabs=false;
        auto flush=[&](){ultraTexCandidates(paragraph.Get(),block,text,pieces,paragraphText.size()-text.size(),selected);text.clear();pieces.clear();};
        auto addMath=[&](IXMLDOMNode* math,IXMLDOMNode* container,bool display){
            Candidate item;item.paragraph=paragraph;item.math=math;item.container=container;item.bodyBlock=block;item.display=display;item.source=nodeText(math);item.textStart=paragraphText.size();
            auto alignment=select(container,L"m:oMathParaPr/m:jc/@m:val");if(alignment){auto value=nodeText(alignment.Get());if(value==L"left"||value==L"right")item.alignment=std::move(value);}
            selected.push_back(std::move(item));++result.remaining;
        };
        for(auto& child:nodes(paragraph.Get(),L"*")){
            if(select(child.Get(),L"self::w:pPr | self::w:bookmarkStart | self::w:bookmarkEnd | self::w:proofErr"))continue;
            if(select(child.Get(),L"self::w:r")){
                for(auto& entry:nodes(child.Get(),L"*")){
                    if(select(entry.Get(),L"self::w:rPr"))continue;
                    if(select(entry.Get(),L"self::w:t")&&fieldDepth==0){auto value=nodeText(entry.Get());pieces.push_back({entry,text.size(),value.size()});text+=value;paragraphText+=value;continue;}
                    // Tab 仍是 LaTeX 拼接边界，但独立公式外侧可作为编号分隔空白。
                    if(select(entry.Get(),L"self::w:tab")&&fieldDepth==0){flush();paragraphText+=L'\t';hasTabs=true;continue;}
                    flush();ordinary=false;
                    if(select(entry.Get(),L"self::w:fldChar")){const auto kind=attribute(entry.Get(),L"w:fldCharType");if(kind==L"begin")++fieldDepth;else if(kind==L"end"&&fieldDepth)--fieldDepth;}
                }
                continue;
            }
            flush();
            if(!fieldDepth&&select(child.Get(),L"self::m:oMath")){addMath(child.Get(),child.Get(),false);continue;}
            if(!fieldDepth&&select(child.Get(),L"self::m:oMathPara")){for(auto& math:nodes(child.Get(),L"m:oMath"))addMath(math.Get(),child.Get(),true);continue;}
            // 域、内容控件、修订和额外容器是扫描边界，不读取其内部文本。
            ordinary=false;
        }
        flush();
        if(ordinary&&selected.size()==first+1){auto& item=selected.back();
            const bool leadingOnly=ultraBlank(paragraphText.substr(0,item.textStart));
            const auto trailing=paragraphText.substr(item.textStart+item.textLength);
            item.manualNumber=leadingOnly&&manualEquationNumber(trailing);
            if(leadingOnly&&(item.manualNumber||(!hasTabs&&ultraBlank(trailing))))item.display=true;
        }
    }
    return selected;
}

bool ultraVisible(IXMLDOMNode* node){
    if(select(node,L"self::w:pPr | self::w:bookmarkStart | self::w:bookmarkEnd | self::w:proofErr"))return false;
    if(select(node,L"self::w:r"))return !ultraBlank(nodeText(node))||select(node,L"*[not(self::w:rPr or self::w:t)]")!=nullptr;
    return true;
}
ComPtr<IXMLDOMNode> ultraNewParagraph(IXMLDOMNode* paragraph){
    auto copy=ultraClone(paragraph,false);ComPtr<IXMLDOMElement> element;check(copy.As(&element),"Split paragraph element");
    for(auto& id:nodes(copy.Get(),L"@*[local-name()='paraId' or local-name()='textId']")){Value name;name.v.vt=VT_BSTR;check(id->get_nodeName(&name.v.bstrVal),"Paragraph identity name");check(element->removeAttribute(name.v.bstrVal),"Do not clone unique paragraph identity");}
    auto props=select(paragraph,L"w:pPr");if(props){props=ultraClone(props.Get());
        for(auto& extra:nodes(props.Get(),L"w:sectPr | w:pageBreakBefore | w:numPr"))check(props->removeChild(extra.Get(),nullptr),"Keep section, page break and list number once");append(copy.Get(),props.Get());}
    return copy;
}
ComPtr<IXMLDOMNode> ultraDisplayParagraph(IXMLDOMNode* paragraph,IXMLDOMNode* run){
    auto children=nodes(paragraph,L"*");size_t at{};while(at<children.size()&&!ultraSame(children[at].Get(),run))++at;require(at<children.size(),"Ultra equation belongs to paragraph");
    bool prefix{},suffix{};for(size_t i=0;i<at;++i)prefix=prefix||ultraVisible(children[i].Get());for(size_t i=at+1;i<children.size();++i)suffix=suffix||ultraVisible(children[i].Get());
    ComPtr<IXMLDOMNode> formula=paragraph,tail;
    if(suffix){tail=ultraNewParagraph(paragraph);for(size_t i=at+1;i<children.size();++i)append(tail.Get(),children[i].Get());ultraAfter(paragraph,tail.Get());}
    if(prefix){formula=ultraNewParagraph(paragraph);append(formula.Get(),run);ultraAfter(paragraph,formula.Get());}
    // 分段只保留一个节属性，并放在最后一段，避免意外产生新节。
    auto last=tail?tail:formula;auto section=select(paragraph,L"w:pPr/w:sectPr");if(section&&!ultraSame(last.Get(),paragraph)){auto props=select(last.Get(),L"w:pPr");append(props.Get(),section.Get());}
    return formula;
}

// 必须按候选的逆序调用：先处理后面的文本，前面各片段的字符偏移才始终有效。
ComPtr<IXMLDOMNode> replaceCandidate(IXMLDOMDocument2*,const Candidate& item,IXMLDOMNode* objectRun){
    ComPtr<IXMLDOMNode> paragraph;
    if(item.tex){
        const auto& first=item.slices.front();auto run=ultraParent(first.text.Get());paragraph=ultraParent(run.Get());
        for(const auto& slice:item.slices){auto value=nodeText(slice.text.Get());value.erase(slice.offset,slice.length);ultraText(slice.text.Get(),value);}
        auto tail=ultraClone(run.Get(),false);auto props=select(run.Get(),L"w:rPr");if(props)append(tail.Get(),ultraClone(props.Get()).Get());
        auto value=nodeText(first.text.Get());auto remainder=ultraClone(first.text.Get());ultraText(remainder.Get(),value.substr(first.offset));
        ultraText(first.text.Get(),value.substr(0,first.offset));append(tail.Get(),remainder.Get());
        ComPtr<IXMLDOMNode> child;check(first.text->get_nextSibling(&child),"Following TeX run content");
        while(child){ComPtr<IXMLDOMNode> next;check(child->get_nextSibling(&next),"Next TeX run content");append(tail.Get(),child.Get());child=std::move(next);}
        ultraAfter(run.Get(),objectRun);ultraAfter(objectRun,tail.Get());
    }else{
        auto parent=ultraParent(item.math.Get());
        if(select(parent.Get(),L"self::m:oMathPara")){
            paragraph=ultraParent(parent.Get());auto tail=ultraClone(parent.Get(),false);auto props=select(parent.Get(),L"m:oMathParaPr");if(props)append(tail.Get(),ultraClone(props.Get()).Get());
            ComPtr<IXMLDOMNode> child;check(item.math->get_nextSibling(&child),"Following display OMML");
            while(child){ComPtr<IXMLDOMNode> next;check(child->get_nextSibling(&next),"Next display OMML");append(tail.Get(),child.Get());child=std::move(next);}
            check(parent->removeChild(item.math.Get(),nullptr),"Remove converted display OMML");ultraAfter(parent.Get(),objectRun);
            if(select(tail.Get(),L"m:oMath"))ultraAfter(objectRun,tail.Get());
            if(!select(parent.Get(),L"m:oMath"))check(paragraph->removeChild(parent.Get(),nullptr),"Remove empty display wrapper");
        }else{paragraph=parent;check(paragraph->replaceChild(objectRun,item.math.Get(),nullptr),"Replace inline OMML");}
    }
    // 仅成功替换、整段只有一式和手动尾号时清理；失败候选完全不改原节点。
    if(item.manualNumber)for(auto& node:nodes(paragraph.Get(),L"w:r/w:t | w:r/w:tab"))check(ultraParent(node.Get())->removeChild(node.Get(),nullptr),"Remove manual equation number and whitespace");
    return item.display?ultraDisplayParagraph(paragraph.Get(),objectRun):paragraph;
}

void ultraSourceSelfTest(){
    const std::wstring math=L"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>";
    auto doc=xmlDocument(L"<w:document "+std::wstring(namespaces)+L"><w:body>"
        L"<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space='preserve'>A $x</w:t><w:t>+1$ B \\(y</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space='preserve'>\\) C $$z$$ D</w:t></w:r></w:p>"
        L"<w:p><w:r><w:t>mixed </w:t></w:r>"+math+L"<w:r><w:t> and $q$ end</w:t></w:r></w:p>"
        L"<w:p><m:oMathPara>"+math+math+L"</m:oMathPara></w:p><w:p>"+math+L"</w:p>"
        L"<w:tbl><w:tr><w:tc><w:p>"+math+L"<w:r><w:t>$ignored$</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        L"<w:p><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:t>$field$</w:t></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r><w:r><w:t>$a$</w:t><w:tab/><w:t>$b$</w:t></w:r></w:p>"
        L"</w:body></w:document>");
    check(doc->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(namespaces).v),"Ultra source test namespaces");UltraResult result;auto list=candidates(doc.Get(),result);
    require(list.size()==10&&result.remaining==4&&list[0].source==L"x+1"&&list[1].source==L"y"&&!list[0].display&&list[2].display,"Scan split runs, mixed formulas and excluded regions");
    require(!list[3].display&&list[5].display&&list[6].display&&list[7].display,"Classify inline and display OMML");
    for(size_t i=list.size();i-->0;){auto run=importXml(doc.Get(),L"<w:r><w:object marker='"+std::to_wstring(i)+L"'/></w:r>");replaceCandidate(doc.Get(),list[i],run.Get());}
    require(nodes(doc.Get(),L"/w:document/w:body/w:p/w:r/w:object").size()==10&&nodes(doc.Get(),L"/w:document/w:body/w:p//m:oMath").empty(),"Replace every body formula independently");
    auto texts=nodes(doc.Get(),L"/w:document/w:body/w:p/w:r/w:t");std::wstring remaining;for(auto& text:texts)remaining+=nodeText(text.Get());
    require(remaining==L"A  B  C  Dmixed  and  end$field$","Preserve surrounding and field text");
    require(nodes(doc.Get(),L"//w:sectPr").size()==1&&select(doc.Get(),L"/w:document/w:body/w:p[3]/w:pPr/w:sectPr")!=nullptr,"Keep section only on final split paragraph");
    require(select(doc.Get(),L"/w:document/w:body/w:p[1]/w:r[w:t='A ']/w:rPr/w:b")&&select(doc.Get(),L"/w:document/w:body/w:p[3]/w:r[w:t=' D']/w:rPr/w:i"),"Preserve split run formats");
    require(nodes(doc.Get(),L"//w:tbl//m:oMath").size()==1&&nodeText(select(doc.Get(),L"//w:tbl//w:t").Get())==L"$ignored$","Leave excluded tables intact");
    // 同段只有一式才清理；扫描不写 DOM，未执行成功替换的候选保留旧编号。
    const std::vector<std::wstring> numbered={
        L"<w:r><w:t>  $x+1$     (27)  </w:t></w:r>",
        L"<w:r><w:tab/><w:t>\\(x+1\\)</w:t><w:tab/><w:t>（3</w:t></w:r><w:r><w:t>－2）　</w:t></w:r>",
        L"<w:r><w:t> $$x+1$$  (3.2) </w:t></w:r>",
        L"<w:r><w:t> \\[x+1\\]  (3-2) </w:t></w:r>",
        L"<w:r><w:tab/><w:t> </w:t></w:r>"+math+L"<w:r><w:t>  (7) </w:t><w:tab/></w:r>",
        L"<m:oMathPara>"+math+L"</m:oMathPara><w:r><w:tab/><w:t> （8） </w:t></w:r>"
    };
    for(const auto& content:numbered){
        auto sample=xmlDocument(L"<w:document "+std::wstring(namespaces)+L"><w:body><w:p><w:pPr><w:spacing w:after='120'/></w:pPr><w:bookmarkStart w:id='0' w:name='formula'/>"+content+L"<w:bookmarkEnd w:id='0'/></w:p></w:body></w:document>");
        check(sample->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(namespaces).v),"Manual number test namespaces");
        const auto before=nodeXml(sample.Get());UltraResult scanned;auto found=candidates(sample.Get(),scanned);
        require(found.size()==1&&found[0].display&&found[0].manualNumber&&nodeXml(sample.Get())==before,"Recognize manual number without mutating an unconverted formula");
        replaceCandidate(sample.Get(),found[0],importXml(sample.Get(),L"<w:r><w:object/></w:r>").Get());
        require(nodes(sample.Get(),L"//w:p").size()==1&&nodes(sample.Get(),L"//w:object").size()==1&&nodes(sample.Get(),L"//w:t | //w:tab | //m:oMath").empty(),"Clear manual number and spaces without creating blank paragraphs");
        require(nodes(sample.Get(),L"//w:bookmarkStart | //w:bookmarkEnd").size()==2&&select(sample.Get(),L"//w:spacing[@w:after='120']"),"Keep bookmarks and paragraph properties while clearing a number");
    }
    const std::vector<std::wstring> preserved={
        L"<w:r><w:t>正文 $x$ (27)</w:t></w:r>",
        L"<w:r><w:t>$x$ (m/s)</w:t></w:r>",
        L"<w:r><w:t>$x$ $y$ (27)</w:t></w:r>",
        L"<w:r><w:t>$x$</w:t><w:fldChar w:fldCharType='begin'/><w:fldChar w:fldCharType='end'/><w:t> (27)</w:t></w:r>",
        L"<w:r><w:t>$x$</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t> (27)</w:t></w:r></w:sdtContent></w:sdt>",
        L"<w:r><w:t>$x$</w:t><w:br/><w:t> (27)</w:t></w:r>"
    };
    for(const auto& content:preserved){
        auto sample=xmlDocument(L"<w:document "+std::wstring(namespaces)+L"><w:body><w:p>"+content+L"</w:p></w:body></w:document>");
        check(sample->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(namespaces).v),"Preserved number test namespaces");UltraResult scanned;auto found=candidates(sample.Get(),scanned);
        require(!found.empty()&&std::none_of(found.begin(),found.end(),[](const Candidate& item){return item.manualNumber;}),"Keep prose, units, multiple formulas and protected boundaries out of cleanup");
    }
    // 没有旧号时保留原先的 Tab 边界，避免升级行外后把 Tab 拆成空段落。
    for(const auto& content:std::vector<std::wstring>{L"<w:r><w:tab/><w:t>$x$</w:t><w:tab/></w:r>",L"<w:r><w:tab/></w:r>"+math+L"<w:r><w:tab/></w:r>"}){
        auto sample=xmlDocument(L"<w:document "+std::wstring(namespaces)+L"><w:body><w:p>"+content+L"</w:p></w:body></w:document>");
        check(sample->setProperty(Value(L"SelectionNamespaces").v.bstrVal,Value(namespaces).v),"Unnumbered Tab test namespaces");UltraResult scanned;auto found=candidates(sample.Get(),scanned);
        require(found.size()==1&&!found[0].display&&!found[0].manualNumber,"Unnumbered Tabs retain the original inline classification");
        replaceCandidate(sample.Get(),found[0],importXml(sample.Get(),L"<w:r><w:object/></w:r>").Get());
        require(nodes(sample.Get(),L"//w:p").size()==1&&nodes(sample.Get(),L"//w:tab").size()==2,"Do not split unnumbered Tabs into empty paragraphs");
    }
}
