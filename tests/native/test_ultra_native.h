// 极速开发核验只编入独立探针，不进入正式 GUI 转换路径。
#pragma once
static void verifyUltraSample(MathTypeApi& api,const std::wstring& input,const std::wstring& directory,bool mixed=false) {
    // 开发检查只作用于新副本：核对样本实际内容，改首式后重开检查其余对象独立性。
    const auto edited=directory+L"\\edited.docx";require(CopyFileW(input.c_str(),edited.c_str(),TRUE)!=0,"Create editable prototype check copy");
    WordSession word;auto documents=get(word.application.Get(),L"Documents");
    auto open=[&](){return asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{edited.c_str(),false,false,false,Value(),Value(),false,Value(),Value(),Value(),Value(),false}));};
    word.document=open();word.verifyOwnership();invoke(word.application.Get(),L"Visible",DISPATCH_PROPERTYPUT,{false});
    const auto before=bodyNativeAssets(word.document.Get());require(before.size()>=2,"Prototype sample has independent objects");
    if(!mixed)require(integer(get(get(word.document.Get(),L"Content").Get(),L"OMaths").Get(),L"Count")==0,"Sample has no unconverted OMML");
    std::string actual="[";
    for(size_t i=0;i<before.size();++i){auto tex=exportNativeTex(api,before[i]);
        const auto style=inspectStyle(before[i]);require(std::stod(style.sizes.at(0).second)==14.0,"Sample actual native size is 14 pt");
        actual+=(i?",":"")+std::string("{\"tex\":")+quote(std::wstring(tex.begin(),tex.end()))+",\"inline\":"+(style.inlineEquation?"true":"false")+"}";
        tex.erase(std::remove_if(tex.begin(),tex.end(),[](char c){return c=='{'||c=='}'||c==' '||c=='\r'||c=='\n'||c=='\t';}),tex.end());
        if(!mixed)require(tex=="\\[x_"+std::to_string(i+1)+"+1\\]","Sample OLE content exactly matches source without accumulation");}
    actual+="]\n";TaskHandle file(CreateFileW((directory+L"\\native.json").c_str(),GENERIC_WRITE,0,nullptr,CREATE_NEW,FILE_ATTRIBUTE_NORMAL,nullptr));require(file.value!=INVALID_HANDLE_VALUE,"New native inspection JSON");DWORD written{};require(WriteFile(file.value,actual.data(),static_cast<DWORD>(actual.size()),&written,nullptr)&&written==actual.size(),"Write native inspection JSON");
    const auto pdf=directory+L"\\original.pdf";invoke(word.document.Get(),L"ExportAsFixedFormat",DISPATCH_METHOD,{pdf.c_str(),17L});
    const auto blank=emptyEquation(api,word);
    auto shapes=get(word.document.Get(),L"InlineShapes");ComPtr<IDispatch> first;const size_t editedIndex=mixed?2:0;size_t ordinal{};
    for(const auto& shape:enumerateCollection(shapes.Get(),integer(shapes.Get(),L"Count"))){if(integer(shape.Get(),L"Type")!=1)continue;if(ordinal++==editedIndex){first=shape;break;}}
    require(first!=nullptr,"Editable prototype MathType shape");auto ole=get(first.Get(),L"OLEFormat");invoke(ole.Get(),L"DoVerb",DISPATCH_METHOD,{2L});auto object=get(ole.Get(),L"Object");
    require(api.writeBytes(object.Get(),3,blank.data(),static_cast<long>(blank.size()))==0,"Reset first check equation");
    const std::wstring replacement=L"\\[q+123\\]";require(api.write(object.Get(),1,replacement.c_str(),static_cast<long>(replacement.size()))==0,"Edit first independent object");
    require(api.close(OLECLOSE_NOSAVE,object.Get())==0,"Close edited check object");object.Reset();ole.Reset();first.Reset();shapes.Reset();
    invoke(word.document.Get(),L"Save",DISPATCH_METHOD);invoke(word.document.Get(),L"Close",DISPATCH_METHOD,{0L});word.document.Reset();word.document=open();
    const auto after=bodyNativeAssets(word.document.Get());require(after.size()==before.size()&&before[editedIndex]!=after[editedIndex],"Selected object persists edited content");
    for(size_t i=0;i<before.size();++i)if(i!=editedIndex)require(before[i]==after[i],"Other equation bytes unchanged after edit");
    require(exportNativeTex(api,after[editedIndex]).find("123")!=std::string::npos,"Edited equation remains editable after reopen");
    std::printf("ULTRA_REOPEN_EDIT_INDEPENDENCE_PASSED COUNT=%zu\n",after.size());word.shutdown();
}

