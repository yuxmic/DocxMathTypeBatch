// 用合成小样确认 MathML 导入和跨区域对象替换；不修改用户文档。
#include "equation_objects.h"

static void createExperimentalFixture(const std::wstring& input,const std::wstring& directory) {
    const auto output=directory+L"\\complex-input.docx";require(CopyFileW(input.c_str(),output.c_str(),TRUE)!=FALSE,"Create independent complex experimental fixture");
    WordSession word;auto documents=get(word.application.Get(),L"Documents");
    word.document=asObject(invoke(documents.Get(),L"Open",DISPATCH_METHOD,{output.c_str(),false,false,false}));word.verifyOwnership();
    auto document=word.document.Get();auto shapes=get(document,L"Shapes");auto anchor=docRange(document,integer(get(document,L"Content").Get(),L"End")-1,integer(get(document,L"Content").Get(),L"End")-1);
    auto first=asObject(invoke(shapes.Get(),L"AddTextbox",DISPATCH_METHOD,{1L,10.0,20.0,120.0,44.0,anchor.Get()}));
    auto second=asObject(invoke(shapes.Get(),L"AddTextbox",DISPATCH_METHOD,{1L,150.0,20.0,120.0,44.0,anchor.Get()}));
    auto firstFrame=get(first.Get(),L"TextFrame"),secondFrame=get(second.Get(),L"TextFrame");invoke(firstFrame.Get(),L"Next",DISPATCH_PROPERTYPUT,{secondFrame.Get()});
    invoke(get(firstFrame.Get(),L"ContainingRange").Get(),L"Text",DISPATCH_PROPERTYPUT,{L"Linked $l_1$ first\rLinked $l_2$ second"});
    auto table=collectionItem(get(document,L"Tables").Get(),1);auto merged=asObject(invoke(table.Get(),L"Cell",DISPATCH_METHOD,{1L,1L}));
    auto next=asObject(invoke(table.Get(),L"Cell",DISPATCH_METHOD,{1L,2L}));invoke(merged.Get(),L"Merge",DISPATCH_METHOD,{next.Get()});
    auto end=get(merged.Get(),L"Range");invoke(end.Get(),L"Collapse",DISPATCH_METHOD,{0L});invoke(end.Get(),L"MoveStart",DISPATCH_METHOD,{1L,-1L});invoke(end.Get(),L"Collapse",DISPATCH_METHOD,{1L});
    auto nested=asObject(invoke(get(document,L"Tables").Get(),L"Add",DISPATCH_METHOD,{end.Get(),1L,1L}));invoke(cellRange(nested.Get(),1,1).Get(),L"Text",DISPATCH_PROPERTYPUT,{L"Nested $n_1$ tail"});
    auto section=collectionItem(get(document,L"Sections").Get(),1);auto header=collectionItem(get(section.Get(),L"Headers").Get(),1);auto headerRange=get(header.Get(),L"Range");const long headerEnd=integer(headerRange.Get(),L"End")-1;invoke(headerRange.Get(),L"SetRange",DISPATCH_METHOD,{headerEnd,headerEnd});
    auto headerTable=asObject(invoke(get(headerRange.Get(),L"Tables").Get(),L"Add",DISPATCH_METHOD,{headerRange.Get(),1L,1L}));invoke(cellRange(headerTable.Get(),1,1).Get(),L"Text",DISPATCH_PROPERTYPUT,{L"Header table $h_2$ tail"});
    auto headerAnchor=get(header.Get(),L"Range");const long anchorEnd=integer(headerAnchor.Get(),L"End")-1;invoke(headerAnchor.Get(),L"SetRange",DISPATCH_METHOD,{anchorEnd,anchorEnd});
    auto box=asObject(invoke(get(header.Get(),L"Shapes").Get(),L"AddTextbox",DISPATCH_METHOD,{1L,10.0,10.0,160.0,42.0,headerAnchor.Get()}));invoke(get(get(box.Get(),L"TextFrame").Get(),L"TextRange").Get(),L"Text",DISPATCH_PROPERTYPUT,{L"Header box $h_3$ tail"});
    auto setup=get(section.Get(),L"PageSetup");invoke(setup.Get(),L"DifferentFirstPageHeaderFooter",DISPATCH_PROPERTYPUT,{-1L});invoke(setup.Get(),L"OddAndEvenPagesHeaderFooter",DISPATCH_PROPERTYPUT,{-1L});
    for(const auto* name:{L"Headers",L"Footers"})for(long kind=2;kind<=3;++kind){auto variant=collectionItem(get(section.Get(),name).Get(),kind);const auto label=std::wstring(name)+std::to_wstring(kind)+L" $v_"+std::to_wstring(kind)+L"$ tail";invoke(get(variant.Get(),L"Range").Get(),L"Text",DISPATCH_PROPERTYPUT,{label.c_str()});}
    const long documentEnd=integer(get(document,L"Content").Get(),L"End")-1;invoke(get(document,L"Sections").Get(),L"Add",DISPATCH_METHOD,{docRange(document,documentEnd,documentEnd).Get(),2L});
    const long notePosition=integer(get(document,L"Content").Get(),L"End")-1;
    auto note=asObject(invoke(get(document,L"Footnotes").Get(),L"Add",DISPATCH_METHOD,{docRange(document,notePosition,notePosition).Get(),Value(),L"Second note $u_1$ tail"}));
    auto noteRange=get(note.Get(),L"Range");const long noteEnd=integer(noteRange.Get(),L"End")-1;invoke(noteRange.Get(),L"SetRange",DISPATCH_METHOD,{noteEnd,noteEnd});
    auto noteTable=asObject(invoke(get(get(note.Get(),L"Range").Get(),L"Tables").Get(),L"Add",DISPATCH_METHOD,{noteRange.Get(),1L,1L}));invoke(cellRange(noteTable.Get(),1,1).Get(),L"Text",DISPATCH_PROPERTYPUT,{L"Second note table $u_2$ tail"});
    invoke(document,L"Save",DISPATCH_METHOD);word.shutdown();std::puts("EXPERIMENTAL_COMPLEX_FIXTURE_CREATED");
}

static void compareExperimentalNative(const std::wstring& first,const std::wstring& second) {
    const auto before=inspectPackage(first),after=inspectPackage(second);std::vector<std::vector<unsigned char>> left,right;
    for(const auto& value:before.equations)left.push_back(nativePayload(value));for(const auto& value:after.equations)right.push_back(nativePayload(value));
    std::sort(left.begin(),left.end());std::sort(right.begin(),right.end());require(!left.empty()&&left==right,"Basic font, spacing and numbering do not change experimental native equations");
    std::puts("EXPERIMENTAL_NATIVE_OPTIONS_ISOLATED");
}

static void checkExperimentalApi(const std::wstring& directory) {
    WordSession word;
    auto documents=get(word.application.Get(),L"Documents");
    word.document=asObject(invoke(documents.Get(),L"Add",DISPATCH_METHOD,{Value(),false,Value(),false}));word.verifyOwnership();
    invoke(word.document.Get(),L"Activate",DISPATCH_METHOD);
    auto scratch=asObject(invoke(documents.Get(),L"Add",DISPATCH_METHOD,{Value(),false,Value(),false}));
    std::wstring problem;MathTypeApi api(findMathTypeInstall(problem));require(problem.empty(),"MathType installation");
    require(api.init(0,10)>0,"Initialize experimental API probe");api.initialized=true;taskMathTypeConnected();
    auto blank=docRange(scratch.Get(),0,0);
    invoke(blank.Get(),L"InsertFile",DISPATCH_METHOD,{(api.install+L"\\Office Support\\BlankEqn.doc").c_str(),L"MTBlankEqn",false,false,false});
    auto fields=get(scratch.Get(),L"Fields");
    for(long i=integer(fields.Get(),L"Count");i>0;--i){auto field=collectionItem(fields.Get(),i);if(integer(field.Get(),L"Type")==35)invoke(field.Get(),L"Unlink",DISPATCH_METHOD);}
    auto shape=collectionItem(get(scratch.Get(),L"InlineShapes").Get(),1);auto ole=get(shape.Get(),L"OLEFormat");
    invoke(ole.Get(),L"DoVerb",DISPATCH_METHOD,{2L});auto object=get(ole.Get(),L"Object");
    const std::wstring mathml=L"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mi>x</mi><mn>2</mn></mfrac></math>";
    const auto status=api.write(object.Get(),2,mathml.c_str(),static_cast<long>(mathml.size()));
    std::printf("MATHML_IMPORT_STATUS=%ld\n",status);require(status==0,"Import MathML with vendor language 2");
    const auto native=readNative(api,object.Get());const auto tex=exportNativeTex(api,native);require(tex.find("frac")!=std::string::npos,"MathML import retains fraction structure");
    require(api.close(OLECLOSE_NOSAVE,object.Get())==0,"Close experimental scratch equation");
    auto formatted=get(get(shape.Get(),L"Range").Get(),L"FormattedText");
    invoke(get(word.document.Get(),L"Content").Get(),L"Text",DISPATCH_PROPERTYPUT,{L"left $x$ right\r"});
    auto table=asObject(invoke(get(word.document.Get(),L"Tables").Get(),L"Add",DISPATCH_METHOD,{docRange(word.document.Get(),integer(get(word.document.Get(),L"Content").Get(),L"End")-1,integer(get(word.document.Get(),L"Content").Get(),L"End")-1).Get(),1L,1L}));
    auto cell=cellRange(table.Get(),1,1);invoke(cell.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"cell $x$ tail"});
    auto section=collectionItem(get(word.document.Get(),L"Sections").Get(),1);auto header=get(collectionItem(get(section.Get(),L"Headers").Get(),1).Get(),L"Range");
    invoke(header.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"header $x$ tail"});
    auto replace=[&](IDispatch* root){auto range=get(root,L"Duplicate");auto find=get(range.Get(),L"Find");invoke(find.Get(),L"Text",DISPATCH_PROPERTYPUT,{L"$x$"});invoke(find.Get(),L"Wrap",DISPATCH_PROPERTYPUT,{0L});invoke(find.Get(),L"MatchWildcards",DISPATCH_PROPERTYPUT,{false});auto found=invoke(find.Get(),L"Execute",DISPATCH_METHOD);require(found.v.vt==VT_BOOL&&found.v.boolVal,"Find experimental placeholder");invoke(range.Get(),L"FormattedText",DISPATCH_PROPERTYPUT,{formatted.Get()});require(integer(get(range.Get(),L"InlineShapes").Get(),L"Count")==1,"FormattedText copies editable object into region");};
    replace(get(word.document.Get(),L"Content").Get());replace(cell.Get());replace(header.Get());
    invoke(scratch.Get(),L"Close",DISPATCH_METHOD,{0L});scratch.Reset();
    invoke(word.document.Get(),L"SaveAs2",DISPATCH_METHOD,{(directory+L"\\experimental-api.docx").c_str(),16L});
    word.shutdown();require(api.term()==0,"Terminate experimental API probe");api.initialized=false;
    std::puts("EXPERIMENTAL_API_PROBE_PASSED");
}
