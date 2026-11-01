// GUI 验收驱动独立持有测试状态，仅通过窗口消息操作真实控件。
#pragma once
#include "gui_support.h"
#include "scan_file.h"
#include <commctrl.h>
#include <share.h>
#include <cmath>

struct GuiAcceptanceView {
    HWND window;UINT dpi;
    const bool& reverseMode;const bool& progressMarquee;
    bool taskEmpty;
    const WorkerState& run;
    const std::vector<QueuedDocument>& documents;
    const std::wstring& install;const std::wstring& input;const std::wstring& output;
};

struct GuiAcceptance {
    std::wstring directory,mode;
    std::vector<std::wstring> inputs;
    bool batch{},initialReverse{};
    int step{};
    size_t refreshes{},backgroundPaints{};
    bool scanProgress{},savedProgress{},reopenMarquee{},parameterPaint{};
    bool active() const {return !directory.empty();}
    bool interactive() const {return active()&&!batch&&mode!=L"preview";}

    void configure(int argc,wchar_t** argv) {
        if(argc==5&&std::wstring(argv[1])==L"--ui-test"){
            inputs.push_back(absolutePath(argv[2]));directory=absolutePath(argv[3]);mode=argv[4];
            initialReverse=mode.rfind(L"reverse-",0)==0;
            if(mode==L"reverse-bracket"){initialReverse=false;step=-1;}
            require(mode==L"normal"||mode==L"stop"||mode==L"reverse-bracket"||mode==L"reverse-dollar"||mode==L"reverse-stop"||mode==L"preview","Known GUI test mode");
        }else if(argc>=5&&std::wstring(argv[1])==L"--ui-batch-test"){
            batch=true;directory=absolutePath(argv[2]);mode=L"batch-"+std::wstring(argv[3]);initialReverse=std::wstring(argv[3])==L"reverse";
            for(int i=4;i<argc;++i)inputs.push_back(absolutePath(argv[i]));
        }else require(argc==1,"Launch GUI without arguments, or use its documented worker/test entry");
    }
    void progress(HWND window,size_t current) const {
        if(active())require((GetWindowLongPtrW(GetDlgItem(window,Progress),GWL_STYLE)&PBS_MARQUEE)==0&&SendMessageW(GetDlgItem(window,Progress),PBM_GETPOS,0,0)==static_cast<LRESULT>(current),"Native progress bar shows reported count");
    }
    void refresh() {if(active())++refreshes;}
    void painted(HDC destination) {if(!destination&&active())++backgroundPaints;}
    bool updated(const WorkerUpdate& update,HWND window,bool marquee,size_t batchPosition,size_t completed) {
        if(!active())return false;
        if(update.reopening)reopenMarquee=marquee&&(GetWindowLongPtrW(GetDlgItem(window,Progress),GWL_STYLE)&PBS_MARQUEE)!=0;
        if(update.scanning&&update.current)scanProgress=true;if(update.saved&&update.current)savedProgress=true;
        return update.event==WorkerEvent::ConversionProgress&&(mode==L"stop"||mode==L"reverse-stop"||(batch&&mode==L"batch-stop"&&batchPosition==1))&&completed==1;
    }
    void trace(const std::string& value) const {
        if(!active())return;
        FILE* file=_wfsopen((directory+L"\\gui-worker.log").c_str(),L"ab",_SH_DENYNO);require(file!=nullptr,"Open test worker trace");
        const bool written=fwrite(value.data(),1,value.size(),file)==value.size()&&fputc('\n',file)!=EOF;
        const bool closed=fclose(file)==0;require(written&&closed,"Write complete test worker trace");
    }
    void batchStarted(size_t index) const {if(batch)trace("UI_BATCH_BEGIN "+std::to_string(index));}
    void scanned(const WorkerState& run,size_t index) const {
        if(active())saveNew(directory+(batch?L"\\scan-"+std::to_wstring(index)+L".json":L"\\gui-scan.json"),"{\"Ready\":"+std::to_string(run.targets)+",\"Rejected\":"+std::to_string(run.rejected)+",\"ExistingMathType\":"+std::to_string(run.existing)+"}\n");
    }
    void readyBatch(const GuiAcceptanceView& view) const {
        require(!view.run.busy&&view.taskEmpty&&!view.run.scanned,"Batch selection does not launch Office");
        saveNew(directory+L"\\gui-ready.json","{\"NoPreScan\":true,\"Documents\":"+std::to_string(view.documents.size())+"}\n");SetTimer(view.window,2,1500,nullptr);
    }
    void finishedBatch(HWND window,ULONGLONG elapsed,bool stopped,const std::vector<QueuedDocument>& documents,const std::wstring& environment) const {
        if(!batch)return;
        std::string json="{\"Status\":\"PASSED\",\"ElapsedMilliseconds\":"+std::to_string(elapsed)+",\"Stopped\":"+(stopped?std::string("true"):std::string("false"))+",\"Items\":[";
        for(size_t i=0;i<documents.size();++i){const auto& item=documents[i];json+=(i?",":"")+std::string("{\"Input\":")+quote(item.path)+",\"State\":"+std::to_string(static_cast<int>(item.state))+",\"Output\":"+quote(item.output)+",\"Error\":"+quote(item.error)+",\"ExitCode\":"+std::to_string(item.exitCode)+",\"Stage\":"+quote(item.stage)+",\"Summary\":"+quote(documentSummary(item,environment))+",\"Details\":"+quote(documentDetails(item,environment))+"}";}
        saveNew(directory+L"\\batch-check.json",json+"]}\n");if(mode!=L"batch-diagnostics")SetTimer(window,3,2000,nullptr);
    }
    void keepReport(const std::wstring& report,size_t index) const {
        if(active())require(CopyFileW(report.c_str(),(directory+(batch?L"\\report-"+std::to_wstring(index)+L".json":L"\\report.json")).c_str(),TRUE)!=0,"Keep test report evidence");
    }
    void finishedItem(const GuiAcceptanceView& view,DWORD code,const std::wstring& task,bool clean,bool cancelled,bool closePending,size_t index) {
        if(batch)saveNew(directory+L"\\item-"+std::to_wstring(index)+L".json","{\"Code\":"+std::to_string(code)+",\"TaskDirectory\":"+quote(task)+",\"Cleaned\":"+(clean?std::string("true"):std::string("false"))+"}\n");
        if(cancelled&&active()&&!batch&&!closePending){
            require(clean&&GetFileAttributesW(view.output.c_str())==INVALID_FILE_ATTRIBUTES,"Cancelled task files are removed");
            saveNew(directory+L"\\gui-check.json","{\"Status\":\"PASSED\",\"Outcome\":\"CANCELLED\",\"Input\":"+quote(view.input)+",\"Output\":"+quote(view.output)+",\"NoOutput\":true,\"NoReport\":true,\"ControlValidation\":true,\"SizeArrows\":true,\"ParameterRepaint\":true,\"KeyboardTabNavigation\":true,\"Layout\":true,\"NoPreScan\":true}\n");step=2;SetTimer(view.window,3,1500,nullptr);
        }
    }
    void failed(const std::wstring& error) const {
        saveNew(directory+L"\\gui-error.json","{\"Error\":"+quote(error)+"}\n");
    }
    void advance(const GuiAcceptanceView& view) {
        const auto window=view.window;
        const auto& run=view.run;
        const auto& reverseMode=view.reverseMode;
        const auto& progressMarquee=view.progressMarquee;
        const auto& documents=view.documents;
        const auto& install=view.install;
        const auto& input=view.input;
        const auto& output=view.output;
        auto control=[&](int id){return GetDlgItem(window,id);};
        auto set=[&](int id,const std::wstring& value){if(windowText(control(id))!=value)SetWindowTextW(control(id),value.c_str());};
        auto refresh=[&]{SendMessageW(window,WM_COMMAND,MAKEWPARAM(ModeBox,CBN_SELCHANGE),reinterpret_cast<LPARAM>(control(ModeBox)));};
        auto layout=[&]{SendMessageW(window,WM_COMMAND,MAKEWPARAM(FollowSize,CBN_SELCHANGE),reinterpret_cast<LPARAM>(control(FollowSize)));};
        auto px=[&](int value){return MulDiv(value,static_cast<int>(view.dpi),96);};
        require(!run.workerFailed,"GUI worker acceptance");
        if(step==0){std::string trace;appendWorkerLog(trace,std::string(MaximumLogBytes+32,'x'));appendWorkerLog(trace,"LAST_FAILURE");require(trace.size()<=MaximumLogBytes&&trace.find("LAST_FAILURE")!=std::string::npos&&trace.find("truncated")!=std::string::npos,"Worker log is bounded and retains its final error");}
        if(step==-1){step=0;SendMessageW(control(Reverse),BM_CLICK,0,0);require(reverseMode&&!run.busy,"Direction switch does not launch Office");return;}
        if(step==0){
            RECT client{};GetClientRect(window,&client);
            for(HWND child=GetWindow(window,GW_CHILD);child;child=GetWindow(child,GW_HWNDNEXT)){
                if(!(GetWindowLongPtrW(child,GWL_STYLE)&WS_VISIBLE))continue;
                RECT bounds{};GetWindowRect(child,&bounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&bounds),2);
                require(bounds.left>=0&&bounds.top>=0&&bounds.right<=client.right&&bounds.bottom<=client.bottom,"Visible controls fit client area");
            }
            require(!run.busy&&view.taskEmpty&&!run.scanned,"File selection does not launch a scan");
            require(GetNextDlgTabItem(window,control(Choose),FALSE)==control(AddFolder),"Keyboard tab navigation");
            if(!reverseMode){
                // 先完成初次绘制，再确认参数交互不会使父窗口重画或覆盖子控件。
                ShowWindow(window,SW_SHOWNOACTIVATE);RedrawWindow(window,nullptr,nullptr,RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW);
                require((GetWindowLongPtrW(window,GWL_STYLE)&WS_CLIPCHILDREN)!=0,"Parent drawing excludes child controls");
                RECT sizeBounds{};GetWindowRect(control(SizeBox),&sizeBounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&sizeBounds),2);
                HDC parentDc=GetDC(window);const bool excluded=!PtVisible(parentDc,(sizeBounds.left+sizeBounds.right)/2,(sizeBounds.top+sizeBounds.bottom)/2);ReleaseDC(window,parentDc);
                require(excluded,"Parent DC cannot overwrite the font size input");
                // 按实际命中控件发送鼠标消息，覆盖输入框遮挡按钮和增减方向的问题。
                auto arrow=[&](bool up){RECT bounds{};GetWindowRect(control(SizeSpin),&bounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&bounds),2);
                    POINT point{(bounds.left+bounds.right)/2,up?bounds.top+(bounds.bottom-bounds.top)/4:bounds.top+3*(bounds.bottom-bounds.top)/4};
                    HWND hit=ChildWindowFromPointEx(window,point,CWP_SKIPINVISIBLE|CWP_SKIPDISABLED);require(hit==control(SizeSpin),"Font size arrow receives mouse click");
                    MapWindowPoints(window,hit,&point,1);const LPARAM position=MAKELPARAM(point.x,point.y);
                    SendMessageW(hit,WM_LBUTTONDOWN,MK_LBUTTON,position);SendMessageW(hit,WM_LBUTTONUP,0,position);};
                auto sizeIs=[&](double expected){require(wcstod(windowText(control(SizeBox)).c_str(),nullptr)==expected,"Font size arrow value");};
                require(SendMessageW(control(FollowSize),CB_GETCURSEL,0,0)==0&&!IsWindowEnabled(control(SizeBox)),"Default size follows source");
                SendMessageW(control(FollowSize),CB_SETCURSEL,1,0);layout();
                set(SizeBox,L"12");const auto refreshesBefore=refreshes;arrow(true);sizeIs(13);
                require(refreshes==refreshesBefore+1,"Font size arrow refreshes state only once");arrow(false);sizeIs(12);
                set(SizeBox,L"11.5");arrow(true);sizeIs(12.5);arrow(false);sizeIs(11.5);
                SendMessageW(control(SizeBox),WM_KEYDOWN,VK_UP,0);SendMessageW(control(SizeBox),WM_KEYUP,VK_UP,0);sizeIs(12.5);
                SendMessageW(control(SizeBox),WM_KEYDOWN,VK_DOWN,0);SendMessageW(control(SizeBox),WM_KEYUP,VK_DOWN,0);sizeIs(11.5);
                set(SizeBox,L"1");arrow(false);sizeIs(1);set(SizeBox,L"256");arrow(true);sizeIs(256);
                SendMessageW(control(ModeBox),CB_SETCURSEL,1,0);refresh();require(!IsWindowEnabled(control(AddNumbers)),"Inline mode has no right numbers");
                SendMessageW(control(ModeBox),CB_SETCURSEL,0,0);set(SizeBox,L"0");refresh();require(!IsWindowEnabled(control(Start)),"Reject zero size");
                set(SizeBox,L"NaN");refresh();require(!IsWindowEnabled(control(Start)),"Reject nonfinite size");
                set(SizeBox,L"12");
                RedrawWindow(window,nullptr,nullptr,RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW);const auto paintsBefore=backgroundPaints;
                for(int change=0;change<20;++change){SendMessageW(control(ModeBox),CB_SETCURSEL,change%2?0:2,0);
                    SendMessageW(window,WM_COMMAND,MAKEWPARAM(ModeBox,CBN_SELCHANGE),reinterpret_cast<LPARAM>(control(ModeBox)));
                    arrow(change%2==0);UpdateWindow(window);}
                require(backgroundPaints==paintsBefore,"Mode and size interactions do not repaint the parent");parameterPaint=true;
            }else{require(!(GetWindowLongPtrW(control(ModeBox),GWL_STYLE)&WS_VISIBLE)&&(GetWindowLongPtrW(control(Bracket),GWL_STYLE)&WS_VISIBLE),"Reverse format controls");
                const bool dollar=SendMessageW(control(Dollar),BM_GETCHECK,0,0)==BST_CHECKED;
                CheckRadioButton(window,Bracket,Dollar,0);refresh();require(!IsWindowEnabled(control(Start)),"Reject missing delimiter");
                CheckRadioButton(window,Bracket,Dollar,dollar?Dollar:Bracket);set(SizeBox,L"NaN");refresh();require(IsWindowEnabled(control(Start)),"Ignore reverse hidden size settings");}
            set(SizeBox,L"18");set(NameBox,L":");refresh();require(!IsWindowEnabled(control(Start)),"Reject invalid output suffix");
            set(NameBox,L"_"+mode);refresh();require(IsWindowEnabled(control(Start)),"Enable without pre-scan");
            saveNew(directory+L"\\gui-ready.json","{\"NoPreScan\":true}\n");step=1;SetTimer(window,2,1500,nullptr);return;
        }
        require(run.summary&&(run.outcome==L"COMPLETE"||run.outcome==L"PARTIAL"),"GUI outcome");
        require(scanProgress&&savedProgress&&reopenMarquee&&!progressMarquee,"Scan and saved verification counters with marquee Word reopen");
        refresh();RECT client{};GetClientRect(window,&client);
        for(int id:{OpenOutput,OpenFolder,OpenReport})if(GetWindowLongPtrW(control(id),GWL_STYLE)&WS_VISIBLE){
            RECT bounds{};GetWindowRect(control(id),&bounds);MapWindowPoints(nullptr,window,reinterpret_cast<POINT*>(&bounds),2);
            require(bounds.bottom<=client.bottom-px(96),"Result buttons stay inside their card");}
        if(!reverseMode){
            // 完成摘要显示后改字号同样只更新输入框，保留公式计数和结果按钮。
            refresh();
            RedrawWindow(window,nullptr,nullptr,RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW);const auto paintsBefore=backgroundPaints;
            const auto savedSize=windowText(control(SizeBox));set(SizeBox,L"18.5");set(SizeBox,savedSize);UpdateWindow(window);
            require(backgroundPaints==paintsBefore&&run.scanned&&!documents.front().output.empty(),"Changing size preserves completed summary without background repaint");
        }
        std::string json="{\"Status\":\"PASSED\",\"MathTypeDirectory\":"+quote(install)+",\"Input\":"+quote(input)+",\"Output\":"+quote(output)+",\"Report\":"+quote(directory+L"\\report.json")+
            ",\"Outcome\":"+quote(run.outcome)+(reverseMode?",\"Direction\":\"REVERSE\",\"DisplayDelimiter\":"+quote(mode==L"reverse-dollar"?L"dollar":L"bracket"):",\"DefaultFont\":true,\"RequestedSizePt\":18")+
            ",\"ControlValidation\":true,\"ScanProgress\":true,\"SavedProgress\":true,\"ReopenMarquee\":true,\"SizeArrows\":"+(reverseMode?std::string("false"):std::string("true"))+",\"ParameterRepaint\":"+(parameterPaint?std::string("true"):std::string("false"))+",\"KeyboardTabNavigation\":true,\"Layout\":true,\"NoPreScan\":true,\"Succeeded\":"+std::to_string(run.succeeded)+",\"FailedPreserved\":"+std::to_string(run.failed)+",\"NotAttempted\":"+std::to_string(run.unattempted)+"}\n";
        saveNew(directory+L"\\gui-check.json",json);step=2;SetTimer(window,3,1500,nullptr);
    }

};
