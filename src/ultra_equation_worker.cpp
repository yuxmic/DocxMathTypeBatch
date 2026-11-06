// 极速独立探针与正式界面调用同一原生引擎；开发验收仅在本入口编入。
#include "ultra_conversion.h"
#include "equation_styles.h"
#include "test_ultra_native.h"

int wmain(int argc,wchar_t** argv){
    setvbuf(stdout,nullptr,_IONBF,0);if(FAILED(OleInitialize(nullptr)))return 1;int code=1;
    try{
        if(argc==2&&std::wstring(argv[1])==L"--xml-self-test"){ultraXmlSelfTest();OleUninitialize();return 0;}
        require(argc>=4,"usage: ultra-equation-worker --convert input output report size [--numbers] [--progress] | --verify input directory");
        TaskHandle mutex(CreateMutexW(nullptr,FALSE,L"Local\\DocxMathTypeBatch.MathTypePreferences.v1"));bool owned{};
        struct Unlock {TaskHandle& mutex;bool& owned;~Unlock(){if(owned)ReleaseMutex(mutex.value);}} unlock{mutex,owned};
        require(mutex.value!=nullptr,"Open MathType task mutex");const auto wait=WaitForSingleObject(mutex.value,0);require(wait==WAIT_OBJECT_0||wait==WAIT_ABANDONED,"MATHTYPE_TASK_BUSY");owned=true;
        std::wstring problem;MathTypeApi api(findMathTypeInstall(problem));
        if(std::wstring(argv[1])==L"--verify"||std::wstring(argv[1])==L"--verify-mixed"){
            require(argc==4,"Verify sample arguments");require(api.init(0,10)>0,"Initialize ultra sample API");api.initialized=true;verifyUltraSample(api,absolutePath(argv[2]),absolutePath(argv[3]),std::wstring(argv[1])==L"--verify-mixed");
        }else{
            require(std::wstring(argv[1])==L"--convert"&&argc>=6,"Ultra conversion arguments");
            const auto input=absolutePath(argv[2]),output=absolutePath(argv[3]),report=absolutePath(argv[4]);
            UltraOptions options;wchar_t* end{};options.points=wcstod(argv[5],&end);require(end&&!*end&&std::isfinite(options.points)&&options.points>=1&&options.points<=256,"Ultra size must be 1..256 pt");
            for(int i=6;i<argc;++i){const std::wstring option=argv[i];if(option==L"--numbers")options.numbers=true;else if(option==L"--progress")options.progress=true;else throw std::runtime_error("Unknown ultra option");}
            // 同一个只读句柄覆盖整个任务，副本从建立后只读取一次。
            TaskHandle source(CreateFileW(input.c_str(),GENERIC_READ,FILE_SHARE_READ,nullptr,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,nullptr));require(source.value!=INVALID_HANDLE_VALUE,"Lock original ultra input");
            require(GetFileAttributesW(report.c_str())==INVALID_FILE_ATTRIBUTES,"New ultra report required");taskCopy(input,output);
            try{ConversionTimings timings;const auto started=GetTickCount64();const auto result=convertUltra(api,output,options,timings);timings.report();
                const auto json="{\"Mode\":\"Ultra\",\"Status\":\""+std::string(result.remaining||result.failed?"PARTIAL":"COMPLETE")+"\",\"Targets\":"+std::to_string(result.targets)+",\"Succeeded\":"+std::to_string(result.succeeded)+",\"Failed\":"+std::to_string(result.failed)+",\"RemainingOMML\":"+std::to_string(result.remaining)+",\"NumbersAdded\":"+std::to_string(result.numbersAdded)+",\"Validation\":\"NONE\",\"WordStarted\":false,\"Seconds\":"+std::to_string((GetTickCount64()-started)/1000.0)+"}\n";
                TaskHandle file(CreateFileW(report.c_str(),GENERIC_WRITE,0,nullptr,CREATE_NEW,FILE_ATTRIBUTE_NORMAL,nullptr));require(file.value!=INVALID_HANDLE_VALUE,"Create new ultra report");DWORD written{};require(WriteFile(file.value,json.data(),static_cast<DWORD>(json.size()),&written,nullptr)&&written==json.size(),"Write ultra report");
            }catch(...){DeleteFileW(output.c_str());throw;}
        }code=0;
    }catch(const std::exception& error){std::printf("ULTRA_FAILED %s\n",error.what());}
    OleUninitialize();return code;
}
