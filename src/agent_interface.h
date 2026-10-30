// 轻量查询不启动 COM；Skill 通过包内相对路径调用程序。
#pragma once
#include "app_version.h"

inline int agentDescribe(int argc,wchar_t** argv) {
    if(argc!=3||std::wstring(argv[2])!=L"--json"){
        std::puts("{\"Status\":\"FAILED\",\"ErrorCode\":\"INVALID_ARGUMENT\",\"Message\":\"Use --describe --json\"}");return 2;
    }
    std::puts("{\"Application\":\"DocxMathTypeBatch\",\"Version\":\"" DMTB_VERSION "\",\"CliContractVersion\":" DMTB_CLI_CONTRACT ",\"ReportSchemaVersion\":1,"
        "\"Platform\":\"windows-x64\",\"Directions\":[\"ToMathType\",\"ToLatex\"],\"Scopes\":[\"all\",\"inline\",\"display\"],"
        "\"BodyOmmlDefault\":true,\"ExperimentalFastConversion\":true,\"ExperimentalRegions\":[\"tables\",\"textboxes\",\"headers\",\"notes\"],"
        "\"ReportStatuses\":[\"COMPLETE\",\"PARTIAL\"],\"Cancellation\":\"STOP_FILE\",\"Timeout\":\"NO_PROGRESS\"}");return 0;
}
inline int agentHelp() {
    std::puts("DocxMathTypeBatch " DMTB_VERSION " (Windows x64)\n"
        "No arguments: open the graphical interface.\n"
        "--describe --json: query identity and CLI contract without starting Word.\n"
        "--worker-convert <input.docx> <new-output.docx> <new-report.json> <paragraph|size=pt> <stop.flag> [options]\n"
        "--worker-reverse <input.docx> <new-output.docx> <new-report.json> <bracket|dollar> <stop.flag> [options]\n"
        "--worker-scan <input.docx> <new-report.json> <forward|reverse> <stop.flag>\n"
        "Options: --scope=all|inline|display --numbers --number-custom\n"
        "  --number-format=(#E1) --number-start=1 --number-chapter=1 --number-section=1\n"
        "  --inline-spacing=follow --display-spacing=follow\n"
        "  --experimental-tables --experimental-textboxes --experimental-headers --experimental-notes\n"
        "  --experimental-fast: save without pre/post content, layout or package verification (may contain errors).\n"
        "Body OMML is included in forward conversion. Extra regions only replace formulas.\n"
        "Scan covers the basic body only. Output/report paths must be new.\n"
        "Exit 0 requires reading COMPLETE/PARTIAL in the report. FastConversion=true means unverified output.\n"
        "Exit 124: idle timeout; 125: service lost; 126: recovery required; 127: cancelled.\n"
        "Create the stop file to cancel and wait for supervision cleanup.");return 0;
}
