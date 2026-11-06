// 工作进程协议纯回归：直接消费文本记录，不创建窗口、进程或 Office 实例。
#include "worker_client.h"
#include "worker_protocol.h"
#include "gui_support.h"
#include "gui_acceptance.h"

int main() {
    try {
        WorkerState state;state.reset(ConversionMode::Standard);
        const auto progress=state.read("UI_STAGE_PROGRESS NEW_STAGE 1 2");
        require(progress.event==WorkerEvent::Progress&&progress.current==1&&progress.total==2&&!state.workerFailed,
                "Unknown display stage keeps valid task progress");
        bool rejected=false;
        try{state.read("UI_STAGE_PROGRESS NEW_STAGE 1 0");}catch(const std::runtime_error&){rejected=true;}
        require(rejected,"Malformed stage progress remains rejected");

        // 极速复用阶段计数，结果必须标记极速且未经核验，不冒充快速或标准。
        WorkerState ultra;ultra.reset(ConversionMode::Ultra);ultra.read("UI_SCAN_COUNTS 0 1000 0 0");
        const auto ultraProgress=ultra.read("UI_STAGE_PROGRESS CONVERTING 24 1000");
        require(ultraProgress.current==24&&ultraProgress.total==1000&&ultraProgress.label.find(L"24 / 1000")!=std::wstring::npos,"Ultra throttled progress carries current formula counters");
        ultra.read("UI_CONVERT_SUMMARY COMPLETE 1000 1000 0 0 0");
        QueuedDocument ultraDocument;ultraDocument.mode=ConversionMode::Ultra;ultraDocument.state=DocumentState::Complete;
        const auto ultraSummary=documentSummary(ultraDocument,L"");
        require(ultra.mode==ConversionMode::Ultra&&ultra.succeeded==1000&&ultraSummary.find(L"极速输出未经核验")!=std::wstring::npos&&ultraSummary.find(L"快速")==std::wstring::npos,"Ultra result retains its actual mode");
        ultra.read("UI_PROCESSING_ISSUE "+wireHex("正文")+" "+wireHex("SCANNING")+" "+wireHex("极速仅处理每段一个 OMML，当前段落包含多个公式，该式保留。")+" "+wireHex("正文块 2")+" "+wireHex("x+y"));
        require(ultra.problems.size()==1&&ultra.problems.front().state==L"REJECTED_PRESERVED"&&ultra.failed==0&&ultra.failures.front().find(L"该式未转换，原文保留")!=std::wstring::npos,"Ultra scanning rejection remains untouched, not a failed conversion");

        // 统一进度只更新转换计数时，自动停止仍在第一条新公式完成后触发。
        GuiAcceptance stopping;stopping.directory=L"synthetic";stopping.mode=L"stop";
        require(stopping.updated({WorkerEvent::ConversionCount,L"",1,2},nullptr,false,0,1)&&
                stopping.updated({WorkerEvent::ConversionProgress,L"",1,2},nullptr,false,0,1)&&
                !stopping.updated({WorkerEvent::Progress,L"",1,2},nullptr,false,0,1),
                "Unified formula counters preserve first-conversion cancellation without stopping on other stage progress");

        // 编号插入完成后仍要更新编号与引用；展示阶段不得改动两种模式的转换统计。
        for(bool fast:{false,true}){
            WorkerState stages;stages.reset(fast?ConversionMode::Fast:ConversionMode::Standard);stages.read("UI_SCAN_COUNTS 3 2 1 1");
            stages.read("UI_ITEM_RESULT 1 SUCCEEDED -");
            stages.read(fast?"UI_ITEM_RESULT 2 FAILED_UNVERIFIED -":"UI_ITEM_RESULT 2 FAILED_PRESERVED -");
            require(stages.read("CONVERT_PROGRESS 2/2").event==WorkerEvent::ConversionProgress,
                    "New formula conversion retains its existing progress display");
            const auto inserted=stages.read("UI_STAGE_PROGRESS NUMBER_EQUATIONS 2 2");
            const auto updating=stages.read("UI_STAGE_PROGRESS NUMBER_UPDATES 0 2");
            const auto references=stages.read("UI_STAGE_PROGRESS NUMBER_REFERENCES 0 1");
            const auto numberedText=stages.read("UI_STAGE_PROGRESS NUMBER_TEXT 1 2");
            require(inserted.current==inserted.total&&updating.event==WorkerEvent::Progress&&updating.current==0&&updating.total==2&&
                    updating.label==L"正在更新公式编号 0 / 2 项"&&references.current==0&&references.total==1&&
                    references.label==L"正在更新编号引用 0 / 1 项"&&numberedText.label==L"正在核对编号后的正文 1 / 2 项",
                    "Numbering progress continues with counter and reference updates after insertion completes");
            for(const auto* phase:{"SNAPSHOT_STRUCTURE","SNAPSHOT_OMML","SNAPSHOT_BODY","SNAPSHOT_OBJECTS","NUMBERING","NUMBER_CHECKPOINT","SAVING","CLOSING_WORK_COPY","EXITING_WORD","PUBLISHING","REPORT_BUILD","COM_UNINITIALIZE","WORD_COM_RELEASE","TASK_CLEANUP","WAIT_WORKER_EXIT"}){
                const auto update=stages.read(std::string("UI_PHASE ")+phase);
                require(update.event==WorkerEvent::Phase&&update.marquee&&update.label!=L"正在完成输出…",
                        "Snapshot and output phases leave completed item progress and display their own activity");
            }
            for(const auto* stage:{"SOURCE_POSITIONS","OMML_SNAPSHOT","BODY_SNAPSHOT"}){
                const auto update=stages.read(std::string("UI_STAGE_PROGRESS ")+stage+" 1 2");
                require(update.event==WorkerEvent::Progress&&update.current==1&&update.total==2&&!update.marquee&&
                        update.label.find(L"正在处理文档")==std::wstring::npos,
                        "Snapshot and formula processing counters have specific stage labels");
            }
            const auto processing=stages.read("UI_STAGE_PROGRESS PROCESS_FORMULAS 3 5");
            const auto conversionCount=stages.read("CONVERT_PROGRESS 2/2");
            require(processing.label==L"正在转换公式…  ·  已完成 3 / 5"&&processing.current==3&&processing.total==5&&
                    conversionCount.event==WorkerEvent::ConversionCount&&conversionCount.current==2&&conversionCount.total==2&&stages.completed==2,
                    "New conversion counters do not overwrite unified new and existing formula progress");
            stages.read("UI_PHASE NUMBERING");
            require(stages.read("CONVERT_PROGRESS 2/2").event==WorkerEvent::ConversionProgress,
                    "The next phase clears unified formula progress without changing conversion counts");
            require(unverifiedMode(stages.mode)==fast&&stages.existing==3&&stages.targets==2&&stages.inlines==1&&stages.completed==2&&
                    stages.succeeded==1&&stages.failed==1&&stages.rejected==1&&stages.unattempted==0&&
                    stages.problems.size()==1&&!stages.summary&&!stages.workerFailed,
                    "Display-only phases preserve normal and fast conversion counts and diagnostics");
            // 子步骤、完成消息及新旧公式统一进度使用相同文案，避免每式循环时文字和计数位置闪动。
            WorkerState steady;steady.reset(fast?ConversionMode::Fast:ConversionMode::Standard);steady.read("UI_SCAN_COUNTS 0 278 0 0");
            const auto stableLabel=steady.read("CONVERT_PROGRESS 9/278").label;
            require(stableLabel==L"正在转换公式…  ·  已完成 9 / 278","Formula count has one stable display format");
            for(const auto* phase:{"FORMULA_SOURCE","FORMULA_STYLE","FORMULA_PREPARING","FORMULA_INSERTING","FORMULA_IMPORTING","FORMULA_CLOSING","FORMULA_LAYOUT","FORMULA_VERIFYING","FORMULA_EXPORTING","FORMULA_REPLACING","FORMULA_CHECKPOINT"}){
                const auto update=steady.read(std::string("UI_PHASE ")+phase);
                require(update.event==WorkerEvent::Status&&update.label==stableLabel&&steady.progressStage==phase&&formulaConversionStage(steady.progressStage),
                        "Normal formula steps retain raw diagnostics while displaying identical text and counts");
            }
            require(steady.read("UI_STAGE_PROGRESS PROCESS_FORMULAS 9 278").label==stableLabel&&
                    steady.read("UI_STAGE_PROGRESS EXP_FORMULAS 9 278").label==stableLabel,
                    "Body and experimental formula counters use the same stable display");
            const auto recovering=steady.read("UI_PHASE FORMULA_RECOVERING");
            require(recovering.label!=stableLabel&&!formulaConversionStage(steady.progressStage)&&!formulaConversionStage("SAVING"),
                    "Recovery and saving retain their own visible stages");
            steady.read("UI_EXPERIMENTAL_PHASE CONVERTING");
            require(steady.read("UI_ITEM_RESULT 1 SUCCEEDED -").label==L"正在转换公式…  ·  已完成 1",
                    "Experimental fallback does not reuse the previous body formula total");
            // 每式细步骤保留总体条及错误上下文，满条之后明确切换收尾阶段。
            WorkerState tail;tail.reset(fast?ConversionMode::Fast:ConversionMode::Standard);tail.read("UI_SCAN_COUNTS 0 1 1 0");tail.read("UI_ITEM_BEGIN 1 1 0 3 247824 -");
            tail.read("UI_STAGE_PROGRESS PROCESS_FORMULAS 0 1");
            const auto importing=tail.read("UI_PHASE FORMULA_IMPORTING");
            require(importing.event==WorkerEvent::Status&&importing.label.find(L"0 / 1")!=std::wstring::npos&&tail.progressStage=="FORMULA_IMPORTING"&&tail.currentProblem.id==1&&
                    tail.read("CONVERT_PROGRESS 1/1").event==WorkerEvent::ConversionCount,
                    "Formula substeps keep unified progress and the active diagnostic context");
            tail.read("UI_ITEM_RESULT 1 SUCCEEDED -");tail.read("UI_STAGE_PROGRESS PROCESS_FORMULAS 1 1");
            const auto reportPhase=tail.read("UI_PHASE REPORT_BUILD");
            require(reportPhase.event==WorkerEvent::Phase&&reportPhase.marquee&&tail.progressStage=="REPORT_BUILD",
                    "Completed formula progress moves to the actual report stage");
            tail.read("UI_STAGE_PROGRESS PUBLISH_FILES 2 2");const auto summary=tail.read("UI_CONVERT_SUMMARY COMPLETE 1 1 0 0 0");
            require(summary.event==WorkerEvent::Phase&&summary.marquee&&tail.progressStage=="WAITING_WORKER"&&tail.summary&&tail.succeeded==1,
                    "A completed publication bar still displays backend shutdown until the worker exits");
            tail.read("UI_STAGE_PROGRESS SOURCE_COMPARE 1 2");tail.read("UI_STAGE_PROGRESS SOURCE_COMPARE 2 2");
            require(tail.progressStage=="SOURCE_COMPARE"&&tail.read("UI_PHASE FUTURE_STEP").label!=L"正在完成输出…"&&tail.progressStage=="FUTURE_STEP",
                    "Stable stage identifiers omit counters and unknown phases do not imply output completion");
            tail.cancelling=true;const auto cleaning=tail.read("UI_PHASE TASK_CLEANUP");const auto cleaned=tail.read("UI_STAGE_PROGRESS TASK_CLEANUP 1 3");
            require(cleaning.event==WorkerEvent::Phase&&cleaning.marquee&&cleaned.event==WorkerEvent::Progress&&tail.progressStage=="TASK_CLEANUP"&&
                    tail.read("UI_COMMIT_READY").event==WorkerEvent::None&&tail.read("UI_ITEM_RESULT 2 SUCCEEDED -").event==WorkerEvent::None&&
                    tail.read("UI_CONVERT_SUMMARY COMPLETE 2 2 0 0 0").event==WorkerEvent::None&&tail.targets==1&&tail.succeeded==1&&tail.completed==1,
                    "Cancellation displays cleanup without accepting results or changing conversion statistics");
            WorkerState extra;extra.reset(fast?ConversionMode::Fast:ConversionMode::Standard);extra.read("UI_EXPERIMENTAL_PHASE CONVERTING");extra.read("UI_STAGE_PROGRESS EXP_FORMULAS 0 1");
            extra.read("UI_PHASE FORMULA_IMPORTING");
            require(extra.read("UI_ITEM_RESULT 1 SUCCEEDED -").event==WorkerEvent::None&&extra.succeeded==1&&
                    extra.read("UI_STAGE_PROGRESS EXP_SAVED_PARTS 1 2").saved,
                    "Experimental item results do not overwrite counted progress and saved package checks are marked correctly");
        }

        state.reset(ConversionMode::Standard);state.read("UI_SCAN_COUNTS 0 1 1 0");
        const auto pageTwo=wireHex(utf8(L"原文第 2 页 · 正文"));
        const auto pageThree=wireHex(utf8(L"原文第 3 页 · 正文"));
        state.read("UI_ITEM_BEGIN 1 1 0 3 "+wireHex("$x$")+" "+pageTwo);
        require(!state.currentFormula.empty(),"Formula context is available before its result");
        state.read("UI_ITEM_RESULT 1 SUCCEEDED -");state.read("UI_CONVERT_SUMMARY COMPLETE 1 1 0 0 0");
        require(state.summary&&state.outcome==L"COMPLETE"&&state.succeeded==1&&state.currentFormula.empty()&&!state.workerFailed,
                "Formula result and final summary keep success counters");

        // 成功公式只留在原始日志，常规和快速摘要都不逐式展开，也不将快速完成表述为核验通过。
        QueuedDocument completedDocument;completedDocument.path=L"synthetic\\success.docx";completedDocument.state=DocumentState::Complete;
        completedDocument.problems=state.problems;completedDocument.layoutChecked=true;
        completedDocument.trace="UI_ITEM_BEGIN 1 1 0 3 "+wireHex("$x$")+" "+pageTwo+"\nUI_ITEM_RESULT 1 SUCCEEDED -\n";
        const auto completedSummary=documentSummary(completedDocument,L"");
        require(completedSummary.find(L"未发现转换问题。")!=std::wstring::npos&&
                completedSummary.find(L"原式片段：")==std::wstring::npos&&completedSummary.find(L"$x$")==std::wstring::npos&&
                completedSummary.find(L"UI_ITEM_BEGIN")==std::wstring::npos&&completedSummary.find(L"0 条版式提醒")==std::wstring::npos&&
                completedSummary.find(L"版式检查范围：")==std::wstring::npos,
                "Normal success summary stays brief and never expands successful formulas or raw records");
        WorkerState fastSuccess;fastSuccess.reset(ConversionMode::Fast);
        fastSuccess.read("UI_ITEM_BEGIN 1 1 0 3 "+wireHex("$x$")+" "+pageTwo);fastSuccess.read("UI_ITEM_RESULT 1 SUCCEEDED -");
        fastSuccess.read("UI_CONVERT_SUMMARY COMPLETE 1 1 0 0 0");
        completedDocument.mode=ConversionMode::Fast;completedDocument.layoutChecked=false;completedDocument.problems=fastSuccess.problems;
        const auto fastCompletedSummary=documentSummary(completedDocument,L"");
        require(fastSuccess.summary&&fastSuccess.succeeded==1&&fastSuccess.problems.empty()&&
                fastCompletedSummary.find(L"未报告转换错误，快速输出未经核验。")!=std::wstring::npos&&
                fastCompletedSummary.find(L"未进行版式检查。")!=std::wstring::npos&&
                fastCompletedSummary.find(L"原式片段：")==std::wstring::npos&&fastCompletedSummary.find(L"$x$")==std::wstring::npos&&
                fastCompletedSummary.find(L"UI_ITEM_BEGIN")==std::wstring::npos&&fastCompletedSummary.find(L"0 条版式提醒")==std::wstring::npos,
                "Fast success summary hides successful formulas and clearly states its unchecked result");

        // 版式提醒属于已转换结果，不能将成功公式算作失败或部分完成。
        const auto bodyRegion=wireHex(utf8(L"正文"));
        const auto wideWarning="UI_LAYOUT_WARNING 1 FORMULA_TOO_WIDE "+bodyRegion+" "+pageTwo+" "+wireHex("$x$")+" "+wireHex(utf8(L"公式宽 180.0 磅，可用宽度 120.0 磅，可能显示不全。"));
        const auto spacingWarning="UI_LAYOUT_WARNING 1 EXACT_LINE_SPACING "+bodyRegion+" "+pageTwo+" "+wireHex("$x$")+" "+wireHex(utf8(L"公式高 22.0 磅，段落固定行距 12.0 磅，可能显示不全。"));
        state.read("UI_LAYOUT_CHECKED 1");state.read(wideWarning);state.read(spacingWarning);
        require(state.layoutChecked&&state.warnings.size()==2&&state.warnings[0].id==1&&
                state.warnings[0].code==L"FORMULA_TOO_WIDE"&&state.warnings[0].region==L"正文"&&
                state.warnings[0].location==L"原文第 2 页 · 正文"&&state.warnings[0].source==L"$x$"&&
                state.warnings[0].message.find(L"180.0 磅")!=std::wstring::npos&&
                state.summary&&state.outcome==L"COMPLETE"&&state.succeeded==1&&state.failed==0&&
                !state.workerFailed&&state.problems.empty()&&state.failures.empty()&&state.error.empty(),
                "Layout warnings preserve original locations and successful conversion counters");
        QueuedDocument warningDocument;warningDocument.path=L"synthetic\\wide.docx";warningDocument.state=DocumentState::Complete;
        warningDocument.warnings=state.warnings;warningDocument.layoutChecked=state.layoutChecked;
        const auto warningSummary=documentSummary(warningDocument,L"");
        require(warningDocument.state==DocumentState::Complete&&warningSummary.find(L"0 条问题 · 2 条版式提醒")!=std::wstring::npos&&
                warningSummary.find(L"原文第 2 页")!=std::wstring::npos&&warningSummary.find(L"原文页码不代表结果页码")!=std::wstring::npos&&
                warningSummary.find(L"原式片段：$x$")!=std::wstring::npos&&warningSummary.find(L"180.0 磅")!=std::wstring::npos&&
                warningSummary.find(L"缩小该式字号")!=std::wstring::npos&&warningSummary.find(L"固定行距改为最小值或单倍行距")!=std::wstring::npos&&
                warningSummary.find(L"FORMULA_TOO_WIDE")==std::wstring::npos,
                "Completed document summary separates layout reminders and gives concrete manual adjustments");
        warningDocument.warnings.clear();warningDocument.layoutChecked=false;
        require(documentSummary(warningDocument,L"").find(L"未进行版式检查")==std::wstring::npos,
                "Absent layout protocol does not claim a reverse task was checked or skipped");

        // 相同原式用位置和编号区分；未知页不能沿用上一式的页码。
        state.reset(ConversionMode::Standard);
        state.read("UI_ITEM_REJECTED 1 2 10 13 "+wireHex("$x$")+" "+wireHex("UNBALANCED_BRACES")+" "+pageTwo);
        require(state.problems.size()==1&&state.problems[0].id==1&&state.problems[0].source==L"$x$"&&
                state.problems[0].location==L"原文第 2 页 · 正文"&&state.problems[0].reason==L"UNBALANCED_BRACES"&&
                state.problems[0].state==L"REJECTED_PRESERVED","Rejected formula retains structured original context");
        state.read("UI_ITEM_BEGIN 2 3 20 23 "+wireHex("$x$")+" "+pageThree);
        state.read("UI_ITEM_RESULT 2 FAILED_PRESERVED "+wireHex("Write TeX into new object"));
        require(state.problems.size()==2&&state.problems[1].id==2&&state.problems[1].location==L"原文第 3 页 · 正文"&&
                state.problems[1].state==L"FAILED_PRESERVED"&&state.currentProblem.id==0,
                "Conversion failure remains distinct from an earlier rejected copy of the same formula");
        state.read("UI_ITEM_BEGIN 3 4 30 33 "+wireHex("$z$")+" -");
        state.read("UI_ITEM_RESULT 3 FAILED_PRESERVED "+wireHex("Import OMML into new object"));
        require(state.problems.size()==3&&state.problems.back().location.empty()&&state.problems.back().source==L"$z$",
                "Unknown source location does not inherit the preceding formula page");
        QueuedDocument document;document.path=L"synthetic\\input.docx";document.state=DocumentState::Partial;
        document.problems=state.problems;
        auto summary=documentSummary(document,L"");
        require(summary.find(L"原文页码未知")!=std::wstring::npos&&summary.find(L"原式片段：$z$")!=std::wstring::npos&&
                summary.find(L"原因：")!=std::wstring::npos&&summary.find(L"建议：")!=std::wstring::npos&&
                summary.find(L"UNBALANCED_BRACES")==std::wstring::npos&&summary.find(L"Word 位置")==std::wstring::npos&&
                summary.find(document.path)==std::wstring::npos,
                "Human summary presents unknown pages, source and advice without raw reason codes or Word coordinates");

        state.reset(ConversionMode::Fast);
        state.read("UI_ITEM_BEGIN 1 1 0 3 "+wireHex("$x$")+" "+pageTwo);
        state.read("UI_ITEM_RESULT 1 FAILED_UNVERIFIED "+wireHex("Write TeX into new object"));
        require(state.problems.size()==1&&state.problems[0].state==L"FAILED_UNVERIFIED"&&state.failed==1,
                "Fast conversion failure never claims that the original formula was verified");
        state.read("UI_LAYOUT_CHECKED 0");
        document.mode=ConversionMode::Fast;document.problems=state.problems;document.layoutChecked=state.layoutChecked;summary=documentSummary(document,L"");
        require(!state.layoutChecked&&state.warnings.empty()&&summary.find(L"未进行版式检查")!=std::wstring::npos&&
                summary.find(L"未经核验")!=std::wstring::npos&&summary.find(L"原式已保留")==std::wstring::npos,
                "Fast failure summary does not promise verified preservation");

        // 接口中断只记录正在处理的上下文，不能把服务故障归因为这一条公式。
        state.reset(ConversionMode::Standard);
        state.read("UI_ITEM_REJECTED 1 2 10 13 "+wireHex("$x$")+" "+wireHex("UNBALANCED_BRACES")+" "+pageTwo);
        state.read("UI_ITEM_BEGIN 2 3 20 23 "+wireHex("$y$")+" "+pageThree);
        state.read("UI_ERROR "+wireHex("TASK_TIMEOUT"));
        require(state.workerFailed&&state.problems.size()==2&&state.problems.back().state==L"TASK_FAILED"&&
                state.problems.back().id==0&&state.problems.back().source.empty()&&
                state.interrupted.id==2&&state.interrupted.source==L"$y$"&&state.interrupted.reason.empty(),
                "Task interruption keeps earlier problems and separates the current formula from the task cause");
        document.mode=ConversionMode::Standard;document.state=DocumentState::Failed;document.problems=state.problems;document.interrupted=state.interrupted;
        summary=documentSummary(document,L"");
        require(summary.find(L"文件处理问题")!=std::wstring::npos&&summary.find(L"中断时正在处理")!=std::wstring::npos&&
                summary.find(L"仅表示中断位置")!=std::wstring::npos,
                "Interrupted summary distinguishes the active formula from the file failure");
        state.read("UI_LAYOUT_CHECKED 1");state.read(wideWarning);state.read("UI_BODY_RESET");
        require(state.problems.empty()&&state.failures.empty()&&state.warnings.empty()&&!state.layoutChecked&&state.currentProblem.id==0&&state.interrupted.id==0,
                "Rebuilding the body clears stale formula problems before replay");

        // 实验区重做保留正文问题，丢弃本轮实验问题，再接受最终的区域问题。
        state.reset(ConversionMode::Standard);
        state.read("UI_ITEM_REJECTED 1 2 10 13 "+wireHex("$x$")+" "+wireHex("UNBALANCED_BRACES")+" "+pageTwo);
        state.read(wideWarning);state.read("UI_LAYOUT_CHECKED 1");
        state.read("UI_EXPERIMENTAL_PHASE PREPARING");
        const auto header=wireHex(utf8(L"第 2 节页眉（主页）"));
        const auto shared=wireHex(utf8(L"第 2 节页眉（主页）（共用内容，可能出现于多页）"));
        state.read("UI_EXPERIMENTAL_REJECTED 2 "+header+" "+wireHex("$x$")+" "+wireHex("EMPTY")+" "+shared);
        state.read("UI_EXPERIMENTAL_ITEM 3 "+header+" "+wireHex("$h$")+" "+shared);
        state.read("UI_LAYOUT_WARNING 3 EXACT_LINE_SPACING "+header+" "+shared+" "+wireHex("$h$")+" "+wireHex(utf8(L"固定行距可能遮挡公式。")));
        state.read("UI_LAYOUT_CHECKED 0");
        state.read("UI_EXPERIMENTAL_RESET");
        require(state.problems.size()==1&&state.problems[0].id==1&&state.warnings.size()==1&&state.warnings[0].code==L"FORMULA_TOO_WIDE"&&
                state.layoutChecked&&state.currentProblem.id==0&&state.currentFormula.empty(),
                "Experimental replay preserves body diagnostics and removes abandoned experimental contexts");
        const auto issue="UI_PROCESSING_ISSUE "+header+" "+wireHex("RECOVERY")+" "+wireHex("REGION_SKIPPED")+" "+shared+" "+wireHex("$recover$");
        state.read(issue);state.read(issue);
        require(state.problems.size()==2&&state.problems.back().id==0&&state.problems.back().region==L"第 2 节页眉（主页）"&&
                state.problems.back().location==L"第 2 节页眉（主页）（共用内容，可能出现于多页）"&&
                state.problems.back().stage==L"RECOVERY"&&state.problems.back().state==L"PROCESSING_ISSUE"&&state.problems.back().source==L"$recover$",
                "Repeated region recovery reports retain one shared header location without inventing a page");
        state.read("UI_EXPERIMENTAL_ITEM 4 "+header+" "+wireHex("$h$")+" "+shared);
        state.read("UI_ITEM_RESULT 4 FAILED_PRESERVED "+wireHex("Write TeX into new object"));
        require(state.problems.size()==3&&state.problems.back().id==4&&state.problems.back().source==L"$h$"&&
                state.problems.back().region==L"第 2 节页眉（主页）"&&state.problems.back().location==L"第 2 节页眉（主页）（共用内容，可能出现于多页）",
                "Experimental formula failure keeps its own region and original source");
        state.read("UI_ITEM_BEGIN 5 5 40 43 "+wireHex("$q$")+" "+pageTwo);state.read("UI_ITEM_END");
        state.read("UI_ERROR "+wireHex("MATHTYPE_SERVICE_DISCONNECTED"));
        require(state.interrupted.id==0,"Ended formula context cannot leak into a later task interruption");
        state.reset(ConversionMode::Standard);
        require(state.problems.empty()&&state.warnings.empty()&&!state.layoutChecked&&state.interrupted.id==0&&state.currentProblem.id==0,
                "The next document starts without the preceding document diagnostics");

        state.reset(ConversionMode::Standard);state.cancelling=true;
        require(state.read("CONVERT_PROGRESS 1/2").event==WorkerEvent::None&&state.completed==0,
                "Cancellation ignores conversion statistics");
        state.read("UI_ERROR "+wireHex("TASK_COPY_CLEANUP_FAILED"));
        require(state.workerFailed&&!state.error.empty(),"Cancellation still reports failed task copy cleanup");
        std::puts("WORKER_PROTOCOL_PASSED");return 0;
    }catch(const std::exception& error){std::printf("WORKER_PROTOCOL_FAILED: %s\n",error.what());return 1;}
}
