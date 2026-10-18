#requires -Version 7.0
# 阶段 6.2.1：真实任务 Word 断连，以及暂停 COM 服务后关闭 GUI、恢复并安全停止。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$faultTempRoot=[IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output\work'))
$faultIdentity=[Guid]::NewGuid().ToString('N')
$faultTask=Join-Path $faultTempRoot ('服务故障-'+$faultIdentity)
$faultPublish=Join-Path $ProjectRoot ('tests\output\服务故障-'+$faultIdentity)
New-Item -ItemType Directory -Path $faultTask|Out-Null
$faultEvidence=[ordered]@{CheckedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');Status='FAILED';Scope='阶段6.2.1真实任务Word断连与可恢复阻塞';Cases=@();Checks=@{};SampleDirectory=$faultPublish;AutomaticHardTimeoutVerified=$false}
$faultSentinel=$null;$faultSentinelDoc=$null;$faultSentinelOwned=$false

# 开发检查使用系统进程/线程句柄；所有破坏动作都先核对进程映像和精确启动时间。
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public static class WordFaultCheck {
    [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr OpenProcess(uint access,bool inherit,uint pid);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool GetProcessTimes(IntPtr process,out long created,out long exited,out long kernel,out long user);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool QueryFullProcessImageName(IntPtr process,uint flags,StringBuilder path,ref uint size);
    [DllImport("kernel32.dll")] public static extern uint WaitForSingleObject(IntPtr handle,uint milliseconds);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool TerminateProcess(IntPtr process,uint code);
    [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll")] static extern IntPtr OpenThread(uint access,bool inherit,uint id);
    [DllImport("kernel32.dll")] static extern uint GetProcessIdOfThread(IntPtr thread);
    [DllImport("kernel32.dll")] static extern uint SuspendThread(IntPtr thread);
    [DllImport("kernel32.dll")] static extern uint ResumeThread(IntPtr thread);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("user32.dll",SetLastError=true)] public static extern bool PostMessage(IntPtr window,uint message,IntPtr wparam,IntPtr lparam);
    [DllImport("user32.dll",SetLastError=true)] public static extern IntPtr SendMessageTimeout(IntPtr window,uint message,IntPtr wparam,IntPtr lparam,uint flags,uint timeout,out IntPtr result);
    public delegate bool EnumProc(IntPtr window,IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr parameter);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window,StringBuilder name,int size);
    public static IntPtr Window(uint pid) {
        IntPtr found=IntPtr.Zero;
        EnumWindows((window,p)=>{uint owner;GetWindowThreadProcessId(window,out owner);var name=new StringBuilder(256);GetClassName(window,name,256);if(owner==pid&&name.ToString()=="DocxMathTypeBatch.B"){found=window;return false;}return true;},IntPtr.Zero);
        return found;
    }
    public static IntPtr VerifiedWord(uint pid,long started,long earliest) {
        IntPtr handle=OpenProcess(0x00101001,false,pid);
        if(handle==IntPtr.Zero) throw new Exception("无法打开任务 Word 句柄");
        try {
            long created,exited,kernel,user;uint size=32768;var image=new StringBuilder((int)size);
            if(!GetProcessTimes(handle,out created,out exited,out kernel,out user)||created!=started||created<earliest||
               !QueryFullProcessImageName(handle,0,image,ref size)||!image.ToString().EndsWith("\\WINWORD.EXE",StringComparison.OrdinalIgnoreCase)||
               WaitForSingleObject(handle,0)!=258) throw new Exception("Word 身份不符，拒绝操作");
            return handle;
        } catch {CloseHandle(handle);throw;}
    }
    public static void Kill(IntPtr verified) {
        if(WaitForSingleObject(verified,0)==258&&!TerminateProcess(verified,78)) throw new Exception("任务 Word 断连失败");
        if(WaitForSingleObject(verified,10000)!=0) throw new Exception("任务 Word 未退出");
    }
    public static List<IntPtr> Pause(uint pid,IntPtr verified) {
        var suspended=new List<IntPtr>();
        try {
            // 仅暂停已核实 Word 的现有线程，并保存句柄；不根据可复用的线程 ID 恢复。
            foreach(ProcessThread item in Process.GetProcessById((int)pid).Threads) {
                if(WaitForSingleObject(verified,0)!=258) throw new Exception("任务 Word 已退出");
                IntPtr thread=OpenThread(0x0042,false,(uint)item.Id);
                if(thread==IntPtr.Zero) continue;
                if(GetProcessIdOfThread(thread)!=pid){CloseHandle(thread);throw new Exception("线程归属不符");}
                uint before=SuspendThread(thread);
                if(before==uint.MaxValue){CloseHandle(thread);throw new Exception("暂停任务线程失败");}
                suspended.Add(thread);
                if(before!=0) throw new Exception("任务线程已被其他操作暂停");
            }
            if(suspended.Count==0) throw new Exception("没有实际暂停任何任务线程");
            return suspended;
        } catch {Resume(suspended);throw;}
    }
    public static void Resume(List<IntPtr> threads) {
        bool failed=false;
        foreach(IntPtr thread in threads){if(ResumeThread(thread)==uint.MaxValue) failed=true;CloseHandle(thread);}
        threads.Clear();if(failed) throw new Exception("恢复任务线程失败");
    }
}
'@
function Wait-FaultRecord {
    param([string]$Path,[string]$Marker,[Diagnostics.Process]$Process)
    $faultDeadline=[DateTime]::UtcNow.AddSeconds(90)
    while([DateTime]::UtcNow -lt $faultDeadline) {
        if(Test-Path -LiteralPath $Path) {
            # 日志仍由工作进程写入，读取端也必须允许写共享，不能用 ReadAllText 的只读共享。
            try {
                $faultReader=[IO.StreamReader]::new([IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite))
                try {$faultText=$faultReader.ReadToEnd()} finally {$faultReader.Dispose()}
            } catch [IO.IOException] {$faultText=''}
            if($faultText.Contains($Marker)) {return $faultText}
        }
        if($Process.HasExited) {throw ('故障介入前任务已经退出：'+$Marker)}
        Start-Sleep -Milliseconds 10
    }
    throw ('没有到达实际故障位置：'+$Marker)
}
try {
    if(-not $SkipBuild) {& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
    $faultBaseline = New-TestFixtures -Native
    $faultExe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
    $faultUnrelated=@(Get-Process -Name WINWORD,MathType -ErrorAction SilentlyContinue|ForEach-Object {[ordered]@{Id=$_.Id;Started=$_.StartTime.ToFileTimeUtc();Name=$_.ProcessName}})
    # 建立一个任务外的哨兵 Word 会话，确认故障期间既有文档及会话不会被误关。
    $faultPriorWord=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
    $faultSentinel=New-Object -ComObject Word.Application;$faultSentinelDoc=$faultSentinel.Documents.Add()
    [uint32]$faultSentinelPid=0;$null=[WordFaultCheck]::GetWindowThreadProcessId([IntPtr]$faultSentinelDoc.ActiveWindow.Hwnd,[ref]$faultSentinelPid)
    $faultSentinelOwned=$faultSentinelPid -ne 0 -and $faultSentinelPid -notin $faultPriorWord
    if(-not $faultSentinelOwned) {throw '拒绝用已有 Word 创建哨兵文档。'}
    $faultSentinel.Visible=$false;$faultSentinel.DisplayAlerts=0
    $faultSentinelDoc.Content.Text='DocxMathTypeBatch unrelated-session sentinel '+$faultIdentity
    $faultSentinelText=$faultSentinelDoc.Content.Text
    $faultOriginalLinks=$faultSentinel.Options.UpdateLinksAtOpen
    foreach($faultCase in @(
        @{Name='worker-forward-disconnect';Gui=$false;Reverse=$false;Pause=$false},
        @{Name='worker-reverse-disconnect';Gui=$false;Reverse=$true;Pause=$false},
        @{Name='gui-forward-disconnect';Gui=$true;Reverse=$false;Pause=$false},
        @{Name='gui-reverse-disconnect';Gui=$true;Reverse=$true;Pause=$false},
        @{Name='gui-forward-block-close';Gui=$true;Reverse=$false;Pause=$true},
        @{Name='gui-reverse-block-close';Gui=$true;Reverse=$true;Pause=$true}
    )) {
        $faultFolder=Join-Path $faultTask $faultCase.Name;New-Item -ItemType Directory -Path $faultFolder|Out-Null
        $faultInput=Join-Path $faultFolder 'input.docx'
        $faultSource=Join-Path $faultBaseline.SampleDirectory $(if($faultCase.Reverse){'forward.docx'}else{'input.docx'})
        Copy-Item -LiteralPath $faultSource -Destination $faultInput
        $faultInputHash=(Get-FileHash -LiteralPath $faultInput -Algorithm SHA256).Hash
        $faultTrace=Join-Path $faultFolder $(if($faultCase.Gui){'gui-worker.log'}else{'worker.log'})
        $faultOut=Join-Path $faultFolder 'output.docx';$faultReport=Join-Path $faultFolder 'report.json'
        $faultStop=Join-Path $faultFolder 'stop.flag'
        if($faultCase.Gui) {$faultArguments=@('--ui-test',$faultInput,$faultFolder,$(if($faultCase.Reverse){'reverse-dollar'}else{'normal'}))}
        else {
            $faultStyle=if($faultCase.Reverse){'bracket'}else{'paragraph'}
            $faultArguments=@($(if($faultCase.Reverse){'--worker-reverse'}else{'--worker-convert'}),$faultInput,$faultOut,$faultReport,$faultStyle,$faultStop)
        }
        $faultEarliest=[DateTime]::UtcNow.ToFileTimeUtc()
        $faultProcess=Start-Process -FilePath $faultExe -ArgumentList @($faultArguments|ForEach-Object {'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $faultFolder $(if($faultCase.Gui){'gui.log'}else{'worker.log'})) -RedirectStandardError (Join-Path $faultFolder 'stderr.log') -PassThru
        $faultWord=[IntPtr]::Zero;$faultThreads=$null;$faultInjected=[DateTime]::UtcNow
        try {
            $faultRecord=Wait-FaultRecord -Path $faultTrace -Marker $(if($faultCase.Pause){'UI_PHASE CONVERTING'}else{'CONVERT_PROGRESS 2/'}) -Process $faultProcess
            $faultWordPid=[uint32]([regex]::Matches($faultRecord,'OWNED_WORD_PID=(\d+)')|Select-Object -Last 1).Groups[1].Value
            $faultStarted=[long]([regex]::Matches($faultRecord,'OWNED_WORD_STARTED=(\d+)')|Select-Object -Last 1).Groups[1].Value
            if($faultWordPid -eq $faultSentinelPid -or $faultWordPid -in @($faultUnrelated.Id)) {throw '故障目标错误，拒绝操作既有会话。'}
            $faultWord=[WordFaultCheck]::VerifiedWord($faultWordPid,$faultStarted,$faultEarliest)
            $faultInjected=[DateTime]::UtcNow
            if($faultCase.Pause) {
                $faultThreads=[WordFaultCheck]::Pause($faultWordPid,$faultWord)
                $faultThreadCount=$faultThreads.Count
                Start-Sleep -Milliseconds 400
                $faultWindow=[WordFaultCheck]::Window([uint32]$faultProcess.Id)
                if($faultWindow -eq [IntPtr]::Zero -or -not [WordFaultCheck]::PostMessage($faultWindow,0x0010,[IntPtr]::Zero,[IntPtr]::Zero)) {throw '未向实际 GUI 发送关闭请求。'}
                # 任务 Word 保持暂停；新取消无需恢复服务，也不得保存部分结果。
                if(-not $faultProcess.WaitForExit(3000)){throw '关闭 GUI 后未立即结束暂停的任务。'}
                foreach($faultThread in $faultThreads){[WordFaultCheck]::CloseHandle($faultThread)|Out-Null};$faultThreads=$null
            } else {[WordFaultCheck]::Kill($faultWord)}
            $faultDoneDeadline=[DateTime]::UtcNow.AddSeconds(40)
            while(-not $faultProcess.WaitForExit(100)) {if([DateTime]::UtcNow -gt $faultDoneDeadline) {throw '任务故障/恢复后未在40秒验收边界内结束。'}}
            $faultElapsed=[math]::Round(([DateTime]::UtcNow-$faultInjected).TotalSeconds,2)
            $faultTraceText=[IO.File]::ReadAllText($faultTrace)
            # 未发布的 .work-* 候选不能当最终输出；单独登记清理情况，不把文件扩展名冒充发布状态。
            $faultProduced=@(Get-ChildItem -LiteralPath $faultFolder -Filter '*.docx'|Where-Object { $_.Name -ne 'input.docx' -and $_.Name -notlike '*.work-*.docx' })
            $faultRemnants=@(Get-ChildItem -LiteralPath $faultFolder -Filter '*.work-*.docx'|Select-Object -ExpandProperty Name)
            # 子进程断连时可能暂时遇到共享锁，由监管进程在子进程退出后兜底；以最终状态为准。
            if($faultRemnants.Count -ne 0 -or $faultTraceText.Contains('TASK_COPY_CLEANUP_FAILED') -or -not $faultTraceText.Contains('TASK_COPY_CLEANUP_PASSED')) {throw '故障退出后仍有未清理的工作副本。'}
            if($faultCase.Pause) {
                if($faultProcess.ExitCode -ne 0 -or $faultProduced.Count -ne 0 -or (Test-Path -LiteralPath $faultReport) -or -not $faultTraceText.Contains('UI_CANCELLED')){throw '关闭取消产生了输出或没有确认清理。'}
                $faultOutcome='CANCELLED'
            } else {
                if($faultProcess.ExitCode -eq 0 -or $faultProduced.Count -ne 0 -or (Test-Path -LiteralPath $faultReport) -or $faultTraceText.Contains('UI_CONVERT_SUMMARY')) {throw '真实服务断连后误发布结果或宣称成功。'}
                $faultOutcome='FAILED_NO_PUBLICATION'
            }
            if((Get-FileHash -LiteralPath $faultInput -Algorithm SHA256).Hash -ne $faultInputHash -or $faultSentinelDoc.Content.Text -cne $faultSentinelText -or $faultSentinel.Documents.Count -ne 1 -or
                $faultSentinel.Options.UpdateLinksAtOpen -ne $faultOriginalLinks) {throw '原件或任务外 Word 会话/选项受到影响。'}
            foreach($faultOld in $faultUnrelated) {if((Get-Process -Id $faultOld.Id -ErrorAction Stop).StartTime.ToFileTimeUtc() -ne $faultOld.Started) {throw '已有 Office/MathType 会话改变。'}}
            $faultEvidence.Cases+=[ordered]@{Name=$faultCase.Name;Outcome=$faultOutcome;WordPid=$faultWordPid;WordStarted=$faultStarted;InjectedAfter=$(if($faultCase.Pause){'CONVERTING'}else{'TWO_ACTUAL_SUCCESSES'});SecondsToFinish=$faultElapsed;InputSHA256=$faultInputHash;OriginalBytesUnchanged=$true;UnrelatedSessionsUnchanged=$true;GuiResponsiveWhileBlocked=$faultCase.Pause;PublicationRefusedWhileBlocked=$faultCase.Pause;ThreadsSuspended=$(if($faultCase.Pause){$faultThreadCount}else{0});UnpublishedRemnants=$faultRemnants}
            Write-Output ('REAL_WORD_FAULT_PASSED: '+$faultCase.Name+' '+$faultOutcome+' '+$faultElapsed+'s')
        } finally {
            if($faultThreads) {[WordFaultCheck]::Resume($faultThreads)}
            if($faultWord -ne [IntPtr]::Zero) {
                if([WordFaultCheck]::WaitForSingleObject($faultWord,0) -eq 258) {[WordFaultCheck]::Kill($faultWord)}
                $null=[WordFaultCheck]::CloseHandle($faultWord)
            }
            if(-not $faultProcess.HasExited) {if(-not $faultProcess.WaitForExit(10000)) {$faultProcess.Kill();$faultProcess.WaitForExit()}}
            $faultProcess.Dispose()
        }
    }
    $faultEvidence.Checks=@{ActualWordProcessDisconnected=$true;BothDirectionsAndGUI=$true;NoUnknownResultPublished=$true;TaskCopiesCleanedAfterDisconnect=$true;OriginalBytesUnchanged=$true;UnrelatedWordAndMathTypeUnchanged=$true;SuspendedComServiceReallyBlocks=$true;CloseRequestsImmediateCancel=$true;NoPartialCopyOnCancel=$true;TaskIdentityIncludesCreationTime=$true;AutomaticHardTimeout=$false;PhysicalMouseKeyboard=$false;OtherMachine=$false}
    $faultEvidence.SourceSHA256=[ordered]@{}
    foreach($faultFile in @('gui.cpp','convert_docx.cpp','docx_scan.h','word_automation.h','mathtype_api.h','task_guard.h','test-faults.ps1')) {$faultEvidence.SourceSHA256[$faultFile]=(Get-FileHash -LiteralPath (Get-TestSourcePath $faultFile) -Algorithm SHA256).Hash}
    $faultEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $faultExe -Algorithm SHA256).Hash
    $faultEvidence.Status='PASSED'
    Write-Output 'REAL_WORD_FAULT_ACCEPTANCE_PASSED'
} catch {$faultEvidence.Error=$_.Exception.Message;throw}
finally {
    if($faultSentinelOwned) {$faultSentinelDoc.Close(0);$faultSentinel.Quit(0)}
    if($faultSentinel) {$null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($faultSentinel)}
    # 验收及失败日志均作为本轮证据保存；只移动已核对的本次临时材料。
    $faultResolved=[IO.Path]::GetFullPath($faultTask);$faultPublished=[IO.Path]::GetFullPath($faultPublish)
    $faultSampleRoot=[IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output'))
    if(-not $faultResolved.StartsWith($faultTempRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or
        -not $faultPublished.StartsWith($faultSampleRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {throw '故障验收材料路径越界。'}
    Move-Item -LiteralPath $faultResolved -Destination $faultPublished
    $faultEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'fault-results.json') -Encoding UTF8
    $faultEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $faultPublished 'fault-results.json') -Encoding UTF8
}
