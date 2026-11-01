#requires -Version 7.0
# 逐阶段真实取消验收：本轮副本、真实停止按钮、阻塞的任务 Word 和发布竞态。
param([switch]$SkipBuild,[string[]]$Case)
$ErrorActionPreference='Stop'
# 先验证场景名，再创建样本或启动构建；未指定时保持全部阶段验收。
$cancelCases=@(
    @{Name='forward-startup';Boundary='WORD_STARTUP'},
    @{Name='forward-scan';Boundary='SCAN_PROGRESS'},
    @{Name='forward-active-ole';Boundary='NATIVE_STYLE_WRITTEN'},
    @{Name='forward-before-save';Boundary='BEFORE_SAVE'},
    @{Name='forward-saved';Boundary='SAVED_DOCUMENT'},
    @{Name='forward-rescan';Boundary='RESCAN_PROGRESS'},
    @{Name='forward-before-publish';Boundary='BEFORE_PUBLISH'},
    @{Name='forward-output-moved';Boundary='AFTER_OUTPUT_MOVE'},
    @{Name='forward-committed';Boundary='AFTER_COMMIT'},
    @{Name='reverse-scan';Boundary='SCAN_PROGRESS';Reverse=$true},
    @{Name='reverse-before-save';Boundary='BEFORE_SAVE';Reverse=$true},
    @{Name='reverse-rescan';Boundary='RESCAN_PROGRESS';Reverse=$true},
    @{Name='reverse-output-moved';Boundary='AFTER_OUTPUT_MOVE';Reverse=$true},
    @{Name='scan-worker';Boundary='SCAN_PROGRESS';Scan=$true},
    @{Name='gui-scan';Boundary='SCAN_PROGRESS';Gui=$true},
    @{Name='gui-reverse-committed';Boundary='AFTER_COMMIT';Gui=$true;Reverse=$true},
    @{Name='gui-blocked-com';Boundary='NATIVE_STYLE_WRITTEN';Gui=$true;Block=$true}
)
if($PSBoundParameters.ContainsKey('Case')){
    $unknown=@($Case|Where-Object {$_ -notin $cancelCases.Name})
    if(-not $Case -or $unknown.Count){throw ('请指定有效的取消场景；不支持：'+($unknown -join ', '))}
    $selectedCases=@($cancelCases|Where-Object {$_.Name -in $Case})
}else{$selectedCases=$cancelCases}
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$cancelRoot=[IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output\work'))
$cancelIdentity=[Guid]::NewGuid().ToString('N')
$cancelTask=Join-Path $cancelRoot ('取消验收-'+$cancelIdentity)
$cancelPublished=Join-Path $ProjectRoot ('tests\output\取消验收-'+$cancelIdentity)
New-Item -ItemType Directory -Path $cancelTask|Out-Null
$cancelExe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
$cancelEvidence=[ordered]@{Status='FAILED';CheckedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');SelectedCases=@($selectedCases.Name);ChecksScope='SelectedCases';Cases=@();Checks=@{};SampleDirectory=$cancelPublished}
$cancelOldBoundary=$env:DOCX_MATH_FAULT_BOUNDARY;$cancelOldGate=$env:DOCX_MATH_FAULT_GATE
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class CancellationWindow {
    public delegate bool EnumProc(IntPtr window,IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr parameter);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window,StringBuilder name,int capacity);
    [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr window,int id);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr window,uint message,IntPtr wparam,IntPtr lparam);
    [DllImport("kernel32.dll")] static extern IntPtr OpenThread(uint access,bool inherit,uint id);
    [DllImport("kernel32.dll")] static extern uint SuspendThread(IntPtr thread);
    [DllImport("kernel32.dll")] static extern uint ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
    public static IntPtr Window(int pid){IntPtr found=IntPtr.Zero;EnumWindows((w,p)=>{uint owner;GetWindowThreadProcessId(w,out owner);var name=new StringBuilder(128);GetClassName(w,name,128);if(owner==pid&&name.ToString()=="DocxMathTypeBatch.B"){found=w;return false;}return true;},IntPtr.Zero);return found;}
    public static void Stop(int pid){var w=Window(pid);var button=GetDlgItem(w,126);if(w==IntPtr.Zero||button==IntPtr.Zero||!PostMessage(w,0x111,new IntPtr(126),button))throw new Exception("未点击本轮 GUI 的安全停止按钮");}
    // 只暂停带创建时间证明的任务 Word；取消后仅关闭线程句柄。
    public static List<IntPtr> Pause(int pid,long started){using(var p=Process.GetProcessById(pid)){if(p.ProcessName!="WINWORD"||p.StartTime.ToFileTimeUtc()!=started)throw new Exception("拒绝不匹配的 Word 身份");var held=new List<IntPtr>();foreach(ProcessThread t in p.Threads){var h=OpenThread(2,false,(uint)t.Id);if(h==IntPtr.Zero||SuspendThread(h)==uint.MaxValue)throw new Exception("无法暂停任务 Word");held.Add(h);}return held;}}
    public static void Release(List<IntPtr> held,bool alive){foreach(var h in held){if(alive)ResumeThread(h);CloseHandle(h);}held.Clear();}
}
'@
function Start-CancelTask {
    param([string[]]$Arguments,[string]$Folder,[string]$Name)
    return Start-Process -FilePath $cancelExe -ArgumentList @($Arguments|ForEach-Object {'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $Folder ($Name+'.log')) -RedirectStandardError (Join-Path $Folder ($Name+'.errors')) -PassThru
}
function Wait-CancelProcess {
    param([Diagnostics.Process]$Process,[int]$Seconds=60)
    $cancelDeadline=[DateTime]::UtcNow.AddSeconds($Seconds)
    while(-not $Process.WaitForExit(100)){if([DateTime]::UtcNow -gt $cancelDeadline){throw '本轮取消任务未在验收期限内退出。'}}
}
try {
    if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target All}
    $cancelBase = New-TestFixtures -Native
    $cancelStyle = 'paragraph'
    foreach($cancelCase in $selectedCases){
        $cancelFolder=Join-Path $cancelTask $cancelCase.Name;New-Item -ItemType Directory -Path $cancelFolder|Out-Null
        $cancelInput=Join-Path $cancelFolder 'input.docx';Copy-Item -LiteralPath $(if($cancelCase.Reverse){$cancelBase.Output}else{$cancelBase.Input}) -Destination $cancelInput
        $cancelInputHash=(Get-FileHash -LiteralPath $cancelInput).Hash
        $cancelOutput=Join-Path $cancelFolder 'output.docx';$cancelReport=Join-Path $cancelFolder 'report.json';$cancelStop=Join-Path $cancelFolder 'stop.flag'
        $cancelOccupied=Join-Path $cancelFolder 'existing.docx';Copy-Item -LiteralPath $cancelInput -Destination $cancelOccupied;$cancelOccupiedHash=(Get-FileHash -LiteralPath $cancelOccupied).Hash
        $env:DOCX_MATH_FAULT_BOUNDARY=$cancelCase.Boundary;$env:DOCX_MATH_FAULT_GATE=Join-Path $cancelFolder 'release.gate'
        if($cancelCase.Gui){$cancelArgs=@('--ui-test',$cancelInput,$cancelFolder,$(if($cancelCase.Reverse){'reverse-bracket'}else{'normal'}))}
        elseif($cancelCase.Scan){$cancelArgs=@('--worker-scan',$cancelInput,$cancelReport,'forward',$cancelStop)}
        else{$cancelArgs=@($(if($cancelCase.Reverse){'--worker-reverse'}else{'--worker-convert'}),$cancelInput,$cancelOutput,$cancelReport,$(if($cancelCase.Reverse){'bracket'}else{$cancelStyle}),$cancelStop)}
        $cancelProcess=Start-CancelTask $cancelArgs $cancelFolder 'worker';$cancelThreads=$null;$cancelWordPid=0
        try {
            $cancelLog=Join-Path $cancelFolder $(if($cancelCase.Gui){'gui-worker.log'}else{'worker.log'})
            $cancelDeadline=[DateTime]::UtcNow.AddSeconds(60)
            do {
                $cancelText=if(Test-Path -LiteralPath $cancelLog){[string](Get-Content -LiteralPath $cancelLog -Raw)}else{''}
                if($cancelProcess.HasExited -or [DateTime]::UtcNow -gt $cancelDeadline){throw ('未到达取消阶段：'+$cancelCase.Name)}
                Start-Sleep -Milliseconds 20
            }until(([string]$cancelText).Contains('FAULT_BOUNDARY='+$cancelCase.Boundary))
            $cancelWordMatch=[regex]::Matches($cancelText,'OWNED_WORD_PID=(\d+)')|Select-Object -Last 1
            $cancelStartedMatch=[regex]::Matches($cancelText,'OWNED_WORD_STARTED=(\d+)')|Select-Object -Last 1
            if($cancelWordMatch){$cancelWordPid=[int]$cancelWordMatch.Groups[1].Value}
            if($cancelCase.Block){
                $cancelThreads=[CancellationWindow]::Pause($cancelWordPid,[long]$cancelStartedMatch.Groups[1].Value)
                [IO.File]::WriteAllText($env:DOCX_MATH_FAULT_GATE,'release');Start-Sleep -Milliseconds 300
            }
            $cancelRequested=[Diagnostics.Stopwatch]::StartNew()
            if($cancelCase.Gui){[CancellationWindow]::Stop($cancelProcess.Id)}else{[IO.File]::WriteAllText($cancelStop,'')}
            if($cancelCase.Gui){
                $cancelCheck=Join-Path $cancelFolder 'gui-check.json'
                while(-not (Test-Path -LiteralPath $cancelCheck)){if($cancelProcess.HasExited -or $cancelRequested.Elapsed.TotalSeconds -gt 3){throw 'GUI 未立即完成取消并清理。'};Start-Sleep -Milliseconds 20}
                $cancelElapsed=$cancelRequested.Elapsed.TotalSeconds
                Wait-CancelProcess $cancelProcess 5
                if((Get-Content -LiteralPath $cancelCheck -Raw|ConvertFrom-Json).Outcome -ne 'CANCELLED'){throw 'GUI 取消状态不符。'}
                if($cancelProcess.ExitCode -ne 0){throw '用户取消被误报为 GUI 错误。'}
            }else{Wait-CancelProcess $cancelProcess 3;$cancelElapsed=$cancelRequested.Elapsed.TotalSeconds;if($cancelProcess.ExitCode -ne 127){throw '后台未返回独立取消状态。'}}
            $cancelText=[IO.File]::ReadAllText($cancelLog)
            $cancelTail=$cancelText.Substring($cancelText.IndexOf('TASK_CANCEL_REQUESTED'))
            if(-not $cancelText.Contains('UI_CANCELLED') -or $cancelTail -match 'UI_PHASE|UI_STAGE_PROGRESS|RECOVERY_BEGIN|BATCH_CONVERSION_PASSED'){throw '取消后仍处理文档或重新启动恢复任务。'}
            if($cancelElapsed -gt 3 -or @(Get-ChildItem -LiteralPath $cancelFolder -Filter '*.docx'|Where-Object Name -notin @('input.docx','existing.docx')).Count -or (Test-Path -LiteralPath $cancelReport) -or @(Get-ChildItem -LiteralPath $cancelFolder -Filter '~$*').Count){throw '取消未立即完成或留下结果/工作副本。'}
            if($cancelWordPid -and (Get-Process -Id $cancelWordPid -ErrorAction SilentlyContinue)){throw '取消后任务 Word 仍在运行。'}
            if((Get-FileHash -LiteralPath $cancelInput).Hash -ne $cancelInputHash -or (Get-FileHash -LiteralPath $cancelOccupied).Hash -ne $cancelOccupiedHash){throw '取消改变了原件或已有文件。'}
            $cancelEvidence.Cases+=@{Name=$cancelCase.Name;Boundary=$cancelCase.Boundary;CancelSeconds=[math]::Round($cancelElapsed,3);Outcome='CANCELLED';NoOutput=$true;NoReport=$true;WordExited=$true;OriginalUnchanged=$true}
            Write-Output ('CANCEL_CASE_PASSED: '+$cancelCase.Name+' '+[math]::Round($cancelElapsed,3)+'s')
        }finally{
            if($cancelThreads){[CancellationWindow]::Release($cancelThreads,[bool](Get-Process -Id $cancelWordPid -ErrorAction SilentlyContinue))}
            if(-not $cancelProcess.HasExited){[IO.File]::WriteAllText($cancelStop,'');$cancelProcess.WaitForExit(5000)|Out-Null}
            $cancelProcess.Dispose()
        }
    }
    # 选中场景取消后，真实 MathType 连接和正常保存核验必须仍可复用。
    $env:DOCX_MATH_FAULT_BOUNDARY=$null;$env:DOCX_MATH_FAULT_GATE=$null
    $cancelNext=Start-CancelTask @('--worker-convert',$cancelBase.Input,(Join-Path $cancelTask 'next.docx'),(Join-Path $cancelTask 'next.json'),$cancelStyle,(Join-Path $cancelTask 'next-stop.flag')) $cancelTask 'next-task'
    try{Wait-CancelProcess $cancelNext;if($cancelNext.ExitCode -ne 0){throw '取消后下一次转换失败。'}}finally{$cancelNext.Dispose()}
    $cancelEvidence.Checks=@{AllStages=($selectedCases.Count -eq $cancelCases.Count);PublicationRollback=(@($selectedCases|Where-Object {$_.Boundary -in @('AFTER_OUTPUT_MOVE','AFTER_COMMIT')}).Count -gt 0);BlockedComImmediate=(@($selectedCases|Where-Object {$_.Block}).Count -gt 0);NoPostCancelProcessing=$true;NoRecoveryWord=$true;NoPartialOutput=$true;OriginalAndExistingFilesUnchanged=$true;NextConversionPassed=$true}
    $cancelEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $cancelExe).Hash
    $cancelEvidence.SourceSHA256=Get-TestSourceHashes
    $cancelEvidence.Status='PASSED';Write-Output 'CANCELLATION_ACCEPTANCE_PASSED'
}finally{
    $env:DOCX_MATH_FAULT_BOUNDARY=$cancelOldBoundary;$env:DOCX_MATH_FAULT_GATE=$cancelOldGate
    $cancelEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $cancelTask 'cancel-results.json') -Encoding utf8
    $cancelResolved=[IO.Path]::GetFullPath($cancelTask);$cancelDestination=[IO.Path]::GetFullPath($cancelPublished)
    if(-not $cancelResolved.StartsWith($cancelRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or -not $cancelDestination.StartsWith([IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output'))+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '取消验收搬迁越界。'}
    Move-Item -LiteralPath $cancelResolved -Destination $cancelDestination
    $cancelEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'cancel-results.json') -Encoding utf8
}
