#requires -Version 7.0
# 阶段 6.2.2：真实 COM 永久阻塞、原生样式写入中断、下一任务和既有服务保持。
param([switch]$SkipBuild, [string[]]$Cases, [ValidateRange(5,60)][int]$TimeoutSeconds=8)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$timeoutIdentity = [Guid]::NewGuid().ToString('N')
$timeoutTempRoot = [IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output\work'))
$timeoutTask = Join-Path $timeoutTempRoot ('超时专项-' + $timeoutIdentity)
$timeoutPublished = Join-Path $ProjectRoot ('tests\output\超时专项-' + $timeoutIdentity)
New-Item -ItemType Directory -Path $timeoutTask | Out-Null
$timeoutEvidence = [ordered]@{ CheckedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Status='FAILED'; Scope='阶段6.3自动硬超时与隔离原生样式'; Cases=@(); Checks=@{}; SampleDirectory=$timeoutPublished }
$timeoutPrevious = @{}
foreach ($timeoutVariable in @('DOCX_MATH_TIMEOUT_SECONDS','DOCX_MATH_RECOVERY_DIRECTORY','DOCX_MATH_FAULT_BOUNDARY','DOCX_MATH_FAULT_GATE')) { $timeoutPrevious[$timeoutVariable]=[Environment]::GetEnvironmentVariable($timeoutVariable) }
$timeoutSentinel=$null; $timeoutSentinelDoc=$null; $timeoutSentinelOwned=$false
$timeoutSentinelTask=$null;$timeoutSentinelNative=$null;$timeoutSentinelHandle=[IntPtr]::Zero;$timeoutSentinelGate=Join-Path $timeoutTask 'sentinel-release.flag'

# 只对已核对的进程句柄注入故障；MathType 必须有本轮/前置自建 GUI 任务的明确来源证据。
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public static class TimeoutFaultCheck {
    [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr OpenProcess(uint access,bool inherit,uint pid);
    [DllImport("kernel32.dll")] static extern bool GetProcessTimes(IntPtr process,out long created,out long exited,out long kernel,out long user);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode)] static extern bool QueryFullProcessImageName(IntPtr process,uint flags,StringBuilder path,ref uint size);
    [DllImport("kernel32.dll")] public static extern uint WaitForSingleObject(IntPtr handle,uint milliseconds);
    [DllImport("kernel32.dll")] static extern bool TerminateProcess(IntPtr process,uint code);
    [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll")] static extern IntPtr OpenThread(uint access,bool inherit,uint id);
    [DllImport("kernel32.dll")] static extern uint GetProcessIdOfThread(IntPtr thread);
    [DllImport("kernel32.dll")] static extern uint SuspendThread(IntPtr thread);
    [DllImport("kernel32.dll")] static extern uint ResumeThread(IntPtr thread);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window,uint message,IntPtr wparam,IntPtr lparam);
    [DllImport("user32.dll")] public static extern IntPtr SendMessageTimeout(IntPtr window,uint message,IntPtr wparam,IntPtr lparam,uint flags,uint timeout,out IntPtr result);
    public delegate bool EnumProc(IntPtr window,IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr window,EnumProc callback,IntPtr parameter);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr window,uint id,ref Guid iid,[MarshalAs(UnmanagedType.Interface)] out object value);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window,StringBuilder name,int size);
    public static object NativeWord(uint expected) {
        object found=null;
        EnumWindows((window,p)=>{uint owner;GetWindowThreadProcessId(window,out owner);if(owner!=expected)return true;
            EnumChildWindows(window,(child,data)=>{var name=new StringBuilder(80);GetClassName(child,name,80);if(name.ToString()=="_WwG"){
                var iid=new Guid("00020400-0000-0000-C000-000000000046");object value;if(AccessibleObjectFromWindow(child,0xFFFFFFF0,ref iid,out value)==0)found=value;}
                return found==null;},IntPtr.Zero);return found==null;},IntPtr.Zero);
        return found;
    }
    public static IntPtr Window(uint pid) {
        IntPtr found=IntPtr.Zero;
        EnumWindows((window,p)=>{uint owner;GetWindowThreadProcessId(window,out owner);var name=new StringBuilder(256);GetClassName(window,name,256);if(owner==pid&&name.ToString()=="DocxMathTypeBatch.B"){found=window;return false;}return true;},IntPtr.Zero);
        return found;
    }
    public static bool HasVisibleWindows(uint pid) {
        bool visible=false;
        EnumWindows((window,p)=>{uint owner;GetWindowThreadProcessId(window,out owner);var name=new StringBuilder(256);GetClassName(window,name,256);if(owner==pid&&IsWindowVisible(window)&&(name.ToString()=="EQNWINCLASS"||name.ToString()=="#32770")){visible=true;return false;}return true;},IntPtr.Zero);
        return visible;
    }
    public static IntPtr Verified(uint pid,long started,string expected) {
        IntPtr handle=OpenProcess(0x00101001,false,pid);
        if(handle==IntPtr.Zero) throw new Exception("无法打开精确任务进程");
        try {
            long created,exited,kernel,user;uint size=32768;var image=new StringBuilder((int)size);
            if(!GetProcessTimes(handle,out created,out exited,out kernel,out user)||created!=started||
               !QueryFullProcessImageName(handle,0,image,ref size)||!image.ToString().EndsWith("\\"+expected,StringComparison.OrdinalIgnoreCase)||
               WaitForSingleObject(handle,0)!=258) throw new Exception("进程身份不符，拒绝故障注入");
            return handle;
        } catch {CloseHandle(handle);throw;}
    }
    public static void Kill(IntPtr verified) {
        if(WaitForSingleObject(verified,0)==258&&!TerminateProcess(verified,78)) throw new Exception("服务断连失败");
        if(WaitForSingleObject(verified,5000)!=0) throw new Exception("服务未退出");
    }
    public static List<IntPtr> Pause(uint pid,IntPtr verified) {
        var suspended=new List<IntPtr>();
        try {
            foreach(ProcessThread item in Process.GetProcessById((int)pid).Threads) {
                if(WaitForSingleObject(verified,0)!=258) throw new Exception("待暂停进程已退出");
                IntPtr thread=OpenThread(0x0042,false,(uint)item.Id);if(thread==IntPtr.Zero)continue;
                if(GetProcessIdOfThread(thread)!=pid){CloseHandle(thread);throw new Exception("线程归属不符");}
                uint before=SuspendThread(thread);if(before==uint.MaxValue){CloseHandle(thread);throw new Exception("暂停线程失败");}
                suspended.Add(thread);if(before!=0)throw new Exception("线程已有外部暂停");
            }
            if(suspended.Count==0)throw new Exception("没有暂停任何线程");return suspended;
        }catch{Resume(suspended,WaitForSingleObject(verified,0)==258);throw;}
    }
    public static void Resume(List<IntPtr> threads,bool alive) {
        bool failed=false;foreach(IntPtr thread in threads){if(alive&&ResumeThread(thread)==uint.MaxValue)failed=true;CloseHandle(thread);}
        threads.Clear();if(failed)throw new Exception("恢复测试服务线程失败");
    }
}
'@
function Read-TimeoutTrace([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    $timeoutStream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $timeoutReader=[IO.StreamReader]::new($timeoutStream)
    try { return $timeoutReader.ReadToEnd() } finally { $timeoutReader.Dispose() }
}
function Wait-TimeoutMarker([string]$Path,[string]$Marker,[Diagnostics.Process]$Process) {
    $timeoutDeadline=[DateTime]::UtcNow.AddSeconds(90)
    while ([DateTime]::UtcNow -lt $timeoutDeadline) {
        $timeoutText=Read-TimeoutTrace $Path
        if ($timeoutText.Contains($Marker)) { return $timeoutText }
        if ($Process.HasExited) { throw ('故障边界之前任务已退出：'+$Marker+'；'+$timeoutText) }
        Start-Sleep -Milliseconds 10
    }
    throw ('没有到达实际接口边界：'+$Marker)
}
function Start-TimeoutWorker([string]$Program,[string[]]$Arguments,[string]$Folder,[string]$Name) {
    $timeoutArguments=@($Arguments | ForEach-Object { '"'+$_+'"' })
    return Start-Process -FilePath $Program -ArgumentList $timeoutArguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $Folder ($Name+'.log')) -RedirectStandardError (Join-Path $Folder ($Name+'.err'))
}
function Wait-TimeoutFinish([Diagnostics.Process]$Process,[int]$Seconds=60) {
    $timeoutDeadline=[DateTime]::UtcNow.AddSeconds($Seconds)
    while (-not $Process.WaitForExit(100)) { if ([DateTime]::UtcNow -gt $timeoutDeadline) { throw '产品未在独立验收调查上限内退出。' } }
}
try {
    if (-not $SkipBuild) { & (Join-Path $ProjectRoot 'build.ps1') -Target All }
    & (Join-Path $ProjectRoot 'build/bin/native-api-probe.exe') --task-power-self-test
    if($LASTEXITCODE -ne 0){throw '任务电源与有效运行时间自检失败。'}
    # 只接纳本轮样本转换产生的服务身份，不读取历史日志或进程归属。
    $timeoutInitialOffice=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { [ordered]@{Id=$_.Id;Started=$_.StartTime.ToFileTimeUtc()} })
    $timeoutInitialMath=@(Get-Process -Name MathType -ErrorAction SilentlyContinue | ForEach-Object { [ordered]@{Id=$_.Id;Started=$_.StartTime.ToFileTimeUtc()} })
    $timeoutBaseline = New-TestFixtures -Native
    $timeoutExe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
    $timeoutCli=Join-Path $ProjectRoot 'build\bin\docx-math-convert.exe'
    $timeoutScan=Join-Path $ProjectRoot 'build\bin\docx-math-scan.exe'
    $timeoutStyle='paragraph'
    $timeoutOwnedMath=@{}
    foreach ($service in $timeoutBaseline.OwnedMath) { $timeoutOwnedMath[$service.Id]=$service.Started }
    $timeoutBeforeWord=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    # 用已有受监管扫描入口创建哨兵，停在开发边界；不通过 COM 激活未知实例。
    $savedSentinel=@{};foreach($variable in @('DOCX_MATH_TIMEOUT_SECONDS','DOCX_MATH_FAULT_BOUNDARY','DOCX_MATH_FAULT_GATE')){$savedSentinel[$variable]=[Environment]::GetEnvironmentVariable($variable)}
    try{
        $env:DOCX_MATH_TIMEOUT_SECONDS='3600';$env:DOCX_MATH_FAULT_BOUNDARY='WORD_STARTUP';$env:DOCX_MATH_FAULT_GATE=$timeoutSentinelGate
        $timeoutSentinelTask=Start-TimeoutWorker $timeoutExe @('--worker-scan',$timeoutBaseline.Input,(Join-Path $timeoutTask 'sentinel-scan.json'),'forward',(Join-Path $timeoutTask 'sentinel-stop.flag')) $timeoutTask 'sentinel'
    }finally{foreach($variable in $savedSentinel.Keys){[Environment]::SetEnvironmentVariable($variable,$savedSentinel[$variable])}}
    $sentinelTrace=Wait-TimeoutMarker (Join-Path $timeoutTask 'sentinel.log') 'FAULT_BOUNDARY=WORD_STARTUP' $timeoutSentinelTask
    [uint32]$timeoutSentinelPid=[uint32]([regex]::Matches($sentinelTrace,'OWNED_WORD_PID=(\d+)')|Select-Object -Last 1).Groups[1].Value
    $sentinelStarted=[long]([regex]::Matches($sentinelTrace,'OWNED_WORD_STARTED=(\d+)')|Select-Object -Last 1).Groups[1].Value
    $timeoutSentinelHandle=[TimeoutFaultCheck]::Verified($timeoutSentinelPid,$sentinelStarted,'WINWORD.EXE')
    $timeoutSentinelOwned=$timeoutSentinelPid -ne 0 -and $timeoutSentinelPid -notin $timeoutBeforeWord
    if (-not $timeoutSentinelOwned) { throw '拒绝使用已有 Word 创建哨兵。' }
    $sentinelDeadline=[DateTime]::UtcNow.AddSeconds(15)
    do{$timeoutSentinelNative=[TimeoutFaultCheck]::NativeWord($timeoutSentinelPid);if(-not $timeoutSentinelNative){Start-Sleep -Milliseconds 25}}while(-not $timeoutSentinelNative -and [DateTime]::UtcNow -lt $sentinelDeadline)
    if(-not $timeoutSentinelNative){throw '自建哨兵 Word 原生对象未就绪。'}
    [uint32]$nativePid=0;$null=[TimeoutFaultCheck]::GetWindowThreadProcessId([IntPtr]$timeoutSentinelNative.Hwnd,[ref]$nativePid)
    if($nativePid -ne $timeoutSentinelPid){throw '哨兵原生窗口的 PID 不匹配。'}
    $timeoutSentinel=$timeoutSentinelNative.Application;$timeoutSentinelDoc=$timeoutSentinelNative.Document
    $timeoutSentinel.Visible=$false; $timeoutSentinel.DisplayAlerts=0
    $timeoutSentinelDoc.Content.Text='timeout unrelated Word sentinel '+$timeoutIdentity
    $timeoutSentinelText=$timeoutSentinelDoc.Content.Text; $timeoutSentinelLinks=$timeoutSentinel.Options.UpdateLinksAtOpen
    $timeoutAllCases=@(
        @{Name='cli-word-startup-permanent';Startup=$true;Scan=$true;Gui=$false;Reverse=$false;Cli=$true;Pending=$false;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='cli-scan-permanent-word';Scan=$true;Gui=$false;Reverse=$false;Cli=$true;Pending=$false;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='cli-forward-permanent-word';Scan=$false;Gui=$false;Reverse=$false;Cli=$true;Pending=$false;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='worker-reverse-permanent-word';Scan=$false;Gui=$false;Reverse=$true;Cli=$false;Pending=$false;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='gui-forward-permanent-word-close';Scan=$false;Gui=$true;Reverse=$false;Cli=$false;Pending=$false;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='gui-reverse-permanent-word-close';Scan=$false;Gui=$true;Reverse=$true;Cli=$false;Pending=$false;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='gui-scan-permanent-word-close';Scan=$true;Gui=$true;Reverse=$false;Cli=$false;Pending=$false;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='native-style-word-timeout';Scan=$false;Gui=$false;Reverse=$false;Cli=$false;Pending=$true;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='gui-native-style-word-close';Scan=$false;Gui=$true;Reverse=$false;Cli=$false;Pending=$true;Math=$false;Disconnect=$false;Retry=$false},
        @{Name='native-style-word-timeout-retry';Scan=$false;Gui=$false;Reverse=$false;Cli=$false;Pending=$true;Math=$false;Disconnect=$false;Retry=$true}
    )
    if ($Cases) { foreach ($timeoutRequested in $Cases) { if ($timeoutRequested -notin $timeoutAllCases.Name) { throw ('未知案例：'+$timeoutRequested) } } }
    foreach ($timeoutCase in @($timeoutAllCases | Where-Object { -not $Cases -or $_.Name -in $Cases })) {
        $timeoutFolder=Join-Path $timeoutTask $timeoutCase.Name; New-Item -ItemType Directory -Path $timeoutFolder | Out-Null
        $timeoutRecovery=Join-Path $timeoutFolder 'recovery'; New-Item -ItemType Directory -Path $timeoutRecovery | Out-Null
        # 缩短开发验证的限时，同一路径用于默认120秒；必要时可给冷启动留更长时间。
        $env:DOCX_MATH_TIMEOUT_SECONDS=[string]$TimeoutSeconds; $env:DOCX_MATH_RECOVERY_DIRECTORY=$timeoutRecovery
        $env:DOCX_MATH_FAULT_BOUNDARY=if ($timeoutCase.Startup) { 'WORD_STARTUP' } elseif ($timeoutCase.Pending) { 'NATIVE_STYLE_WRITTEN' } else { $null }
        $env:DOCX_MATH_FAULT_GATE=Join-Path $timeoutFolder 'release.flag'
        $timeoutInput=Join-Path $timeoutFolder 'input.docx'
        Copy-Item -LiteralPath (Join-Path $timeoutBaseline.SampleDirectory $(if ($timeoutCase.Reverse) { 'forward.docx' } else { 'input.docx' })) -Destination $timeoutInput
        $timeoutInputHash=(Get-FileHash -LiteralPath $timeoutInput).Hash
        $timeoutOutput=Join-Path $timeoutFolder 'output.docx'; $timeoutReport=Join-Path $timeoutFolder 'output.docx.json'
        $timeoutProgram=$timeoutExe
        if ($timeoutCase.Gui) { $timeoutArgs=@('--ui-test',$timeoutInput,$timeoutFolder,$(if ($timeoutCase.Reverse) { 'reverse-dollar' } else { 'normal' })) }
        elseif ($timeoutCase.Scan) { $timeoutProgram=$timeoutScan; $timeoutArgs=@('--scan',$timeoutInput,$timeoutReport) }
        else {
            $timeoutMode=if ($timeoutCase.Cli) { if ($timeoutCase.Reverse) { '--reverse' } else { '--convert' } } else { if ($timeoutCase.Reverse) { '--worker-reverse' } else { '--worker-convert' } }
            if ($timeoutCase.Cli) { $timeoutProgram=$timeoutCli }
            $timeoutArgs=@($timeoutMode,$timeoutInput,$timeoutOutput,$timeoutReport,$(if ($timeoutCase.Reverse) { 'bracket' } else { $timeoutStyle }))
            if (-not $timeoutCase.Cli) { $timeoutArgs+=Join-Path $timeoutFolder 'stop.flag' }
        }
        $timeoutProcess=Start-TimeoutWorker $timeoutProgram $timeoutArgs $timeoutFolder 'task'
        $timeoutTrace=if ($timeoutCase.Gui) { Join-Path $timeoutFolder 'gui-worker.log' } else { Join-Path $timeoutFolder 'task.log' }
        $timeoutWord=[IntPtr]::Zero; $timeoutService=[IntPtr]::Zero; $timeoutThreads=$null; $timeoutMathThreads=$false
        $timeoutCaseRecord=[ordered]@{Name=$timeoutCase.Name;Status='FAILED';InputSHA256=$timeoutInputHash;TimeoutSeconds=$TimeoutSeconds}
        try {
            $timeoutMarker=if ($timeoutCase.Startup) { 'FAULT_BOUNDARY=WORD_STARTUP' } elseif ($timeoutCase.Pending) { 'FAULT_BOUNDARY=NATIVE_STYLE_WRITTEN' } elseif ($timeoutCase.Scan) { 'OWNED_WORD_NATIVE_BRIDGE_READY' } else { 'CONVERT_PROGRESS 2/' }
            $timeoutText=Wait-TimeoutMarker $timeoutTrace $timeoutMarker $timeoutProcess
            $timeoutWordPid=[uint32]([regex]::Matches($timeoutText,'OWNED_WORD_PID=(\d+)') | Select-Object -Last 1).Groups[1].Value
            $timeoutWordStarted=[long]([regex]::Matches($timeoutText,'OWNED_WORD_STARTED=(\d+)') | Select-Object -Last 1).Groups[1].Value
            $timeoutWord=[TimeoutFaultCheck]::Verified($timeoutWordPid,$timeoutWordStarted,'WINWORD.EXE')
            $timeoutService=$timeoutWord; $timeoutServicePid=$timeoutWordPid
            $timeoutMathMatch=[regex]::Match($timeoutText,'SHARED_MATHTYPE_PID=(\d+) STARTED=(\d+)')
            if ($timeoutMathMatch.Success) {
                $timeoutMathPid=[int]$timeoutMathMatch.Groups[1].Value; $timeoutMathStarted=[long]$timeoutMathMatch.Groups[2].Value
                if ($timeoutMathPid -notin $timeoutInitialMath.Id) { $timeoutOwnedMath[$timeoutMathPid]=$timeoutMathStarted }
            }
            if ($timeoutCase.Math) {
                if (-not $timeoutMathMatch.Success -or -not $timeoutOwnedMath.ContainsKey($timeoutMathPid) -or $timeoutOwnedMath[$timeoutMathPid] -ne $timeoutMathStarted) { throw 'MathType 缺少精确测试创建来源，拒绝注入故障。' }
                if ([TimeoutFaultCheck]::HasVisibleWindows([uint32]$timeoutMathPid)) { throw 'MathType 存在可见编辑器，拒绝注入故障。' }
                $timeoutService=[TimeoutFaultCheck]::Verified([uint32]$timeoutMathPid,$timeoutMathStarted,'MathType.exe'); $timeoutServicePid=[uint32]$timeoutMathPid; $timeoutMathThreads=$true
            }
            $timeoutInjected=[DateTime]::UtcNow
            if ($timeoutCase.Disconnect) { [TimeoutFaultCheck]::Kill($timeoutService) }
            else { $timeoutThreads=[TimeoutFaultCheck]::Pause($timeoutServicePid,$timeoutService) }
            if ($timeoutCase.Pending -or $timeoutCase.Startup) { [IO.File]::WriteAllText($env:DOCX_MATH_FAULT_GATE,'release') }
            if ($timeoutCase.Gui) {
                $timeoutWindow=[TimeoutFaultCheck]::Window([uint32]$timeoutProcess.Id)
                if ($timeoutWindow -eq [IntPtr]::Zero) { throw '未找到实际 GUI。' }
                [IntPtr]$timeoutReply=[IntPtr]::Zero
                if ([TimeoutFaultCheck]::SendMessageTimeout($timeoutWindow,0,[IntPtr]::Zero,[IntPtr]::Zero,2,300,[ref]$timeoutReply) -eq [IntPtr]::Zero) { throw '故障期间 GUI 不响应。' }
                if (-not [TimeoutFaultCheck]::PostMessage($timeoutWindow,0x0010,[IntPtr]::Zero,[IntPtr]::Zero)) { throw '关闭请求失败。' }
                $timeoutCaseRecord.GuiResponsive=$true; $timeoutCaseRecord.CloseRequested=$true
            }
            # 核对第二个转换任务被互斥门禁拒绝，不创建另一组 Office/工作副本。
            if ($timeoutCase.Name -eq 'cli-forward-permanent-word') {
                $timeoutBusy=Start-TimeoutWorker $timeoutCli @('--convert',$timeoutInput,(Join-Path $timeoutFolder 'busy.docx'),(Join-Path $timeoutFolder 'busy.json'),$timeoutStyle) $timeoutFolder 'busy'
                try { Wait-TimeoutFinish $timeoutBusy 15; if ($timeoutBusy.ExitCode -eq 0 -or -not (Read-TimeoutTrace (Join-Path $timeoutFolder 'busy.log')).Contains('4D415448545950455F5441534B5F42555359')) { throw '并发转换门禁未生效。' } } finally { $timeoutBusy.Dispose() }
                $timeoutCaseRecord.ConcurrentTaskRefused=$true
            }
            Wait-TimeoutFinish $timeoutProcess
            $timeoutCaseRecord.SecondsToFinish=[math]::Round(([DateTime]::UtcNow-$timeoutInjected).TotalSeconds,2)
            $timeoutText=Read-TimeoutTrace $timeoutTrace
            $timeoutCaseRecord.ExitCode=$timeoutProcess.ExitCode
            if (($timeoutCase.Gui -and $timeoutProcess.ExitCode -ne 0) -or (-not $timeoutCase.Gui -and $timeoutProcess.ExitCode -eq 0) -or $timeoutText.Contains('UI_CONVERT_SUMMARY') -or @((Get-ChildItem -LiteralPath $timeoutFolder -Filter '*.docx') | Where-Object Name -ne 'input.docx').Count -ne 0 -or (Test-Path -LiteralPath $timeoutReport)) { throw '故障后误发布结果或留下工作副本。' }
            if (-not $timeoutText.Contains('TASK_COPY_CLEANUP_PASSED') -or $timeoutText.Contains('TASK_COPY_CLEANUP_FAILED') -or [TimeoutFaultCheck]::WaitForSingleObject($timeoutWord,0) -ne 0) { throw '任务 Word 或工作副本没有清理。' }
            if($timeoutCase.Gui){if(-not $timeoutText.Contains('UI_CANCELLED') -or $timeoutCaseRecord.SecondsToFinish -gt 3){throw '关闭 GUI 未立即取消并清理。'}}
            elseif(-not $timeoutCase.Disconnect -and -not $timeoutText.Contains('TASK_HARD_TIMEOUT_PASSED')) { throw '没有产品内部硬超时证据。' }
            $timeoutMarkerPath=Join-Path $timeoutRecovery 'pending-preferences.bin'
            if ($timeoutCase.Pending) {
                if ((Test-Path -LiteralPath $timeoutMarkerPath) -or $timeoutText.Contains('PREFERENCE_RECOVERY_REQUIRED')) { throw '原生样式写入不应遗留共享偏好恢复标记。' }
                $timeoutCaseRecord.NoPendingSharedStyle=$true
            }
            if ($timeoutCase.Retry) {
                $env:DOCX_MATH_FAULT_BOUNDARY=$null; $env:DOCX_MATH_FAULT_GATE=$null
                $timeoutRetry=Start-TimeoutWorker $timeoutCli @('--convert',$timeoutInput,$timeoutOutput,$timeoutReport,$timeoutStyle) $timeoutFolder 'next-task'
                try { Wait-TimeoutFinish $timeoutRetry 90; if ($timeoutRetry.ExitCode -ne 0 -or (Test-Path -LiteralPath $timeoutMarkerPath)) { throw '样式写入中断后下一任务未正常完成。' } } finally { $timeoutRetry.Dispose() }
                $timeoutResult=Get-Content -LiteralPath $timeoutReport -Raw | ConvertFrom-Json
                if ($timeoutResult.Status -ne 'COMPLETE' -or -not $timeoutResult.WordReopenVerified) { throw '下一任务输出未核验。' }
                $timeoutAudit=& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/audit_docx.py') $timeoutOutput
                if ($LASTEXITCODE -ne 0) { throw '下一任务 DOCX 包检查失败。' }
                $timeoutCaseRecord.RestoredNextTask='COMPLETE'
            }
            if ((Get-FileHash -LiteralPath $timeoutInput).Hash -ne $timeoutInputHash -or $timeoutSentinelDoc.Content.Text -cne $timeoutSentinelText -or $timeoutSentinel.Documents.Count -ne 1 -or $timeoutSentinel.Options.UpdateLinksAtOpen -ne $timeoutSentinelLinks) { throw '原件或任务外 Word 哨兵受到影响。' }
            foreach ($timeoutOld in $timeoutInitialOffice) { if ((Get-Process -Id $timeoutOld.Id).StartTime.ToFileTimeUtc() -ne $timeoutOld.Started) { throw '既有 Word 身份改变。' } }
            foreach ($timeoutOld in $timeoutInitialMath) { if (-not $timeoutOwnedMath.ContainsKey([int]$timeoutOld.Id) -and (Get-Process -Id $timeoutOld.Id).StartTime.ToFileTimeUtc() -ne $timeoutOld.Started) { throw '非测试 MathType 身份改变。' } }
            $timeoutCaseRecord.Status='PASSED'; $timeoutCaseRecord.WordPid=$timeoutWordPid; $timeoutCaseRecord.WordStarted=$timeoutWordStarted
            $timeoutCaseRecord.OriginalBytesUnchanged=$true; $timeoutCaseRecord.UnrelatedSessionsUnchanged=$true; $timeoutCaseRecord.NoUnverifiedPublication=$true; $timeoutCaseRecord.OwnedCopiesCleaned=$true
            if ($timeoutCase.Math) { $timeoutCaseRecord.SharedMathTypePid=$timeoutMathPid; $timeoutCaseRecord.SharedMathTypeStarted=$timeoutMathStarted }
            Write-Output ('PRODUCT_TIMEOUT_CASE_PASSED: '+$timeoutCase.Name+' '+$timeoutCaseRecord.SecondsToFinish+'s')
        } catch { $timeoutCaseRecord.Error=$_.Exception.Message; throw }
        finally {
            $timeoutEvidence.Cases+=$timeoutCaseRecord
            if ($timeoutThreads) { [TimeoutFaultCheck]::Resume($timeoutThreads,([TimeoutFaultCheck]::WaitForSingleObject($timeoutService,0) -eq 258)) }
            if ($timeoutWord -ne [IntPtr]::Zero -and [TimeoutFaultCheck]::WaitForSingleObject($timeoutWord,0) -eq 258) { [TimeoutFaultCheck]::Kill($timeoutWord) }
            if (-not $timeoutProcess.HasExited) { if (-not $timeoutProcess.WaitForExit(10000)) { $timeoutProcess.Kill(); $timeoutProcess.WaitForExit() } }
            if ($timeoutMathThreads -and $timeoutService -ne [IntPtr]::Zero) { $null=[TimeoutFaultCheck]::CloseHandle($timeoutService) }
            if ($timeoutWord -ne [IntPtr]::Zero) { $null=[TimeoutFaultCheck]::CloseHandle($timeoutWord) }
            $timeoutProcess.Dispose()
        }
    }
    $timeoutEvidence.Checks=[ordered]@{AutomaticHardTimeoutVerified=$true;PermanentWordBlockVerified=$true;NoUnknownPublication=$true;OwnedWordAndCopiesCleaned=$true;TaskIdentityIncludesCreationTime=$true;UnrelatedSessionsPreserved=$true;SharedMathTypeNeverKilledByProduct=$true;NoPendingSharedStyles=(@($timeoutEvidence.Cases | Where-Object NoPendingSharedStyle -EQ $true).Count -eq 3);NextConversionAfterTimeout=(@($timeoutEvidence.Cases | Where-Object RestoredNextTask -EQ 'COMPLETE').Count -eq 1);PhysicalMouseKeyboard=$false;OtherMachine=$false}
    $timeoutEvidence.FullSuite=-not $Cases
    $timeoutEvidence.SourceSHA256=[ordered]@{}
    foreach ($timeoutSource in @('task_guard.h','gui.cpp','convert_docx.cpp','docx_scan.h','word_automation.h','mathtype_api.h','native_api_probe.cpp','test-timeouts.ps1','build.ps1')) { $timeoutEvidence.SourceSHA256[$timeoutSource]=(Get-FileHash -LiteralPath (Get-TestSourcePath $timeoutSource)).Hash }
    $timeoutEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $timeoutExe).Hash
    $timeoutEvidence.CliExecutableSHA256=(Get-FileHash -LiteralPath $timeoutCli).Hash; $timeoutEvidence.ScanExecutableSHA256=(Get-FileHash -LiteralPath $timeoutScan).Hash
    $timeoutEvidence.Status='PASSED'; Write-Output 'PRODUCT_TIMEOUT_ACCEPTANCE_PASSED'
} catch { $timeoutEvidence.Error=$_.Exception.Message; throw }
finally {
    foreach ($timeoutVariable in $timeoutPrevious.Keys) { [Environment]::SetEnvironmentVariable($timeoutVariable,$timeoutPrevious[$timeoutVariable]) }
    if ($timeoutSentinelDoc) { $null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($timeoutSentinelDoc) }
    if ($timeoutSentinel) { $null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($timeoutSentinel) }
    if ($timeoutSentinelNative) { $null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($timeoutSentinelNative) }
    if($timeoutSentinelTask){
        [IO.File]::WriteAllText($timeoutSentinelGate,'release')
        if(-not $timeoutSentinelTask.WaitForExit(15000)){
            [IO.File]::WriteAllText((Join-Path $timeoutTask 'sentinel-stop.flag'),'stop')
            if(-not $timeoutSentinelTask.WaitForExit(10000)){$timeoutSentinelTask.Kill();$timeoutSentinelTask.WaitForExit()}
        }
        $timeoutSentinelTask.Dispose()
    }
    if($timeoutSentinelHandle -ne [IntPtr]::Zero){$null=[TimeoutFaultCheck]::CloseHandle($timeoutSentinelHandle)}
    # 失败材料同样保存，避免排查日志在 finally 中消失；移动前核对两个绝对路径的边界。
    $timeoutResolved=[IO.Path]::GetFullPath($timeoutTask); $timeoutTarget=[IO.Path]::GetFullPath($timeoutPublished)
    $timeoutSamples=[IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output'))
    if (-not $timeoutResolved.StartsWith($timeoutTempRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or -not $timeoutTarget.StartsWith($timeoutSamples+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw '专项材料路径越界。' }
    Move-Item -LiteralPath $timeoutResolved -Destination $timeoutTarget
    $timeoutEvidence | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $TestOutputRoot 'timeout-results.json') -Encoding utf8
    $timeoutEvidence | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $timeoutTarget 'timeout-results.json') -Encoding utf8
}
