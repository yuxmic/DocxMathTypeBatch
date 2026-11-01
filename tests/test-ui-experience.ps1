#requires -Version 7.0
# 真实界面体验验收：拖放不启动 Office、后台窗口不显示、不抢焦点、输出无 JSON。
param([switch]$SkipBuild,[switch]$Reverse,[switch]$ShowUi)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
public static class ExperienceCheck {
    public delegate void EventProc(IntPtr hook,uint evt,IntPtr hwnd,int obj,int child,uint thread,uint time);
    public delegate bool EnumProc(IntPtr hwnd,IntPtr data);
    [StructLayout(LayoutKind.Sequential)] public struct Message {public IntPtr hwnd;public uint message;public UIntPtr wparam;public IntPtr lparam;public uint time;public int x,y;public uint privateValue;}
    [DllImport("user32.dll")] static extern IntPtr SetWinEventHook(uint min,uint max,IntPtr dll,EventProc callback,uint pid,uint tid,uint flags);
    [DllImport("user32.dll")] static extern bool UnhookWinEvent(IntPtr hook);
    [DllImport("user32.dll")] static extern int GetMessage(out Message message,IntPtr hwnd,uint min,uint max);
    [DllImport("user32.dll")] static extern bool PostThreadMessage(uint thread,uint message,UIntPtr wparam,IntPtr lparam);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr hwnd,uint flag);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr parent,IntPtr after,string cls,string title);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int length);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd,StringBuilder text,int length);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd,uint message,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd,uint message,IntPtr wp,IntPtr lp);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalAlloc(uint flags,UIntPtr size);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr memory);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr memory);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc,IntPtr data);
    static Thread watcher;static uint watcherId;static EventProc callback;static ManualResetEvent ready=new ManualResetEvent(false);static bool hooksReady;
    static List<string> events=new List<string>();static List<string> rawEvents=new List<string>();static HashSet<IntPtr> visibleBefore=new HashSet<IntPtr>();
    public static string Text(IntPtr hwnd){var text=new StringBuilder(4096);GetWindowText(hwnd,text,text.Capacity);return text.ToString();}
    public static string[] Texts(IntPtr parent){var result=new List<string>();IntPtr child=IntPtr.Zero;while((child=FindWindowEx(parent,child,"Static",null))!=IntPtr.Zero)result.Add(Text(child));return result.ToArray();}
    public static string[] Events {get {lock(events)return events.ToArray();}}
    public static string[] RawEvents {get {lock(events)return rawEvents.ToArray();}}
    public static bool IsForeground(IntPtr hwnd){return GetForegroundWindow()==hwnd;}
    // 只观察顶层 Word/MathType 窗口，不更改任何窗口状态。
    static void Observe(IntPtr hook,uint evt,IntPtr hwnd,int obj,int child,uint thread,uint time){
        // SHOW 事件即使在回调时已被隐藏也保留，避免漏掉极短的闪窗。
        if(hwnd==IntPtr.Zero||obj!=0||child!=0||GetAncestor(hwnd,2)!=hwnd)return;
        uint pid;GetWindowThreadProcessId(hwnd,out pid);string process;
        try{process=Process.GetProcessById((int)pid).ProcessName;}catch(ArgumentException){return;}
        if(process!="WINWORD"&&process!="MathType")return;
        if(evt==0x8002&&visibleBefore.Contains(hwnd))return;
        var cls=new StringBuilder(128);GetClassName(hwnd,cls,128);
        var visible=IsWindowVisible(hwnd);var foreground=GetForegroundWindow()==hwnd;
        var entry=evt+" "+process+" pid="+pid+" hwnd="+hwnd+" class="+cls+" visible="+visible+" foreground="+foreground+" time="+time+" title="+Text(hwnd)+(cls.ToString()=="#32770"?" dialog="+string.Join(" | ",Texts(hwnd)):"");
        lock(events){rawEvents.Add(entry);if(evt==0x8002||visible||foreground)events.Add(entry);}
    }
    public static void Start(){
        EnumWindows((hwnd,data)=>{if(IsWindowVisible(hwnd))visibleBefore.Add(hwnd);return true;},IntPtr.Zero);
        watcher=new Thread(()=>{watcherId=GetCurrentThreadId();callback=Observe;
            var show=SetWinEventHook(0x8002,0x8002,IntPtr.Zero,callback,0,0,0);var focus=SetWinEventHook(3,3,IntPtr.Zero,callback,0,0,0);
            hooksReady=show!=IntPtr.Zero&&focus!=IntPtr.Zero;
            ready.Set();Message message;while(GetMessage(out message,IntPtr.Zero,0,0)>0){}UnhookWinEvent(show);UnhookWinEvent(focus);
        });watcher.IsBackground=true;watcher.Start();ready.WaitOne();
        if(!hooksReady){Stop();throw new Exception("窗口事件监测未成功启动");}
    }
    public static void Stop(){PostThreadMessage(watcherId,0x12,UIntPtr.Zero,IntPtr.Zero);watcher.Join();}
    // WM_DROPFILES 使用窗口自己的拖放入口，文档路径按 Unicode 传递。
    public static void Drop(IntPtr hwnd,string path){
        byte[] text=Encoding.Unicode.GetBytes(path+"\0\0");IntPtr block=GlobalAlloc(0x42,(UIntPtr)(20+text.Length));IntPtr data=GlobalLock(block);
        Marshal.WriteInt32(data,20);Marshal.WriteInt32(data,16,1);Marshal.Copy(text,0,IntPtr.Add(data,20),text.Length);GlobalUnlock(block);
        if(!PostMessage(hwnd,0x233,block,IntPtr.Zero))throw new Exception("Post file drop failed");
    }
}
'@
$experienceTask=Join-Path $ProjectRoot ('tests\output\work\体验-'+[Guid]::NewGuid().ToString('N'))
$experiencePublished=Join-Path $ProjectRoot ('tests\output\UI体验-'+[IO.Path]::GetFileName($experienceTask))
New-Item -ItemType Directory -Path $experienceTask | Out-Null
$experienceProcess=$null
$experiencePassed=$false
$experienceWatching=$false
$experienceOfficeBefore=@(Get-Process WINWORD,MathType -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$experienceTempBefore=@(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'DocxMathTypeBatch-*' | Select-Object -ExpandProperty FullName)
try {
    $experienceBaseline = New-TestFixtures -Native:$Reverse
    $experienceOfficeBefore=@(Get-Process WINWORD,MathType -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $experienceInput=Join-Path $experienceTask 'input.docx';Copy-Item -LiteralPath $(if($Reverse){$experienceBaseline.Output}else{$experienceBaseline.Input}) -Destination $experienceInput
    $experienceHash=(Get-FileHash -LiteralPath $experienceInput).Hash
    [ExperienceCheck]::Start()
    $experienceWatching=$true
    $experienceProcess=Start-Process -FilePath (Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe') -WindowStyle $(if($ShowUi){'Normal'}else{'Hidden'}) -PassThru
    $experienceDeadline=[DateTime]::UtcNow.AddSeconds(15)
    while([GuiBatchCheck]::Find($experienceProcess.Id) -eq [IntPtr]::Zero){
        if([DateTime]::UtcNow -gt $experienceDeadline){throw '界面启动超时。'}
        Start-Sleep -Milliseconds 100;$experienceProcess.Refresh()
    }
    $experienceWindow=[GuiBatchCheck]::Find($experienceProcess.Id)
    if($Reverse){[ExperienceCheck]::SendMessage([GuiBatchCheck]::GetDlgItem($experienceWindow,102),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null}
    [ExperienceCheck]::Drop($experienceWindow,$experienceInput)
    Start-Sleep -Milliseconds 1000
    $experienceOfficeAfter=@(Get-Process WINWORD,MathType -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    if(@($experienceOfficeAfter | Where-Object {$_ -notin $experienceOfficeBefore}).Count){throw '拖入文档启动了 Office。'}
    if(-not (Save-GuiWindow -PidValue $experienceProcess.Id -Path (Join-Path $experienceTask '准备就绪.png'))){throw '准备状态截图失败。'}
    $experienceStart=[ExperienceCheck]::FindWindowEx($experienceWindow,[IntPtr]::Zero,'Button','开始转换（1）')
    if($experienceStart -eq [IntPtr]::Zero){throw '未找到开始按钮。'}
    $experienceGuiForeground=[ExperienceCheck]::IsForeground($experienceWindow)
    [ExperienceCheck]::SendMessage($experienceStart,0xF5,[IntPtr]::Zero,[IntPtr]::Zero) | Out-Null
    $experienceDeadline=[DateTime]::UtcNow.AddSeconds(180)
    do {
        Start-Sleep -Milliseconds 100
        if($experienceProcess.HasExited){throw ('界面进程提前退出，退出码：'+$experienceProcess.ExitCode)}
        if([DateTime]::UtcNow -gt $experienceDeadline){throw '真实转换超时。'}
        $experienceComplete=[ExperienceCheck]::FindWindowEx($experienceWindow,[IntPtr]::Zero,'Static','批次完成')
        $experienceFailed=[ExperienceCheck]::FindWindowEx($experienceWindow,[IntPtr]::Zero,'Static','转换未完成')
        if($experienceFailed -ne [IntPtr]::Zero){
            $experienceMessages=[ExperienceCheck]::Texts($experienceWindow)
            $experienceMessages | Set-Content -LiteralPath (Join-Path $experienceTask 'error.txt') -Encoding utf8
            throw ('真实转换失败：'+($experienceMessages -join ' | '))
        }
    }while($experienceComplete -eq [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
    if(-not (Save-GuiWindow -PidValue $experienceProcess.Id -Path (Join-Path $experienceTask '转换完成.png'))){throw '完成状态截图失败。'}
    [ExperienceCheck]::Stop()
    $experienceWatching=$false
    $experienceEvents=[ExperienceCheck]::Events
    [IO.File]::WriteAllLines((Join-Path $experienceTask 'window-events.log'),[string[]]$experienceEvents)
    $experienceOutput=Join-Path $experienceTask $(if($Reverse){'input_LaTeX.docx'}else{'input_MathType.docx'})
    if(-not (Test-Path -LiteralPath $experienceOutput) -or (Test-Path -LiteralPath ($experienceOutput+'.json'))){throw '输出缺失或生成了旁置 JSON。'}
    if((Get-FileHash -LiteralPath $experienceInput).Hash -ne $experienceHash){throw '输入被改变。'}
    $experienceTempAfter=@(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'DocxMathTypeBatch-*' | Select-Object -ExpandProperty FullName)
    if(@($experienceTempAfter | Where-Object {$_ -notin $experienceTempBefore}).Count){throw '任务报告临时目录未清理。'}
    $experienceShown=@($experienceEvents|Where-Object {$_ -like '32770 *'}).Count
    $experienceForeground=@($experienceEvents|Where-Object {$_ -like '3 *'}).Count
    [ordered]@{Status=$(if($experienceEvents.Count){'FAILED'}else{'PASSED'});Direction=$(if($Reverse){'reverse'}else{'forward'});VisibleLaunch=$ShowUi.IsPresent;GuiForegroundBeforeConversion=$experienceGuiForeground;NoPreScan=$true;NoVisibleOfficeWindows=($experienceShown -eq 0);NoOfficeForegroundEvents=($experienceForeground -eq 0);OfficeShowEvents=$experienceShown;OfficeForegroundEvents=$experienceForeground;RawOfficeEvents=[ExperienceCheck]::RawEvents.Count;NoOutputSidecar=$true;TemporaryReportRemoved=$true;InputBytesUnchanged=$true;Output=$experienceOutput} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $experienceTask 'experience.json') -Encoding utf8
    if($experienceEvents.Count){throw ('外部窗口显示或抢焦点：'+($experienceEvents -join '; '))}
    $experiencePassed=$true
    Write-Output ('UI_EXPERIENCE_PASSED: '+$experienceTask)
} finally {
    if($experienceProcess -and -not $experienceProcess.HasExited){[ExperienceCheck]::PostMessage([GuiBatchCheck]::Find($experienceProcess.Id),0x10,[IntPtr]::Zero,[IntPtr]::Zero) | Out-Null;$experienceProcess.WaitForExit(15000) | Out-Null}
    if($experienceProcess){$experienceProcess.Dispose()}
    if($experienceWatching){[ExperienceCheck]::Stop()}
    [IO.File]::WriteAllLines((Join-Path $experienceTask 'window-events.log'),[string[]][ExperienceCheck]::Events)
    [IO.File]::WriteAllLines((Join-Path $experienceTask 'raw-window-events.log'),[string[]][ExperienceCheck]::RawEvents)
    if(-not $experiencePassed){Write-Output ('UI_EXPERIENCE_EVIDENCE: '+$experienceTask)}else{
        # 把本轮证据移到验证目录，正常任务临时文件已由产品自行清理。
        Move-Item -LiteralPath $experienceTask -Destination $experiencePublished
        $experienceEvidence=Get-Content -LiteralPath (Join-Path $experiencePublished 'experience.json') -Raw|ConvertFrom-Json -AsHashtable
        $experienceEvidence.Output=Join-Path $experiencePublished ([IO.Path]::GetFileName($experienceOutput));$experienceEvidence.Directory=$experiencePublished
        $experienceEvidence.ReadyScreenshot=Join-Path $experiencePublished '准备就绪.png';$experienceEvidence.ResultScreenshot=Join-Path $experiencePublished '转换完成.png'
        $experienceEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe')).Hash
        $experienceEvidence.SourceSHA256=Get-TestSourceHashes
        $experienceEvidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $experiencePublished 'experience.json') -Encoding utf8
        $experienceEvidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'ui-experience-results.json') -Encoding utf8
    }
}
