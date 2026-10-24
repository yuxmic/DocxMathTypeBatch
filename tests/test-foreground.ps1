#requires -Version 7.0
# 只向自建输入框发送字符和快捷键，检查转换期间输入焦点及主动切换窗口是否保持。
param([switch]$SkipBuild,[string]$BaselineExecutable)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FocusInputCheck {
    [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left,Top,Right,Bottom;}
    [StructLayout(LayoutKind.Sequential)] public struct Gui {public uint Size,Flags;public IntPtr Active,Focus,Capture,MenuOwner,MoveSize,Caret;public Rect CaretRect;}
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint thread,ref Gui info);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint first,uint second,bool attach);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] static extern IntPtr SetFocus(IntPtr window);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll",EntryPoint="GetWindowLongPtrW")] static extern IntPtr GetWindowLongPtr(IntPtr window,int index);
    [DllImport("user32.dll",EntryPoint="SendMessageW",CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr window,uint message,IntPtr wparam,IntPtr lparam);
    public static string Describe(IntPtr window) {
        uint pid;var fg=GetForegroundWindow();GetWindowThreadProcessId(fg,out pid);uint own;uint thread=GetWindowThreadProcessId(window,out own);
        var info=new Gui();info.Size=(uint)Marshal.SizeOf<Gui>();GetGUIThreadInfo(thread,ref info);
        return "foreground="+fg+" pid="+pid+" enabled="+IsWindowEnabled(fg)+" exstyle="+GetWindowLongPtr(fg,-20).ToInt64()+" active="+info.Active+" focus="+info.Focus;
    }
    public static uint ForegroundPid(){uint pid;GetWindowThreadProcessId(GetForegroundWindow(),out pid);return pid;}
    public static bool Focused(IntPtr window,IntPtr edit) {
        if(GetForegroundWindow()!=window)return false;
        uint pid;uint thread=GetWindowThreadProcessId(window,out pid);var info=new Gui();info.Size=(uint)Marshal.SizeOf<Gui>();
        return GetGUIThreadInfo(thread,ref info)&&info.Focus==edit;
    }
    public static bool ActivateTestWindow(IntPtr window,IntPtr edit) {
        // 仅用于测试准备和模拟用户切换；不发送全局按键，立即解除临时输入队列关联。
        uint pid;uint foreground=GetWindowThreadProcessId(GetForegroundWindow(),out pid),current=GetCurrentThreadId();
        bool attached=foreground!=0&&foreground!=current&&AttachThreadInput(current,foreground,true);
        try{SetForegroundWindow(window);SetFocus(edit);}finally{if(attached)AttachThreadInput(current,foreground,false);}
        return Focused(window,edit);
    }
}
'@
$run=Join-Path $TestOutputRoot ('输入焦点-'+[Guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $run|Out-Null
$fixtures=New-TestFixtures -Native
$forms=@();$edits=@()
$evidence=[ordered]@{Status='FAILED';Directory=$run;Cases=@();InputScope='Synthetic WM_CHAR and F6 to owned controls only';PhysicalKeyboard=$false}
try{
    foreach($index in 0..1){
        $form=[Windows.Forms.Form]::new();$form.Text='转换期间输入测试 '+('AB'[$index]);$form.Width=500;$form.Height=240;$form.StartPosition='Manual';$form.Left=80+$index*520;$form.Top=100
        $edit=[Windows.Forms.TextBox]::new();$edit.Multiline=$true;$edit.Dock='Fill';$edit.MaxLength=100000;$form.Controls.Add($edit)
        $form.ActiveControl=$edit
        $edit.Add_KeyDown({if($_.KeyCode -eq [Windows.Forms.Keys]::F6){$script:receivedShortcuts++;$_.Handled=$true}})
        $forms+=$form;$edits+=$edit;$form.Show()
    }
    [Windows.Forms.Application]::DoEvents()
    if(-not [FocusInputCheck]::ActivateTestWindow($forms[0].Handle,$edits[0].Handle)){throw '无法激活自建输入测试窗口，未开始转换。'}
    if(-not [FocusInputCheck]::ActivateTestWindow($forms[1].Handle,$edits[1].Handle) -or [FocusInputCheck]::Focused($forms[0].Handle,$edits[0].Handle)){throw '输入焦点检测的主动切换对照失败。'}
    $evidence.FocusDetectorControl=$true
    $programs=@()
    if($BaselineExecutable){$programs+=@{Name='baseline';Path=[IO.Path]::GetFullPath($BaselineExecutable)}}
    $programs+=@{Name='patched';Path=(Join-Path $ProjectRoot 'build/bin/DocxMathTypeBatch.exe')}
    foreach($program in $programs){foreach($direction in @('forward','reverse')){
        foreach($edit in $edits){$edit.Clear()}
        $script:receivedShortcuts=0;$expected=@([Text.StringBuilder]::new(),[Text.StringBuilder]::new());$samples=0;$lost=0;$shortcuts=0;$active=0;$switched=$false;$wasFocused=$true;$transitions=[Collections.Generic.List[object]]::new();$lossByPid=@{};$firstTextMismatch=$null
        $name=$program.Name+'-'+$direction;$sourceFile=if($direction -eq 'reverse'){$fixtures.Output}else{$fixtures.Input}
        $hash=(Get-FileHash -LiteralPath $sourceFile).Hash;$output=Join-Path $run ($name+'.docx');$report=Join-Path $run ($name+'.json');$log=Join-Path $run ($name+'.log')
        $forms[0].ActiveControl=$edits[0];$edits[0].Select($edits[0].TextLength,0)
        if(-not [FocusInputCheck]::ActivateTestWindow($forms[0].Handle,$edits[0].Handle)){throw '自建输入窗口未获得实际键盘焦点。'}
        $arguments=@($(if($direction -eq 'reverse'){'--worker-reverse'}else{'--worker-convert'}),$sourceFile,$output,$report,$(if($direction -eq 'reverse'){'bracket'}else{'paragraph'}),(Join-Path $run ($name+'.stop')))
        $timer=[Diagnostics.Stopwatch]::StartNew()
        $process=Start-Process -FilePath $program.Path -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -PassThru -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($name+'.errors'))
        try{
            while(-not $process.WaitForExit(10)){
                [Windows.Forms.Application]::DoEvents()
                if(-not $switched -and $timer.ElapsedMilliseconds -ge 1800){
                    # 用户若已主动切到其他应用，不把前台拉回测试窗口。
                    $foregroundOwner=[FocusInputCheck]::ForegroundPid()
                    $stream=[IO.File]::Open($log,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);$reader=[IO.StreamReader]::new($stream)
                    try{$liveLog=$reader.ReadToEnd()}finally{$reader.Dispose()}
                    $taskPids=@([regex]::Matches($liveLog,'OWNED_WORD_PID=(\d+)')|ForEach-Object {[uint32]$_.Groups[1].Value})
                    if($foregroundOwner -eq $PID -or $foregroundOwner -in $taskPids){$active=1;$forms[1].ActiveControl=$edits[1];$edits[1].Select($edits[1].TextLength,0);if(-not [FocusInputCheck]::ActivateTestWindow($forms[1].Handle,$edits[1].Handle)){throw '模拟用户切换输入窗口失败。'}}
                    $switched=$true
                }
                ++$samples
                $focused=[FocusInputCheck]::Focused($forms[$active].Handle,$edits[$active].Handle)
                if($focused -ne $wasFocused){$transitions.Add(@{Milliseconds=$timer.ElapsedMilliseconds;Focused=$focused;State=[FocusInputCheck]::Describe($forms[$active].Handle)});$wasFocused=$focused}
                if($focused){
                    $selection=$edits[$active].SelectionStart;$selectionLength=$edits[$active].SelectionLength
                    $character='abc123中文'[$samples%8];[FocusInputCheck]::SendMessage($edits[$active].Handle,0x102,[IntPtr][int]$character,[IntPtr]1)|Out-Null;$null=$expected[$active].Append($character)
                    if(-not $firstTextMismatch -and $edits[$active].Text -cne $expected[$active].ToString()){$firstTextMismatch=@{Milliseconds=$timer.ElapsedMilliseconds;Editor=$active;Selection=$selection;SelectionLength=$selectionLength;ExpectedLength=$expected[$active].Length;ActualLength=$edits[$active].TextLength}}
                    if($samples%10 -eq 0){++$shortcuts;[FocusInputCheck]::SendMessage($edits[$active].Handle,0x100,[IntPtr]0x75,[IntPtr]1)|Out-Null;[FocusInputCheck]::SendMessage($edits[$active].Handle,0x101,[IntPtr]0x75,[IntPtr]1)|Out-Null}
                }else{++$lost;$owner=[FocusInputCheck]::ForegroundPid();$lossByPid[$owner]=1+$lossByPid[$owner]}
                if(-not $forms[0].Enabled -or -not $forms[1].Enabled){throw '任务外的输入窗口被禁用。'}
            }
            [Windows.Forms.Application]::DoEvents()
            $conversion=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
            $textMatches=$edits[0].Text -ceq $expected[0].ToString() -and $edits[1].Text -ceq $expected[1].ToString()
            $entry=[ordered]@{Program=$program.Name;Direction=$direction;ExecutableSHA256=(Get-FileHash -LiteralPath $program.Path).Hash;ExitCode=$process.ExitCode;Samples=$samples;LostFocusSamples=$lost;Characters=$edits[0].Text.Length+$edits[1].Text.Length;CharactersMatch=$textMatches;ShortcutsExpected=$shortcuts;ShortcutsReceived=$script:receivedShortcuts;UserWindowSwitch=$switched;Status=$conversion.Status;Succeeded=$conversion.Succeeded;InputUnchanged=((Get-FileHash -LiteralPath $sourceFile).Hash -eq $hash)}
            $entry.FocusTransitions=$transitions.ToArray()
            $entry.FirstTextMismatch=$firstTextMismatch
            $ownedPids=@([regex]::Matches([IO.File]::ReadAllText($log),'OWNED_WORD_PID=(\d+)')|ForEach-Object {[uint32]$_.Groups[1].Value}|Select-Object -Unique)
            $wordLoss=0;foreach($owner in $ownedPids){$wordLoss+=$lossByPid[$owner]};$entry.TaskWordFocusSamples=$wordLoss
            $unexpectedLoss=$wordLoss+$lossByPid[[uint32]$PID]+$lossByPid[[uint32]0];$entry.UnexpectedFocusSamples=$unexpectedLoss;$entry.OtherApplicationSamples=$lost-$unexpectedLoss
            $restored=[regex]::Match([IO.File]::ReadAllText($log),'TASK_FOREGROUND_RESTORED=(\d+) FAILED_REQUESTS=(\d+)');$entry.Restored=$(if($restored.Success){[int]$restored.Groups[1].Value}else{0});$entry.FailedRestoreRequests=$(if($restored.Success){[int]$restored.Groups[2].Value}else{0})
            $evidence.Cases+=$entry;$evidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $run 'foreground-results.json') -Encoding utf8
            if($process.ExitCode -ne 0 -or $conversion.Status -ne 'COMPLETE' -or -not $entry.InputUnchanged){throw ('转换核验失败：'+$name)}
            if($program.Name -eq 'patched' -and ($unexpectedLoss -ne 0 -or -not $textMatches -or $shortcuts -ne $script:receivedShortcuts -or -not $switched -or -not [IO.File]::ReadAllText($log).Contains('TASK_FOREGROUND_GUARD_READY'))){throw ('补丁未保持连续输入焦点：'+$name)}
            if($program.Name -eq 'patched' -and ($entry.Characters -lt 16 -or $shortcuts -lt 1)){throw ('输入测试覆盖不足，尚未验证字符与快捷键：'+$name)}
            Write-Output ('FOREGROUND_CASE '+(([pscustomobject]$entry|Select-Object Program,Direction,Samples,LostFocusSamples,TaskWordFocusSamples,CharactersMatch,ShortcutsReceived,UserWindowSwitch,Restored,FailedRestoreRequests)|ConvertTo-Json -Compress))
        }finally{
            if(-not $process.HasExited){[IO.File]::WriteAllText((Join-Path $run ($name+'.stop')),'stop');if(-not $process.WaitForExit(15000)){$process.Kill();$process.WaitForExit()}}
            $process.Dispose()
        }
    }}
    $evidence.Status='PASSED'
}finally{
    foreach($form in $forms){$form.Close();$form.Dispose()}
    $evidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $run 'foreground-results.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'foreground-results.json') -Encoding utf8
}
