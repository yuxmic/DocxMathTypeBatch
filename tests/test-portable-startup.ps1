#requires -Version 7.0
# 仅将 EXE 复制到空目录启动，防止开发目录中的外置清单掩盖发布依赖。
param([switch]$SkipBuild,[string]$Executable=(Join-Path $PSScriptRoot '../build/bin/DocxMathTypeBatch.exe'))
$ErrorActionPreference='Stop'
$projectRoot=Split-Path -Parent $PSScriptRoot
if(-not $SkipBuild){& (Join-Path $projectRoot 'build.ps1') -Target Gui}
. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class PortableStartupCheck {
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window,StringBuilder text,int capacity);
    public static string Text(IntPtr window){var text=new StringBuilder(1024);GetWindowText(window,text,text.Capacity);return text.ToString();}
}
'@
$folder=Join-Path $projectRoot ('tests/output/portable-startup-'+[Guid]::NewGuid().ToString('N'))
$isolated=Join-Path $folder 'single-exe'
New-Item -ItemType Directory -Path $isolated -Force|Out-Null
$exe=Join-Path $isolated 'DocxMathTypeBatch.exe'
Copy-Item -LiteralPath $Executable -Destination $exe
$evidence=[ordered]@{Status='FAILED';ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;SingleExeOnly=(@(Get-ChildItem -LiteralPath $isolated -Force).Count -eq 1);NoConversion=$true}
$process=$null;$window=[IntPtr]::Zero;$dialog=[IntPtr]::Zero
try{
    $process=Start-Process -FilePath $exe -WorkingDirectory $isolated -WindowStyle Hidden -PassThru
    if(-not $process.WaitForInputIdle(5000)){throw '独立 EXE 未进入窗口消息循环。'}
    $controls=@($process.Modules|Where-Object {$_.ModuleName -ieq 'comctl32.dll'})
    $evidence.CommonControlsV6=@($controls|Where-Object {$_.FileVersionInfo.FileMajorPart -ge 6}).Count -gt 0
    # 首次空闲可能发生在窗口创建期间，等待控件初始化结束或明确报错。
    $deadline=[DateTime]::UtcNow.AddSeconds(5);$columns=0
    do{
        $dialog=[GuiBatchCheck]::Dialog($process.Id,'DocxMathTypeBatch')
        if($dialog -ne [IntPtr]::Zero){throw ('独立 EXE 启动失败：'+[PortableStartupCheck]::Text([GuiBatchCheck]::GetDlgItem($dialog,65535)))}
        $window=[GuiBatchCheck]::Find($process.Id)
        $queue=[GuiBatchCheck]::GetDlgItem($window,138)
        $header=[GuiBatchCheck]::SendMessage($queue,0x101F,[IntPtr]::Zero,[IntPtr]::Zero)
        $columns=[GuiBatchCheck]::SendMessage($header,0x1200,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32()
        if($columns -eq 5 -or $process.HasExited){break}
        Start-Sleep -Milliseconds 100
    }while([DateTime]::UtcNow -lt $deadline)
    $evidence.QueueColumns=$columns
    if($window -eq [IntPtr]::Zero -or $columns -ne 5){throw '独立 EXE 未完成原五列界面初始化。'}
    if(-not $evidence.CommonControlsV6){throw '独立 EXE 未激活 Common-Controls 6。'}
    $evidence.Status='PASSED'
}catch{
    $evidence.Error=$_.Exception.Message
    throw
}finally{
    # 仅关闭本轮启动的进程及其窗口，不影响其他转换或 Office 会话。
    if($process){
        if($dialog -ne [IntPtr]::Zero){[GuiBatchCheck]::PostMessage($dialog,0x111,[IntPtr]1,[IntPtr]::Zero)|Out-Null}
        if($window -ne [IntPtr]::Zero){[GuiBatchCheck]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null}
        $closed=$process.WaitForExit(2000)
        if(-not $closed){$process.Kill();$process.WaitForExit()}
        $evidence.ExitCode=$process.ExitCode
        if($evidence.Status -eq 'PASSED' -and (-not $closed -or $evidence.ExitCode -ne 0)){
            $evidence.Status='FAILED';$evidence.Error='独立 EXE 未正常关闭，或退出码不是 0。'
        }
        $process.Dispose()
    }
    $evidence|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $folder 'startup-result.json') -Encoding utf8
    Write-Output ('STARTUP_EVIDENCE: '+$folder)
}
if($evidence.Status -ne 'PASSED'){throw $evidence.Error}
Write-Output 'PORTABLE_STARTUP_PASSED'
