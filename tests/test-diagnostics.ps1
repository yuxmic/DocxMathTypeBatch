#requires -Version 7.0
# 真实 GUI 诊断验收：任务清理后仍能查看本份文档的日志，不额外写用户目录。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
$run=Join-Path $ProjectRoot ('tests\output\转换日志-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
$baseline = New-TestFixtures -Native
$evidence=[ordered]@{Status='FAILED';Directory=$run;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@()}
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class DiagnosticUi {
 public delegate bool EnumProc(IntPtr window,IntPtr data);
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr data);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window,StringBuilder text,int max);
 [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr window,int id);
 [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr window,uint message,IntPtr wp,IntPtr lp);
 [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window,uint message,IntPtr wp,IntPtr lp);
 [DllImport("user32.dll",EntryPoint="SendMessageW",CharSet=CharSet.Unicode)] static extern IntPtr ReadText(IntPtr window,uint message,IntPtr wp,StringBuilder text);
 public static IntPtr Dialog(uint pid) {
  IntPtr found=IntPtr.Zero;EnumWindows((window,data)=>{uint owner;GetWindowThreadProcessId(window,out owner);var title=new StringBuilder(256);GetWindowText(window,title,256);if(owner==pid&&title.ToString()=="转换详情"){found=window;return false;}return true;},IntPtr.Zero);return found;
 }
 public static string Text(IntPtr window) {
  int length=SendMessage(window,0xE,IntPtr.Zero,IntPtr.Zero).ToInt32();var text=new StringBuilder(length+1);ReadText(window,0xD,(IntPtr)text.Capacity,text);return text.ToString();
 }
}
'@
try{
 foreach($kind in @('success','invalid','locked','partial')){
  $folder=Join-Path $run $kind;New-Item -ItemType Directory -Path $folder|Out-Null
  $source=Join-Path $folder 'input.docx'
  if($kind -eq 'invalid'){[IO.File]::WriteAllText($source,'not a docx')}
  else{
   $original=if($kind -eq 'partial'){Join-Path $baseline.Directory 'boundaries.docx'}else{Join-Path $baseline.Directory 'mixed.docx'}
   Copy-Item -LiteralPath $original -Destination $source
  }
  $hash=(Get-FileHash -LiteralPath $source).Hash;$lock=$null;$process=$null
  try{
   if($kind -eq 'locked'){$lock=[IO.File]::Open($source,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read)}
   $process=Start-Process -FilePath $exe -ArgumentList @('--ui-batch-test',('"'+$folder+'"'),'diagnostics',('"'+$source+'"')) -WindowStyle Hidden -PassThru
   $deadline=[DateTime]::UtcNow.AddSeconds(180);$checkFile=Join-Path $folder 'batch-check.json'
   while(-not (Test-Path -LiteralPath $checkFile)){
    if($process.HasExited -or [DateTime]::UtcNow -gt $deadline){throw ('日志任务未完成：'+$kind)}
    Start-Sleep -Milliseconds 30
   }
   $window=[GuiBatchCheck]::Find($process.Id)
   # 诊断自检模式完成后保留窗口，等主线程收尾后实际打开详情。
   [DiagnosticUi]::SendMessage($window,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
   $result=Get-Content -LiteralPath $checkFile -Raw|ConvertFrom-Json;$item=$result.Items[0]
   $state=if($kind -eq 'success'){2}elseif($kind -eq 'partial'){3}else{4}
   if($item.State -ne $state -or -not $item.Details.Contains('本次后台日志') -or -not $item.Details.Contains($source)){throw ('未保留对应文档日志：'+$kind)}
   if($kind -eq 'success'){
    if(-not $item.Details.Contains('BATCH_CONVERSION_PASSED') -or $item.Details -notmatch 'ITEM_BEGIN ID=\d+ PARAGRAPH=\d+ START=\d+'){throw '成功日志缺少实际公式位置或完成记录。'}
   }elseif($kind -eq 'partial'){
    if(-not $item.Details.Contains('ITEM_REJECTED') -or -not $item.Details.Contains('REASON=')){throw '部分转换日志缺少保留位置与原因。'}
   }else{
    if(-not $item.Details.Contains('BATCH_CONVERSION_FAILED') -or $item.ExitCode -eq 0 -or -not $item.Stage){throw '失败日志缺少原始原因、阶段或退出码。'}
   }
   # 使用与右键菜单相同的详情入口，核对真实只读滚动框中的文本。
   [DiagnosticUi]::PostMessage($window,0x111,[IntPtr]129,[IntPtr]::Zero)|Out-Null
   $dialog=[IntPtr]::Zero;$deadline=[DateTime]::UtcNow.AddSeconds(10)
   while($dialog -eq [IntPtr]::Zero){if([DateTime]::UtcNow -gt $deadline){throw '转换详情窗口未打开。'};Start-Sleep -Milliseconds 20;$dialog=[DiagnosticUi]::Dialog($process.Id)}
   $visible=[DiagnosticUi]::Text([DiagnosticUi]::GetDlgItem($dialog,301))
   if($visible -ne $item.Details -or [DiagnosticUi]::Text([DiagnosticUi]::GetDlgItem($dialog,302)) -ne '复制全部'){throw '详情窗口内容或复制入口不正确。'}
   if($kind -eq 'invalid'){
    # 截图只展示本次创建的应用窗口，不控制或关闭用户已有的 Office。
    [GuiBatchCheck]::ShowWindow($dialog,4)|Out-Null
    Add-Type -AssemblyName System.Drawing
    $oldDpi=[GuiBatchCheck]::SetThreadDpiAwarenessContext([IntPtr](-4));$rect=[GuiBatchCheck+Rect]::new()
    try{
     if(-not [GuiBatchCheck]::GetWindowRect($dialog,[ref]$rect)){throw '无法读取日志窗口尺寸。'}
     $bitmap=[Drawing.Bitmap]::new($rect.Right-$rect.Left,$rect.Bottom-$rect.Top);$graphics=[Drawing.Graphics]::FromImage($bitmap)
     try{$dc=$graphics.GetHdc();try{if(-not [GuiBatchCheck]::PrintWindow($dialog,$dc,2)){throw '日志窗口截图失败。'}}finally{$graphics.ReleaseHdc($dc)};$bitmap.Save((Join-Path $folder '转换详情.png'))}finally{$graphics.Dispose();$bitmap.Dispose()}
    }finally{[GuiBatchCheck]::SetThreadDpiAwarenessContext($oldDpi)|Out-Null}
   }
   [DiagnosticUi]::PostMessage($dialog,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
   [DiagnosticUi]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
   if(-not $process.WaitForExit(15000) -or $process.ExitCode -ne 0){throw '诊断验收界面没有正常退出。'}
   $cleanup=Get-Content -LiteralPath (Join-Path $folder 'item-0.json') -Raw|ConvertFrom-Json
   if(-not $cleanup.Cleaned -or (Test-Path -LiteralPath $cleanup.TaskDirectory)){throw '日志导致原临时目录未清理。'}
   $evidence.Cases+=@{Name=$kind;State=$item.State;LocationAndReasonRetained=$true;DetailsWindowVerified=$true;TaskCleaned=$true;Directory=$folder}
   Write-Output ('DIAGNOSTICS_CASE_PASSED: '+$kind)
  }finally{
   if($lock){$lock.Dispose()}
   if($process){if(-not $process.HasExited){$window=[GuiBatchCheck]::Find($process.Id);[DiagnosticUi]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null;$null=$process.WaitForExit(15000)};$process.Dispose()}
  }
  if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw '诊断检查修改了原件。'}
 }
 $evidence.Status='PASSED'
}finally{
 $evidence.SourceSHA256=@{};foreach($file in @('gui.cpp','convert_docx.cpp','docx_scan.h','test-diagnostics.ps1','test-gui.ps1')){$evidence.SourceSHA256[$file]=(Get-FileHash -LiteralPath (Get-TestSourcePath $file)).Hash}
 $evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $run 'diagnostics-results.json') -Encoding utf8
 $evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'diagnostics-results.json') -Encoding utf8
}
