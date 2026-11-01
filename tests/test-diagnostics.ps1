#requires -Version 7.0
# 真实 GUI 诊断验收：任务清理后仍能查看本份文档的日志，不额外写用户目录。
param([switch]$SkipBuild,[ValidateSet('All','Pages','Success','Invalid','Locked','Partial')][string[]]$Case=@('All'))
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target $(if($Case -contains 'All' -or $Case -contains 'Pages'){'All'}else{'Gui'})}
. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
$run=Join-Path $ProjectRoot ('tests\output\转换日志-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
$selectedCase=@($Case)
$names=@('success','invalid','locked','partial'|Where-Object { $selectedCase -contains 'All' -or $selectedCase -contains $_ })
$pages=$selectedCase -contains 'All' -or $selectedCase -contains 'Pages'
if($pages){$names+=@('pages','pages-reflow')}
$baseline=if(@($names|Where-Object {$_ -in @('success','locked','partial')}).Count){New-TestFixtures}
$helper=Join-Path $PSScriptRoot 'helpers/test_diagnostics.py'
if($pages){& python -B -X utf8 $helper fixture $run;if($LASTEXITCODE -ne 0){throw '分页诊断样本生成失败。'}}
$evidence=[ordered]@{Status='FAILED';Case=$selectedCase;Directory=$run;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@()}
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
 [DllImport("user32.dll",EntryPoint="SendMessageW",CharSet=CharSet.Unicode)] public static extern IntPtr SetText(IntPtr window,uint message,IntPtr wp,string text);
 [DllImport("user32.dll",EntryPoint="SendMessageW")] static extern IntPtr Selection(IntPtr window,uint message,out int start,out int end);
 public static IntPtr Dialog(uint pid) {
  IntPtr found=IntPtr.Zero;EnumWindows((window,data)=>{uint owner;GetWindowThreadProcessId(window,out owner);var title=new StringBuilder(256);GetWindowText(window,title,256);if(owner==pid&&title.ToString()=="转换详情"){found=window;return false;}return true;},IntPtr.Zero);return found;
 }
 public static string Text(IntPtr window) {
  int length=SendMessage(window,0xE,IntPtr.Zero,IntPtr.Zero).ToInt32();var text=new StringBuilder(length+1);ReadText(window,0xD,(IntPtr)text.Capacity,text);return text.ToString();
 }
 public static bool AllSelected(IntPtr window) {int start,end;Selection(window,0xB0,out start,out end);return start==0&&end==Text(window).Length;}
}
'@
function Wait-DiagnosticDialog([int]$ProcessId,[string]$Title){
 $deadline=[DateTime]::UtcNow.AddSeconds(10)
 # 窗口标题先于子控件创建，必须等待初始化完成后再读按钮和正文。
 do{
  $dialog=[GuiBatchCheck]::Dialog($ProcessId,$Title)
  $ready=$dialog -ne [IntPtr]::Zero -and [DiagnosticUi]::GetDlgItem($dialog,1) -ne [IntPtr]::Zero
  if($Title -eq '转换详情'){$ready=$ready -and [DiagnosticUi]::GetDlgItem($dialog,303) -ne [IntPtr]::Zero -and [DiagnosticUi]::GetDlgItem($dialog,302) -ne [IntPtr]::Zero}
  if([DateTime]::UtcNow -gt $deadline){throw ('窗口未就绪：'+$Title)}
  if(-not $ready){Start-Sleep -Milliseconds 20}
 }while(-not $ready)
 return $dialog
}
function Assert-DiagnosticView([IntPtr]$Dialog,[string]$Expected){
 $edit=[DiagnosticUi]::GetDlgItem($Dialog,301)
 if([DiagnosticUi]::Text($edit) -ne $Expected){throw '详情窗口当前视图与对应验收文本不一致。'}
 # 复制按钮必须选择当前视图全文，不能仍选取隐藏的另一份技术日志。
 [DiagnosticUi]::SendMessage($edit,0xB1,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
 [DiagnosticUi]::SendMessage($Dialog,0x111,[IntPtr]302,[IntPtr]::Zero)|Out-Null
 if(-not [DiagnosticUi]::AllSelected($edit)){throw '复制没有选择当前视图的全部内容。'}
}
try{
 foreach($kind in $names){
  $folder=Join-Path $run $kind;New-Item -ItemType Directory -Path $folder|Out-Null
  $source=Join-Path $folder 'input.docx'
  if($kind -eq 'invalid'){[IO.File]::WriteAllText($source,'not a docx')}
  else{
   $original=if($kind -like 'pages*'){Join-Path $run 'pages-input.docx'}elseif($kind -eq 'partial'){Join-Path $baseline.Directory 'boundaries.docx'}else{Join-Path $baseline.Directory 'mixed.docx'}
   Copy-Item -LiteralPath $original -Destination $source
  }
  $hash=(Get-FileHash -LiteralPath $source).Hash;$lock=$null;$process=$null
  $previousRegion=$env:DOCX_MATH_REGION_FAULT;$previousBody=$env:DOCX_MATH_BODY_FAULT
  try{
   $env:DOCX_MATH_REGION_FAULT=$null;$env:DOCX_MATH_BODY_FAULT=$null
   if($kind -eq 'pages-reflow'){
    $fixture=Get-Content -LiteralPath (Join-Path $run 'pages-fixture.json') -Raw|ConvertFrom-Json
    $env:DOCX_MATH_BODY_FAULT=([string]$fixture.FailureParagraph+'|PREPARE')
   }
   if($kind -eq 'locked'){$lock=[IO.File]::Open($source,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read)}
   $process=Start-Process -FilePath $exe -ArgumentList @('--ui-batch-test',('"'+$folder+'"'),'diagnostics',('"'+$source+'"')) -WindowStyle Hidden -PassThru
   if($kind -like 'pages*'){
    # 沿用现有批处理验收入口，暂停自动开始后操作真实高级设置控件。
    $deadline=[DateTime]::UtcNow.AddSeconds(15)
    while(-not (Test-Path -LiteralPath (Join-Path $folder 'gui-ready.json'))){if($process.HasExited -or [DateTime]::UtcNow -gt $deadline){throw '分页诊断界面未就绪。'};Start-Sleep -Milliseconds 20}
    $window=[GuiBatchCheck]::Find($process.Id);[GuiBatchCheck]::KillTimer($window,[UIntPtr]2)|Out-Null
    [GuiBatchCheck]::PostMessage([GuiBatchCheck]::GetDlgItem($window,143),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    $advanced=Wait-DiagnosticDialog $process.Id '高级设置'
    foreach($id in @(241,243)){
     $checkbox=[GuiBatchCheck]::GetDlgItem($advanced,$id)
     if([GuiBatchCheck]::SendMessage($checkbox,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){[GuiBatchCheck]::SendMessage($checkbox,0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null}
    }
    [GuiBatchCheck]::PostMessage($advanced,0x111,[IntPtr]1,[IntPtr]::Zero)|Out-Null
    while([GuiBatchCheck]::Dialog($process.Id,'高级设置') -ne [IntPtr]::Zero){if([DateTime]::UtcNow -gt $deadline){throw '分页诊断高级设置未提交。'};Start-Sleep -Milliseconds 20}
    if($kind -eq 'pages-reflow'){
     [GuiBatchCheck]::SendMessage([GuiBatchCheck]::GetDlgItem($window,136),0x14E,[IntPtr]1,[IntPtr]::Zero)|Out-Null
     [GuiBatchCheck]::SendMessage($window,0x111,[IntPtr](0x10000+136),[GuiBatchCheck]::GetDlgItem($window,136))|Out-Null
     [DiagnosticUi]::SetText([GuiBatchCheck]::GetDlgItem($window,113),0xC,[IntPtr]::Zero,'72')|Out-Null
    }
    [GuiBatchCheck]::PostMessage([GuiBatchCheck]::GetDlgItem($window,125),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
   }
   $deadline=[DateTime]::UtcNow.AddSeconds(180);$checkFile=Join-Path $folder 'batch-check.json'
   while(-not (Test-Path -LiteralPath $checkFile)){
    if($process.HasExited -or [DateTime]::UtcNow -gt $deadline){throw ('日志任务未完成：'+$kind)}
    Start-Sleep -Milliseconds 30
   }
   $window=[GuiBatchCheck]::Find($process.Id)
   # 诊断自检模式完成后保留窗口，等主线程收尾后实际打开详情。
   [DiagnosticUi]::SendMessage($window,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
   $result=Get-Content -LiteralPath $checkFile -Raw|ConvertFrom-Json;$item=$result.Items[0]
   $state=if($kind -eq 'success'){2}elseif($kind -eq 'partial' -or $kind -like 'pages*'){3}else{4}
   if($item.State -ne $state -or -not $item.Details.Contains('本次后台日志') -or -not $item.Details.Contains($source) -or -not $item.Summary){throw ('未保留对应文档摘要和日志：'+$kind)}
   if($kind -eq 'success'){
    if(-not $item.Details.Contains('BATCH_CONVERSION_PASSED') -or $item.Details -notmatch 'ITEM_BEGIN ID=\d+ PARAGRAPH=\d+ START=\d+'){throw '成功日志缺少实际公式位置或完成记录。'}
   }elseif($kind -eq 'partial' -or $kind -like 'pages*'){
    if(-not $item.Details.Contains('ITEM_REJECTED') -or -not $item.Details.Contains('REASON=')){throw '部分转换日志缺少保留位置与原因。'}
   }else{
    if(-not $item.Details.Contains('BATCH_CONVERSION_FAILED') -or $item.ExitCode -eq 0 -or -not $item.Stage){throw '失败日志缺少原始原因、阶段或退出码。'}
   }
   # 使用与右键菜单相同的详情入口，核对真实只读滚动框中的文本。
   [DiagnosticUi]::PostMessage($window,0x111,[IntPtr]129,[IntPtr]::Zero)|Out-Null
   $dialog=Wait-DiagnosticDialog $process.Id '转换详情'
   if([DiagnosticUi]::Text([DiagnosticUi]::GetDlgItem($dialog,303)) -ne '原始日志'){throw '详情默认没有显示人读摘要。'}
   Assert-DiagnosticView $dialog $item.Summary
   [DiagnosticUi]::SendMessage($dialog,0x111,[IntPtr]303,[IntPtr]::Zero)|Out-Null
   if([DiagnosticUi]::Text([DiagnosticUi]::GetDlgItem($dialog,303)) -ne '返回摘要'){throw '原始日志切换入口错误。'}
   Assert-DiagnosticView $dialog $item.Details
   [DiagnosticUi]::SendMessage($dialog,0x111,[IntPtr]303,[IntPtr]::Zero)|Out-Null
   Assert-DiagnosticView $dialog $item.Summary
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
   if($kind -eq 'pages-reflow'){
    # 受监管探针只读打开本轮输入、输出，证明分页确实变化，不从日志自身推断。
    $probe=Join-Path $ProjectRoot 'build/bin/native-api-probe.exe'
    & $probe --diagnostic-pages $source $item.Output (Join-Path $folder 'page-probe.json') *> (Join-Path $folder 'page-probe.log')
    if($LASTEXITCODE -ne 0){throw '独立分页探测失败，请查看 page-probe.log。'}
   }
   if($kind -like 'pages*'){
    & python -B -X utf8 $helper check $folder $(if($kind -eq 'pages-reflow'){'reflow'}else{'normal'})
    if($LASTEXITCODE -ne 0){throw ('分页诊断独立核对失败：'+$kind)}
   }
   $evidence.Cases+=@{Name=$kind;State=$item.State;LocationAndReasonRetained=$true;DetailsWindowVerified=$true;SummaryAndRawSwitchVerified=$true;CopyCurrentViewSelected=$true;FrozenSourcePagesVerified=($kind -like 'pages*');PaginationChangedVerified=($kind -eq 'pages-reflow');TaskCleaned=$true;Directory=$folder}
   Write-Output ('DIAGNOSTICS_CASE_PASSED: '+$kind)
  }finally{
   $env:DOCX_MATH_REGION_FAULT=$previousRegion;$env:DOCX_MATH_BODY_FAULT=$previousBody
   if($lock){$lock.Dispose()}
   if($process){if(-not $process.HasExited){$window=[GuiBatchCheck]::Find($process.Id);[DiagnosticUi]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null;$null=$process.WaitForExit(15000)};$process.Dispose()}
  }
  if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw '诊断检查修改了原件。'}
 }
 $evidence.Status='PASSED'
}finally{
 $evidence.SourceSHA256=Get-TestSourceHashes
 $evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $run 'diagnostics-results.json') -Encoding utf8
 $evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'diagnostics-results.json') -Encoding utf8
}
