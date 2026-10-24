#requires -Version 7.0
# 使用合成文档与独立程序/Skill副本验收连接和自动化，结束后恢复原有本机绑定。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
$run=Join-Path $TestOutputRoot ('Agent-'+[Guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($run)
$exe=Join-Path $ProjectRoot 'build/bin/DocxMathTypeBatch.exe'
$skill=Join-Path $run '独立 Skill'
Copy-Item -LiteralPath (Join-Path $ProjectRoot 'skills/docx-latex-to-mathtype') -Destination $skill -Recurse
. (Join-Path $skill 'scripts/common.ps1')
$shell5=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
$shell7=(Get-Command pwsh).Source
$evidence=[ordered]@{Status='FAILED';ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@();Directory=$run}
$keyPath='Software\DocxMathTypeBatch';$saved=@{};$pauseProcess=$null;$pauseStop=$null
$key=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($keyPath)
try{
    foreach($name in @('AgentExecutable','PointSize')){
        $present=$key -and $name -in $key.GetValueNames()
        $saved[$name]=@{Present=[bool]$present;Value=$(if($present){$key.GetValue($name)}else{$null});Kind=$(if($present){$key.GetValueKind($name)}else{$null})}
    }
}finally{if($key){$key.Dispose()}}
# 快照含本机设置，仅留在忽略的测试目录中，供测试异常退出后人工恢复。
$saved|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $run 'local-settings-backup.json') -Encoding utf8

function Assert-Agent($Condition,[string]$Message){if(-not $Condition){throw $Message}}
function Start-AgentProcess([string]$Name,[string]$Program,[string[]]$Arguments){
    Start-Process -FilePath $Program -ArgumentList @($Arguments|ForEach-Object{ConvertTo-ConverterArgument $_}) -WorkingDirectory $run -WindowStyle Hidden -RedirectStandardOutput (Join-Path $run ($Name+'.jsonl')) -RedirectStandardError (Join-Path $run ($Name+'.errors')) -PassThru
}
function Invoke-AgentScript([string]$Name,[string]$Shell,[string]$Script,[string[]]$Arguments,[int]$ExpectedExit=0){
    $process=Start-AgentProcess $Name $Shell (@('-NoProfile','-File',(Join-Path $skill ('scripts/'+$Script)))+$Arguments)
    try{
        Assert-Agent ($process.WaitForExit(180000)) ('脚本测试超时：'+$Name)
        Assert-Agent ($process.ExitCode -eq $ExpectedExit) ($Name+' 退出码错误；查看本轮 jsonl/errors。')
    }finally{$process.Dispose()}
    $lines=@(Get-Content -LiteralPath (Join-Path $run ($Name+'.jsonl'))|Where-Object{$_})
    $events=@($lines|ForEach-Object{$_|ConvertFrom-Json})
    Assert-Agent ($events.Count -gt 0) ('缺少 JSON 事件：'+$Name)
    $result=$events[-1];$evidence.Cases+=@{Name=$Name;Status=$result.Status;ErrorCode=$result.ErrorCode;ExitCode=$ExpectedExit}
    Write-Host ('AGENT_CASE_PASSED: '+$Name)
    return $result
}
function Open-AgentGui([string]$Path){
    $process=Start-Process -FilePath $Path -WindowStyle Hidden -PassThru
    try{
        $clock=[Diagnostics.Stopwatch]::StartNew();$window=[IntPtr]::Zero
        while($clock.Elapsed.TotalSeconds -lt 15 -and $window -eq [IntPtr]::Zero){$window=[GuiBatchCheck]::Find($process.Id);Start-Sleep -Milliseconds 100}
        Assert-Agent ($window -ne [IntPtr]::Zero) '未找到本次 GUI 窗口'
        # 登记在窗口创建完成后执行，等待它写入当前路径。
        while($clock.Elapsed.TotalSeconds -lt 15 -and (Get-RegisteredConverter) -ne $Path){Start-Sleep -Milliseconds 100}
        Assert-Agent ((Get-RegisteredConverter) -eq $Path) '正常 GUI 未登记自身位置'
        [void][GuiBatchCheck]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)
        Assert-Agent ($process.WaitForExit(10000)) '本次 GUI 未正常关闭'
    }finally{if(-not $process.HasExited){[void][GuiBatchCheck]::PostMessage([GuiBatchCheck]::Find($process.Id),0x10,[IntPtr]::Zero,[IntPtr]::Zero);[void]$process.WaitForExit(10000)};$process.Dispose()}
}

try{
    $wordBefore=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
    foreach($entry in @(@('describe',@('--describe','--json'),0),@('help',@('--help'),0),@('unknown',@('--unknown-agent-command'),2),@('bad-describe',@('--describe'),2))){
        $process=Start-AgentProcess $entry[0] $exe $entry[1]
        try{Assert-Agent ($process.WaitForExit(5000) -and $process.ExitCode -eq $entry[2]) '轻量入口失败或弹窗阻塞'}finally{$process.Dispose()}
    }
    $description=Get-Content -LiteralPath (Join-Path $run 'describe.jsonl') -Raw|ConvertFrom-Json
    Assert-Agent ($description.Version -eq '0.1.2' -and $description.CliContractVersion -eq 1 -and $description.BodyOmmlDefault) '程序版本或接口能力不符'
    $wordAfter=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
    Assert-Agent (-not @($wordAfter|Where-Object{$_ -notin $wordBefore}).Count) '能力查询启动了 Word'
    $evidence.Cases+=@{Name='native-query-help-and-errors';Status='PASSED';WordNotStarted=$true}

    # Windows 自己解析命令行，独立检查空参数、引号、中文与尾反斜杠的引用规则。
    Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class AgentArgumentCheck {
    [DllImport("shell32.dll", CharSet=CharSet.Unicode)] static extern IntPtr CommandLineToArgvW(string text,out int count);
    [DllImport("kernel32.dll")] static extern IntPtr LocalFree(IntPtr memory);
    public static string[] Split(string text) {int count;var p=CommandLineToArgvW(text,out count);try{var values=new string[count];for(int i=0;i<count;i++)values[i]=Marshal.PtrToStringUni(Marshal.ReadIntPtr(p,i*IntPtr.Size));return values;}finally{LocalFree(p);}}
}
'@
    $values=@('', '尾部\', '中文 空格 & 单引号''', '字面量 $(不执行)', '编号 "#E1"', '反斜杠\"引号')
    $actual=[AgentArgumentCheck]::Split('test.exe '+(@($values|ForEach-Object{ConvertTo-ConverterArgument $_}) -join ' '))
    Assert-Agent (-not (Compare-Object $values @($actual|Select-Object -Skip 1))) 'Windows 参数引用未保持原值'
    $evidence.Cases+=@{Name='native-argument-roundtrip';Status='PASSED'}
    $key=[Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($keyPath);try{$key.DeleteValue('AgentExecutable',$false)}finally{$key.Dispose()}
    foreach($shell in @($shell5,$shell7)){
        $tag=if($shell -eq $shell5){'ps51'}else{'ps7'}
        $result=Invoke-AgentScript ($tag+'-unbound') $shell 'connect.ps1' @() 1
        Assert-Agent ($result.ErrorCode -eq 'NEEDS_CONNECTION') '未连接状态不明确'
    }
    $portableFolder=Join-Path $run '独立程序 中文 & 空格';[void][IO.Directory]::CreateDirectory($portableFolder)
    $portable=Join-Path $portableFolder '转换器 已重命名.exe';Copy-Item -LiteralPath $exe -Destination $portable
    Open-AgentGui $portable
    foreach($shell in @($shell5,$shell7)){
        $tag=if($shell -eq $shell5){'ps51'}else{'ps7'}
        $result=Invoke-AgentScript ($tag+'-registered') $shell 'connect.ps1' @()
        Assert-Agent ($result.ExecutablePath -eq $portable) 'Skill 未读取 GUI 登记的位置'
    }
    $result=Invoke-AgentScript 'missing-file' $shell5 'connect.ps1' @('-ExePath',(Join-Path $run 'missing.exe')) 1
    Assert-Agent ($result.ErrorCode -eq 'EXECUTABLE_NOT_FOUND' -and (Get-RegisteredConverter) -eq $portable) '失败连接改写了旧绑定'
    $result=Invoke-AgentScript 'wrong-program' $shell7 'connect.ps1' @('-ExePath',(Join-Path $env:WINDIR 'System32/whoami.exe')) 1
    Assert-Agent ($result.ErrorCode -eq 'UNSUPPORTED_EXECUTABLE' -and (Get-RegisteredConverter) -eq $portable) '未拒绝其他产品'
    # 合成查询程序只模拟协议，不访问 Office；验证不兼容和有界查询超时。
    $fakeSource=Join-Path $run 'query-fixture.cs';$fakeExe=Join-Path $run 'query-fixture.exe'
    @'
// 仅用于模拟能力查询边界。
using System; using System.Reflection; using System.Threading;
[assembly: AssemblyProduct("DocxMathTypeBatch")]
class QueryFixture { static void Main() {if(Environment.GetEnvironmentVariable("DOCX_MATH_QUERY_TEST")=="timeout")Thread.Sleep(15000);Console.WriteLine("{\"Application\":\"DocxMathTypeBatch\",\"CliContractVersion\":999,\"ReportSchemaVersion\":1}");} }
'@|Set-Content -LiteralPath $fakeSource -Encoding utf8
    & (Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe') /nologo /target:exe ('/out:'+$fakeExe) $fakeSource
    Assert-Agent ($LASTEXITCODE -eq 0) '查询测试程序编译失败'
    $result=Invoke-AgentScript 'incompatible-contract' $shell7 'connect.ps1' @('-ExePath',$fakeExe) 1
    Assert-Agent ($result.ErrorCode -eq 'INCOMPATIBLE_CONTRACT') '不兼容接口未被拒绝'
    $previousQuery=$env:DOCX_MATH_QUERY_TEST
    try{$env:DOCX_MATH_QUERY_TEST='timeout';$result=Invoke-AgentScript 'query-timeout' $shell5 'connect.ps1' @('-ExePath',$fakeExe) 1}finally{$env:DOCX_MATH_QUERY_TEST=$previousQuery}
    Assert-Agent ($result.ErrorCode -eq 'QUERY_TIMEOUT') '能力查询未有界结束'
    $movedFolder=Join-Path $run '移动后的目录';[void][IO.Directory]::CreateDirectory($movedFolder)
    $moved=Join-Path $movedFolder 'DocxMathTypeBatch.exe'
    foreach($path in @($portable,$moved)){Assert-Agent ([IO.Path]::GetFullPath($path).StartsWith($run+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) '测试移动越界'}
    Move-Item -LiteralPath $portable -Destination $moved
    $result=Invoke-AgentScript 'moved-not-connected' $shell7 'connect.ps1' @() 1
    Assert-Agent ($result.ErrorCode -eq 'EXECUTABLE_NOT_FOUND') '移动后未要求重连'
    Open-AgentGui $moved
    $null=Invoke-AgentScript 'moved-auto-connect' $shell5 'connect.ps1' @()
    $null=Invoke-AgentScript 'explicit-bind' $shell7 'connect.ps1' @('-ExePath',$moved)

    & python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_experimental.py') review-fixture $run
    Assert-Agent ($LASTEXITCODE -eq 0) 'Agent 合成样本创建失败'
    $inputFile=Join-Path $run "输入 中文 ' & 空格.docx";Copy-Item -LiteralPath (Join-Path $run 'review-positions-input.docx') -Destination $inputFile
    $hash=(Get-FileHash -LiteralPath $inputFile).Hash;$forward=Join-Path $run '自动转换.docx';$partial=Join-Path $run '部分完成.docx'
    $result=Invoke-AgentScript 'forward-ps51' $shell5 'convert.ps1' @('-InputPath',$inputFile,'-OutputPath',$forward,'-TaskDirectory',(Join-Path $run 'forward-task'),'-Tables','-RightNumbers','-FontSize','13.5','-NumberStart','7','-Scope','ALL')
    Assert-Agent ($result.Status -eq 'COMPLETE' -and $result.Succeeded -eq 10 -and $result.NumbersAdded -eq 2) 'PowerShell 5.1 转换结果不符'
    $result=Invoke-AgentScript 'reverse-ps7' $shell7 'convert.ps1' @('-InputPath',$forward,'-OutputPath',(Join-Path $run '反向转换.docx'),'-TaskDirectory',(Join-Path $run 'reverse-task'),'-Direction','ToLatex','-Tables','-DisplayDelimiter','Dollar','-Scope','All')
    Assert-Agent ($result.Status -eq 'COMPLETE' -and $result.Succeeded -eq 10) '反向脚本结果不符'
    $result=Invoke-AgentScript 'partial-ps51' $shell5 'convert.ps1' @('-InputPath',(Join-Path $run 'review-preserved-input.docx'),'-OutputPath',$partial,'-TaskDirectory',(Join-Path $run 'partial-task'),'-Tables','-RightNumbers')
    Assert-Agent ($result.Status -eq 'PARTIAL' -and $result.FailedPreserved -eq 1 -and $result.RejectedPreserved -eq 1 -and [IO.File]::Exists($partial)) '脚本未正确交付部分完成'
    Assert-Agent (-not @($result.PreservedItems|Where-Object{$_.psobject.Properties.Name -contains 'Source'}).Count) '默认摘要暴露了完整源码'
    $outputHash=(Get-FileHash -LiteralPath $forward).Hash
    $result=Invoke-AgentScript 'refuse-overwrite' $shell7 'convert.ps1' @('-InputPath',$inputFile,'-OutputPath',$forward) 1
    Assert-Agent ($result.ErrorCode -eq 'OUTPUT_EXISTS_OR_INVALID' -and (Get-FileHash -LiteralPath $forward).Hash -eq $outputHash) '覆盖保护无效'
    $result=Invoke-AgentScript 'reverse-invalid-options' $shell5 'convert.ps1' @('-InputPath',$forward,'-OutputPath',(Join-Path $run '不应生成.docx'),'-Direction','ToLatex','-RightNumbers') 1
    Assert-Agent ($result.ErrorCode -eq 'INVALID_OPTIONS') '反向无效设置未拒绝'

    $pauseTask=Join-Path $run 'cancel-task';$pauseOutput=Join-Path $run 'cancelled.docx';$pauseStop=Join-Path $pauseTask 'stop.flag'
    $previousBoundary=$env:DOCX_MATH_FAULT_BOUNDARY;$previousGate=$env:DOCX_MATH_FAULT_GATE
    try{
        $env:DOCX_MATH_FAULT_BOUNDARY='EXPERIMENTAL_REPLACED';$env:DOCX_MATH_FAULT_GATE=Join-Path $run 'unused-gate.flag'
        $pauseProcess=Start-AgentProcess 'cancel' $shell7 @('-NoProfile','-File',(Join-Path $skill 'scripts/convert.ps1'),'-InputPath',$inputFile,'-OutputPath',$pauseOutput,'-TaskDirectory',$pauseTask,'-Tables')
    }finally{$env:DOCX_MATH_FAULT_BOUNDARY=$previousBoundary;$env:DOCX_MATH_FAULT_GATE=$previousGate}
    $clock=[Diagnostics.Stopwatch]::StartNew();$paused=$false
    while($clock.Elapsed.TotalSeconds -lt 75 -and -not $pauseProcess.HasExited){
        $log=Join-Path $pauseTask 'worker.log'
        if([IO.File]::Exists($log)){
            $stream=[IO.File]::Open($log,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
            $reader=[IO.StreamReader]::new($stream,[Text.Encoding]::UTF8)
            try{$text=$reader.ReadToEnd()}finally{$reader.Dispose()}
            if($text.Contains('FAULT_BOUNDARY=EXPERIMENTAL_REPLACED')){$paused=$true;break}
        }
        Start-Sleep -Milliseconds 100
    }
    Assert-Agent $paused '取消测试未到实际替换边界'
    $result=Invoke-AgentScript 'busy' $shell5 'convert.ps1' @('-InputPath',$inputFile,'-OutputPath',(Join-Path $run 'busy.docx'),'-TaskDirectory',(Join-Path $run 'busy-task')) 1
    Assert-Agent ($result.ErrorCode -eq 'BUSY') '任务忙未被明确返回'
    [IO.File]::WriteAllText($pauseStop,'stop')
    Assert-Agent ($pauseProcess.WaitForExit(30000) -and $pauseProcess.ExitCode -eq 127) '安全停止未完成'
    $cancel=@(Get-Content -LiteralPath (Join-Path $run 'cancel.jsonl')|Where-Object{$_}|ForEach-Object{$_|ConvertFrom-Json})[-1]
    Assert-Agent ($cancel.Status -eq 'CANCELLED' -and -not [IO.File]::Exists($pauseOutput) -and -not [IO.File]::Exists((Join-Path $pauseTask 'result.json'))) '取消后交付了未完成结果'
    Assert-Agent (-not @(Get-ChildItem -LiteralPath $run -Filter 'cancelled.docx.work-*').Count) '取消后遗留工作副本'
    Assert-Agent ((Get-FileHash -LiteralPath $inputFile).Hash -eq $hash) 'Agent 转换改变了输入原件'
    $evidence.Cases+=@{Name='safe-cancel-after-replacement';Status='CANCELLED';OriginalUnchanged=$true;NoOutput=$true}
    $evidence.Status='PASSED'
}finally{
    if($pauseProcess){if(-not $pauseProcess.HasExited){[IO.File]::WriteAllText($pauseStop,'stop');[void]$pauseProcess.WaitForExit(30000)};$pauseProcess.Dispose()}
    $key=[Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($keyPath)
    try{foreach($name in $saved.Keys){$value=$saved[$name];if($value.Present){$key.SetValue($name,$value.Value,$value.Kind)}else{$key.DeleteValue($name,$false)}}}finally{$key.Dispose()}
    $evidence.LocalSettingsRestored=$true
    $evidence.SourceSHA256=[ordered]@{}
    foreach($name in @('src/app_version.h','src/agent_interface.h','src/gui.cpp','src/convert_docx.cpp','assets/app.rc','tests/test-agent.ps1')){$evidence.SourceSHA256[$name]=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot $name)).Hash}
    foreach($file in Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'skills/docx-latex-to-mathtype') -File -Recurse){$relative=[IO.Path]::GetRelativePath($ProjectRoot,$file.FullName).Replace('\','/');$evidence.SourceSHA256[$relative]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $evidence|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $run 'agent-results.json') -Encoding utf8
}
