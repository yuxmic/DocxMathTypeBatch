#requires -Version 7.0
# 开发计时入口与正式 GUI 复用原生引擎，不再维护独立的 Python 包写入实现。
param([Parameter(Mandatory)][string]$InputFile,[Parameter(Mandatory)][string]$OutputFile,[ValidateRange(1,256)][double]$Points=14,[switch]$Numbers,[switch]$Progress)
$ErrorActionPreference='Stop'
$project=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$inputPath=(Resolve-Path -LiteralPath $InputFile).Path
$outputPath=[IO.Path]::GetFullPath($OutputFile)
$report=$outputPath+'.ultra.json'
if((Test-Path -LiteralPath $outputPath) -or (Test-Path -LiteralPath $report)){throw '极速输出文件必须不存在。'}
$exe=Join-Path $project 'build/bin/ultra-equation-worker.exe'
$info=[Diagnostics.ProcessStartInfo]::new($exe)
$info.UseShellExecute=$false;$info.CreateNoWindow=$true
$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
foreach($arg in @('--convert',$inputPath,$outputPath,$report,$Points.ToString([Globalization.CultureInfo]::InvariantCulture))){$info.ArgumentList.Add($arg)}
if($Numbers){$info.ArgumentList.Add('--numbers')}
if($Progress){$info.ArgumentList.Add('--progress')}
$timer=[Diagnostics.Stopwatch]::StartNew()
$process=[Diagnostics.Process]::Start($info)
try{
    $stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
    $process.WaitForExit();$logs=$stdout.Result+$stderr.Result
    if($process.ExitCode -ne 0){throw $logs}
    $timer.Stop();$data=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
    # 报告由本次工作进程创建，补充外层启动耗时和逐阶段计时证据。
    $data|Add-Member WallSeconds $timer.Elapsed.TotalSeconds
    $data|Add-Member Workers 1
    $data|Add-Member WorkerLogs $logs
    $data|Add-Member ExecutableSHA256 (Get-FileHash -LiteralPath $exe).Hash
    [IO.File]::WriteAllText($report,($data|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
    $data|ConvertTo-Json -Depth 6
}finally{
    if(-not $process.HasExited){$process.Kill();$process.WaitForExit()}
    $process.Dispose()
}
