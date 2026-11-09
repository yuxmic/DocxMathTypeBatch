#requires -Version 5.1
# 只读扫描正文公式。完整清单留在本机，标准输出只返回统计和路径。
param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [string]$TaskDirectory,
    [ValidateRange(5,3600)][int]$TimeoutSeconds=120
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'common.ps1')
$exitCode=1
try{
    $inputFile=[IO.Path]::GetFullPath($InputPath)
    if(-not [IO.File]::Exists($inputFile) -or [IO.Path]::GetExtension($inputFile) -ine '.docx'){Throw-ConverterError 'INVALID_INPUT' '输入必须是已保存的 DOCX 文件。'}
    $connection=Get-ConverterConnection
    if($connection.Description.TargetedBodyConversion -ne $true){Throw-ConverterError 'TARGET_SELECTION_UNSUPPORTED' '包内 EXE 不支持逐式选择，请更新完整 Skill 包。'}
    $task=New-ConverterTaskDirectory $TaskDirectory
    $reportPath=Join-Path $task 'scan.json';$stopFile=Join-Path $task 'stop.flag';$logPath=Join-Path $task 'worker.log'
    $arguments=@('--worker-scan',$inputFile,$reportPath,'forward',$stopFile,'--exact-targets')
    $started=[ordered]@{SchemaVersion=1;Event='STARTED';Status='RUNNING';ReadOnly=$true;ReportPath=$reportPath;TaskDirectory=$task;StopFile=$stopFile;LogPath=$logPath}
    Invoke-ConverterWorker $connection.Path $arguments $TimeoutSeconds $started ([ref]$exitCode)
    if($exitCode -ne 0){
        $failure=Get-ConverterWorkerError $exitCode $logPath
        Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='FINISHED';Status=$(if($exitCode -eq 127){'CANCELLED'}else{'FAILED'});ErrorCode=$failure.Code;Reason=$failure.Reason;Stage=$failure.Stage;Message=$failure.Message;ExitCode=$exitCode;ReadOnly=$true;TaskDirectory=$task;LogPath=$logPath})
    }else{
        $report=Read-ConverterScanReport $reportPath 'INVALID_RESULT'
        if($report.Status -cnotin @('SCANNED','SKIPPED_PROTECTED_DOCUMENT') -or $report.Input -ine $inputFile){Throw-ConverterError 'INVALID_RESULT' '扫描结果未满足本次调用约定。'}
        Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='FINISHED';Status=$report.Status;ExitCode=0;ReadOnly=$true;Input=$inputFile;SourceSHA256=$report.SourceSHA256;Ready=$report.Ready;Rejected=$report.Rejected;ReportPath=$reportPath;TaskDirectory=$task;StopFile=$stopFile;LogPath=$logPath})
    }
}catch{Write-ConverterFailure $_;$exitCode=1}
exit $exitCode
