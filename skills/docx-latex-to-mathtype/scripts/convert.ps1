#requires -Version 5.1
# 转换始终使用已连接的单文件 EXE；完整报告和日志只留在本次任务目录。
param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [ValidateSet('ToMathType','ToLatex')][string]$Direction='ToMathType',
    [ValidateSet('all','inline','display')][string]$Scope='all',
    [Nullable[double]]$FontSize,
    [switch]$RightNumbers,
    [string]$NumberFormat='(#E1)',
    [ValidateRange(1,999999)][int]$NumberStart=1,
    [ValidateRange(1,999999)][int]$NumberChapter=1,
    [ValidateRange(1,999999)][int]$NumberSection=1,
    [switch]$CustomNumber,
    [string]$InlineSpacing='follow',
    [string]$DisplaySpacing='follow',
    [ValidateSet('bracket','dollar')][string]$DisplayDelimiter='bracket',
    [switch]$Tables,[switch]$Textboxes,[switch]$Headers,[switch]$Notes,
    [string]$TaskDirectory,
    [ValidateRange(5,3600)][int]$TimeoutSeconds=120
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'common.ps1')
$process=$null;$processStarted=$false;$logStream=$null;$errorsStream=$null;$stopFile=$null;$exitCode=1
try{
    # PowerShell 的 ValidateSet 接受大小写变体，传给区分大小写的原生参数前统一拼写。
    $Scope=$Scope.ToLowerInvariant();$DisplayDelimiter=$DisplayDelimiter.ToLowerInvariant()
    $inputFile=[IO.Path]::GetFullPath($InputPath);$outputFile=[IO.Path]::GetFullPath($OutputPath)
    if(-not [IO.File]::Exists($inputFile) -or [IO.Path]::GetExtension($inputFile) -ine '.docx'){Throw-ConverterError 'INVALID_INPUT' '输入必须是已保存的 DOCX 文件。'}
    if([IO.Path]::GetExtension($outputFile) -ine '.docx' -or [IO.File]::Exists($outputFile) -or [IO.Directory]::Exists($outputFile)){Throw-ConverterError 'OUTPUT_EXISTS_OR_INVALID' '输出必须是尚未存在的 DOCX 路径。'}
    if(-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($outputFile))){Throw-ConverterError 'OUTPUT_DIRECTORY_MISSING' '请先创建输出文件夹。'}
    $reverse=$Direction -eq 'ToLatex'
    if($reverse -and ($RightNumbers -or $null -ne $FontSize -or $InlineSpacing -ne 'follow' -or $DisplaySpacing -ne 'follow')){Throw-ConverterError 'INVALID_OPTIONS' '反向转换不添加编号，也不设置字号或段落行距。'}
    if($null -ne $FontSize -and (-not ($FontSize -ge 1 -and $FontSize -le 256))){Throw-ConverterError 'INVALID_OPTIONS' '字号范围为 1–256 磅。'}
    $connection=Get-ConverterConnection
    if(-not $TaskDirectory){$TaskDirectory=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('DocxMathTypeBatch/agent-tasks/'+[Guid]::NewGuid().ToString('N'))}
    $task=[IO.Path]::GetFullPath($TaskDirectory)
    if(Test-Path -LiteralPath $task){Throw-ConverterError 'TASK_DIRECTORY_EXISTS' '每次调用必须使用全新的任务目录，以免混用旧报告或停止标记。'}
    [void][IO.Directory]::CreateDirectory($task)
    $reportPath=Join-Path $task 'result.json';$stopFile=Join-Path $task 'stop.flag';$logPath=Join-Path $task 'worker.log';$errorsPath=Join-Path $task 'worker.errors'
    $style=if($reverse){$DisplayDelimiter}elseif($null -ne $FontSize){'size='+$FontSize.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture)}else{'paragraph'}
    $arguments=@($(if($reverse){'--worker-reverse'}else{'--worker-convert'}),$inputFile,$outputFile,$reportPath,$style,$stopFile,('--scope='+$Scope))
    if($RightNumbers){$arguments+=@('--numbers',('--number-format='+$NumberFormat),('--number-start='+$NumberStart),('--number-chapter='+$NumberChapter),('--number-section='+$NumberSection));if($CustomNumber){$arguments+='--number-custom'}}
    if(-not $reverse){$arguments+=@(('--inline-spacing='+$InlineSpacing),('--display-spacing='+$DisplaySpacing))}
    foreach($entry in @(@($Tables,'tables'),@($Textboxes,'textboxes'),@($Headers,'headers'),@($Notes,'notes'))){if($entry[0]){$arguments+='--experimental-'+$entry[1]}}
    $start=[Diagnostics.ProcessStartInfo]::new();$start.FileName=$connection.Path
    $start.Arguments=(@($arguments|ForEach-Object{ConvertTo-ConverterArgument $_}) -join ' ')
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    $start.StandardOutputEncoding=[Text.UTF8Encoding]::new($false);$start.StandardErrorEncoding=[Text.UTF8Encoding]::new($false)
    $start.EnvironmentVariables['DOCX_MATH_TIMEOUT_SECONDS']=$TimeoutSeconds.ToString()
    [void]$start.EnvironmentVariables.Remove('DOCX_MATH_GUI_COMMIT')
    $process=[Diagnostics.Process]::new();$process.StartInfo=$start
    [void]$process.Start();$processStarted=$true
    # 异步拷贝到本机文件，不把整份逐式日志驻留内存或发送给模型。
    $logStream=[IO.FileStream]::new($logPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,[IO.FileOptions]::Asynchronous)
    $errorsStream=[IO.FileStream]::new($errorsPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,[IO.FileOptions]::Asynchronous)
    $logCopy=$process.StandardOutput.BaseStream.CopyToAsync($logStream);$errorsCopy=$process.StandardError.BaseStream.CopyToAsync($errorsStream)
    Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='STARTED';Status='RUNNING';Output=$outputFile;TaskDirectory=$task;StopFile=$stopFile;LogPath=$logPath})
    while(-not $process.WaitForExit(250)){}
    $process.WaitForExit();[void]$logCopy.GetAwaiter().GetResult();[void]$errorsCopy.GetAwaiter().GetResult();$exitCode=$process.ExitCode
    $logStream.Dispose();$logStream=$null;$errorsStream.Dispose();$errorsStream=$null
    if($exitCode -ne 0){
        $failure=Get-ConverterWorkerError $exitCode $logPath
        Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='FINISHED';Status=$(if($exitCode -eq 127){'CANCELLED'}else{'FAILED'});ErrorCode=$failure.Code;Message=$failure.Message;ExitCode=$exitCode;Output=$outputFile;OutputExists=[IO.File]::Exists($outputFile);TaskDirectory=$task;LogPath=$logPath})
    }else{
        if(-not [IO.File]::Exists($reportPath) -or -not [IO.File]::Exists($outputFile)){Throw-ConverterError 'INVALID_RESULT' '程序退出成功，但输出或报告缺失。'}
        $report=[IO.File]::ReadAllText($reportPath,[Text.Encoding]::UTF8)|ConvertFrom-Json
        if($report.SchemaVersion -ne 1 -or $report.CliContractVersion -ne 1 -or $report.Status -notin @('COMPLETE','PARTIAL') -or -not $report.OriginalFileBytesUnchanged -or -not $report.WordReopenVerified){Throw-ConverterError 'INVALID_RESULT' '转换结果未满足调用契约的核验要求。'}
        $preserved=@($report.Items|Where-Object{$_.Status -ne 'SUCCEEDED'}|Select-Object Id,Region,Kind,Status,Reason)
        Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='FINISHED';Status=$report.Status;ExitCode=0;ApplicationVersion=$report.ApplicationVersion;Output=$outputFile;ReportPath=$reportPath;TaskDirectory=$task;StopFile=$stopFile;LogPath=$logPath;Targets=$report.Targets;Succeeded=$report.Succeeded;FailedPreserved=$report.FailedPreserved;RejectedPreserved=$report.RejectedPreserved;NotAttempted=$report.NotAttempted;NumbersAdded=$report.NumbersAdded;NumbersUpdated=$report.NumbersUpdated;NumberReferencesUpdated=$report.NumberReferencesUpdated;ExistingDisplaysFormatted=$report.ExistingDisplaysFormatted;Issues=@($report.Issues);PreservedItems=$preserved;OriginalFileBytesUnchanged=$true;WordReopenVerified=$true;MathematicalEquivalenceVerified=$false})
    }
}catch{Write-ConverterFailure $_;$exitCode=1}
finally{
    if($process){
        if($processStarted -and -not $process.HasExited){
            # 脚本被中断时只请求本任务监管进程清理，不结束其他 Word 或 MathType。
            if($stopFile){[IO.File]::WriteAllText($stopFile,'stop')}
            [void]$process.WaitForExit(30000)
        }
        $process.Dispose()
    }
    if($logStream){$logStream.Dispose()};if($errorsStream){$errorsStream.Dispose()}
}
exit $exitCode
