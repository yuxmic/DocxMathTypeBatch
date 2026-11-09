#requires -Version 5.1
# 通过包内程序的无界面 worker 转换；默认只输出计数和问题短片段。
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
    [switch]$SkipOmml,
    [switch]$FastConversion,
    [string]$TaskDirectory,
    [ValidateRange(5,3600)][int]$TimeoutSeconds=120,
    [switch]$UltraConversion,
    [string]$ScanReportPath,
    [int[]]$TargetId
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'common.ps1')
$exitCode=1;$targetList=$null
try{
    # PowerShell 的 ValidateSet 接受大小写变体，传给区分大小写的原生参数前统一拼写。
    $Scope=$Scope.ToLowerInvariant();$DisplayDelimiter=$DisplayDelimiter.ToLowerInvariant()
    $inputFile=[IO.Path]::GetFullPath($InputPath);$outputFile=[IO.Path]::GetFullPath($OutputPath)
    if(-not [IO.File]::Exists($inputFile) -or [IO.Path]::GetExtension($inputFile) -ine '.docx'){Throw-ConverterError 'INVALID_INPUT' '输入必须是已保存的 DOCX 文件。'}
    if([IO.Path]::GetExtension($outputFile) -ine '.docx' -or [IO.File]::Exists($outputFile) -or [IO.Directory]::Exists($outputFile)){Throw-ConverterError 'OUTPUT_EXISTS_OR_INVALID' '输出必须是尚未存在的 DOCX 路径。'}
    if(-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($outputFile))){Throw-ConverterError 'OUTPUT_DIRECTORY_MISSING' '请先创建输出文件夹。'}
    $reverse=$Direction -eq 'ToLatex'
    $targeted=$PSBoundParameters.ContainsKey('ScanReportPath') -or $PSBoundParameters.ContainsKey('TargetId')
    if($targeted){
        if(-not $ScanReportPath -or -not $TargetId.Count){Throw-ConverterError 'TARGET_SELECTION_INVALID' '必须同时指定扫描清单路径和至少一个公式 ID。'}
        if($reverse -or $FastConversion -or $UltraConversion -or $SkipOmml -or $Scope -ne 'all' -or $Tables -or $Textboxes -or $Headers -or $Notes -or
           $RightNumbers -or $CustomNumber -or $InlineSpacing -ne 'follow' -or $DisplaySpacing -ne 'follow'){
            Throw-ConverterError 'TARGET_SELECTION_UNSUPPORTED' '逐式选择仅支持标准正向正文，不能合用范围筛选、额外位置、右编号或行距设置。'
        }
    }
    if($UltraConversion){
        if($null -eq $FontSize){Throw-ConverterError 'INVALID_OPTIONS' '极速模式必须通过 -FontSize 指定字号。'}
        if($reverse -or $FastConversion -or $SkipOmml -or $Scope -ne 'all' -or $Tables -or $Textboxes -or $Headers -or $Notes -or
           $InlineSpacing -ne 'follow' -or $DisplaySpacing -ne 'follow' -or $CustomNumber -or $NumberFormat -cne '(#E1)' -or
           $NumberStart -ne 1 -or $NumberChapter -ne 1 -or $NumberSection -ne 1){
            Throw-ConverterError 'INVALID_OPTIONS' '极速模式只支持正向转换全部正文 LaTeX 和 Word 原生公式、固定字号及可选默认右编号。'
        }
    }
    if($reverse -and ($RightNumbers -or $null -ne $FontSize -or $InlineSpacing -ne 'follow' -or $DisplaySpacing -ne 'follow')){Throw-ConverterError 'INVALID_OPTIONS' '反向转换不添加编号，也不设置字号或段落行距。'}
    if($null -ne $FontSize -and (-not ($FontSize -ge 1 -and $FontSize -le 256))){Throw-ConverterError 'INVALID_OPTIONS' '字号范围为 1–256 磅。'}
    $connection=Get-ConverterConnection
    if($UltraConversion -and 'Ultra' -notin $connection.Description.ConversionModes){Throw-ConverterError 'UNSUPPORTED_OPTION' '包内 EXE 不支持极速转换，请更新完整 Skill 包。'}
    if($FastConversion -and -not $connection.Description.ExperimentalFastConversion){Throw-ConverterError 'UNSUPPORTED_OPTION' '当前 EXE 不支持实验性快速转换，请更新程序。'}
    if($SkipOmml -and -not $reverse -and -not $connection.Description.OptionalOmmlConversion){Throw-ConverterError 'UNSUPPORTED_OPTION' '当前 EXE 不支持保留 Word 原生公式的选项，请更新程序。'}
    if($targeted){
        if($connection.Description.TargetedBodyConversion -ne $true){Throw-ConverterError 'TARGET_SELECTION_UNSUPPORTED' '包内 EXE 不支持逐式选择，请更新完整 Skill 包。'}
        $scan=Read-ConverterScanReport ([IO.Path]::GetFullPath($ScanReportPath)) 'TARGET_SELECTION_INVALID'
        if($scan.Status -cne 'SCANNED' -or $scan.Input -ine $inputFile){Throw-ConverterError 'TARGET_SELECTION_INVALID' '请使用当前输入文件成功扫描后生成的清单。'}
        $targetLines=@(Get-ConverterTargetLines $scan $TargetId)
    }
    $task=New-ConverterTaskDirectory $TaskDirectory
    $reportPath=Join-Path $task 'result.json';$stopFile=Join-Path $task 'stop.flag';$logPath=Join-Path $task 'worker.log'
    $style=if($reverse){$DisplayDelimiter}elseif($null -ne $FontSize){'size='+$FontSize.ToString('0.##',[Globalization.CultureInfo]::InvariantCulture)}else{'paragraph'}
    $arguments=@($(if($reverse){'--worker-reverse'}else{'--worker-convert'}),$inputFile,$outputFile,$reportPath,$style,$stopFile,('--scope='+$Scope))
    if($RightNumbers){$arguments+=@('--numbers',('--number-format='+$NumberFormat),('--number-start='+$NumberStart),('--number-chapter='+$NumberChapter),('--number-section='+$NumberSection));if($CustomNumber){$arguments+='--number-custom'}}
    if(-not $reverse){$arguments+=@(('--inline-spacing='+$InlineSpacing),('--display-spacing='+$DisplaySpacing))}
    foreach($entry in @(@($Tables,'tables'),@($Textboxes,'textboxes'),@($Headers,'headers'),@($Notes,'notes'))){if($entry[0]){$arguments+='--experimental-'+$entry[1]}}
    if($FastConversion){$arguments+='--experimental-fast'}
    if($UltraConversion){$arguments+='--ultra'}
    if($SkipOmml -and -not $reverse){$arguments+='--skip-omml'}
    if($targeted){
        $targetList=Join-Path $task 'targets.txt'
        [IO.File]::WriteAllLines($targetList,[string[]]$targetLines,[Text.Encoding]::ASCII)
        $arguments+='--target-list-file='+$targetList
    }
    $started=[ordered]@{SchemaVersion=1;Event='STARTED';Status='RUNNING';Output=$outputFile;TaskDirectory=$task;StopFile=$stopFile;LogPath=$logPath}
    Invoke-ConverterWorker $connection.Path $arguments $TimeoutSeconds $started ([ref]$exitCode)
    if($exitCode -ne 0){
        $failure=Get-ConverterWorkerError $exitCode $logPath
        Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='FINISHED';Status=$(if($exitCode -eq 127){'CANCELLED'}else{'FAILED'});ErrorCode=$failure.Code;Reason=$failure.Reason;Stage=$failure.Stage;Message=$failure.Message;ExitCode=$exitCode;FastConversion=[bool]$FastConversion;VerificationMode='NOT_COMPLETED';Output=$outputFile;OutputExists=[IO.File]::Exists($outputFile);TaskDirectory=$task;LogPath=$logPath})
    }else{
        if(-not [IO.File]::Exists($reportPath) -or -not [IO.File]::Exists($outputFile)){Throw-ConverterError 'INVALID_RESULT' '程序退出成功，但输出或报告缺失。'}
        $report=[IO.File]::ReadAllText($reportPath,[Text.Encoding]::UTF8)|ConvertFrom-Json
        # 极速报告不进行 Word 重开，单独核对实际模式和未核验标志。
        $validMode=if($UltraConversion){
            $report.ConversionMode -eq 'Ultra' -and $report.Validation -eq 'NONE' -and $report.WordStarted -eq $false -and $report.WordReopenVerified -eq $false
        }else{
            $report.ConversionMode -ne 'Ultra' -and ($FastConversion -or ($report.OriginalFileBytesUnchanged -and $report.WordReopenVerified))
        }
        $validTargets=-not $targeted -or ($report.TargetedBodyConversion -eq $true -and $report.Targets -eq $TargetId.Count)
        if($report.SchemaVersion -ne 1 -or $report.CliContractVersion -ne 1 -or $report.Status -notin @('COMPLETE','PARTIAL') -or [bool]$report.FastConversion -ne [bool]$FastConversion -or ($SkipOmml -and -not $reverse -and $report.ConvertOmml -ne $false) -or -not $validMode -or -not $validTargets){Throw-ConverterError 'INVALID_RESULT' '转换结果未满足本次模式的调用契约。'}
        $preserved=@($report.Items|Where-Object{$_.Status -in @('FAILED_PRESERVED','FAILED_UNVERIFIED','REJECTED_PRESERVED','NOT_ATTEMPTED')}|ForEach-Object{
            $item=ConvertTo-ConverterProblem $_ $_.Status $null;$item.Kind=$_.Kind;$item
        })
        $issues=@($report.Issues|ForEach-Object{
            $item=ConvertTo-ConverterProblem $_ 'PROCESSING_ISSUE' $null;$item.Key=$_.Key;$item.OriginalPreserved=$_.OriginalPreserved
            if($UltraConversion){$item.BodyBlock=$_.BodyBlock}
            $item
        })
        $warnings=@($report.Warnings|ForEach-Object{
            $item=ConvertTo-ConverterProblem $_ 'WARNING' 'LAYOUT_CHECK';$item.Code=$_.Code;$item.Reason=$_.Code;$item.Actual=$_.Actual;$item.Limit=$_.Limit;$item
        })
        $mode=if($UltraConversion){'Ultra'}elseif($FastConversion){'Fast'}else{'Standard'}
        $verification=if($UltraConversion){'ULTRA_UNVERIFIED'}elseif($FastConversion){'FAST_UNVERIFIED'}else{'STANDARD_VERIFIED'}
        $finished=[ordered]@{SchemaVersion=1;Event='FINISHED';Status=$report.Status;ExitCode=0;ApplicationVersion=$report.ApplicationVersion;ConvertOmml=$(if($UltraConversion){$true}else{$report.ConvertOmml});FastConversion=[bool]$report.FastConversion;ConversionMode=$mode;VerificationMode=$verification;Validation=$report.Validation;Output=$outputFile;ReportPath=$reportPath;TaskDirectory=$task;StopFile=$stopFile;LogPath=$logPath;Targets=$report.Targets;Succeeded=$report.Succeeded;FailedPreserved=$report.FailedPreserved;FailedUnverified=$report.FailedUnverified;RejectedPreserved=$report.RejectedPreserved;NotAttempted=$report.NotAttempted;NumbersAdded=$report.NumbersAdded;NumbersUpdated=$report.NumbersUpdated;NumberReferencesUpdated=$report.NumberReferencesUpdated;ExistingDisplaysFormatted=$report.ExistingDisplaysFormatted;Issues=$issues;Warnings=$warnings;LayoutWarningsChecked=[bool]$report.LayoutWarningsChecked;LayoutWarningScope=$report.LayoutWarningScope;PreservedItems=$preserved;OriginalFileBytesUnchanged=$report.OriginalFileBytesUnchanged;WordStarted=$report.WordStarted;WordReopenVerified=$report.WordReopenVerified;MathematicalEquivalenceVerified=$false}
        if($targeted){$finished['TargetedBodyConversion']=$true;$finished['SelectedTargetIds']=$TargetId;$finished['UnselectedFormulas']=$report.UnselectedFormulas}
        Write-ConverterJson $finished
    }
}catch{Write-ConverterFailure $_;$exitCode=1}
finally{
    if($targetList -and [IO.File]::Exists($targetList)){Remove-Item -LiteralPath $targetList -ErrorAction SilentlyContinue}
}
exit $exitCode
