# 快速模式与四类实验位置组合回归：使用合成文档，不读取私人样本。
param([switch]$SkipBuild,[ValidateSet('conversion','blocked')][string[]]$Case=@('conversion','blocked'))
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
$run=Join-Path $TestOutputRoot ('快速实验区域-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run|Out-Null
$helper=Join-Path $PSScriptRoot 'helpers/test_experimental.py'
if($Case -contains 'conversion'){
    & python -B -X utf8 $helper fixture $run
    if($LASTEXITCODE -ne 0){throw '实验区域样本生成失败。'}
}
$source=Join-Path $run 'input.docx'
$exe=Join-Path $ProjectRoot 'build/bin/DocxMathTypeBatch.exe'
$evidence=[ordered]@{Status='FAILED';Directory=$run;Cases=@();ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash}
try{
    if($Case -contains 'conversion'){
    foreach($name in @('fast-forward','fast-reverse','normal-forward')){
        $reverse=$name -eq 'fast-reverse';$fast=$name -ne 'normal-forward'
        $inputFile=if($reverse){Join-Path $run 'fast-forward.docx'}else{$source}
        $hash=(Get-FileHash -LiteralPath $inputFile).Hash
        $output=Join-Path $run ($name+'.docx');$report=Join-Path $run ($name+'.json');$log=Join-Path $run ($name+'.log')
        $arguments=@($(if($reverse){'--worker-reverse'}else{'--worker-convert'}),$inputFile,$output,$report,$(if($reverse){'bracket'}else{'paragraph'}),(Join-Path $run 'stop.flag'))
        $arguments+=@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        if($fast){$arguments+='--experimental-fast'}
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
        try{
            if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('实验区域任务失败：'+$name+'；日志：'+$log)}
        }finally{$process.Dispose()}
        $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
        if($result.Status -ne 'COMPLETE' -or $result.Succeeded -ne 21 -or $result.Experimental.Succeeded -ne 18 -or $result.Experimental.Issues.Count -ne 0){throw ('实验区域组合失败：'+$name+'；报告：'+$report+"`n"+($result.Experimental.Issues|ConvertTo-Json -Depth 4))}
        if((Get-FileHash -LiteralPath $inputFile).Hash -ne $hash){throw '输入文档被修改。'}
        if($fast -and ($result.Validation -ne 'SKIPPED_EXPERIMENTAL_FAST' -or $result.PostConversionScans -ne 0 -or $result.WordReopenVerified)){throw '快速模式错误启用了保存后核验。'}
        $trace=[IO.File]::ReadAllText($log)
        if($fast -ne $trace.Contains('FAST_WORD_EXIT_MS=')){throw '快速与常规模式的 Word 收尾路径不符。'}
        # 从输出包独立核对全部区域的对象、公式及表格属性，不仅检查成功计数。
        if($reverse){& python -B -X utf8 $helper check-reverse $source $output}
        else{& python -B -X utf8 $helper check $source $output all $source}
        if($LASTEXITCODE -ne 0){throw ('输出内容检查失败：'+$name)}
        $evidence.Cases+=@{Name=$name;Succeeded=$result.Succeeded;ExperimentalSucceeded=$result.Experimental.Succeeded;InputUnchanged=$true;Report=$report}
        Write-Output ('FAST_EXPERIMENTAL_CASE_PASSED: '+$name)
    }
    }
    if($Case -contains 'blocked'){
        & python -B -X utf8 $helper blocked-fixture $run
        if($LASTEXITCODE -ne 0){throw '入口拒绝样本生成失败。'}
        foreach($kind in @('tracking','protected')){foreach($direction in @('forward','reverse')){foreach($fast in @($false,$true)){
            # 两种模式、两个方向均须在初始化 MathType 前拒绝，且不交付输出。
            $name=$kind+'-'+$direction+$(if($fast){'-fast'}else{'-normal'})
            $inputFile=Join-Path $run ($kind+'.docx');$hash=(Get-FileHash -LiteralPath $inputFile).Hash
            $output=Join-Path $run ($name+'.docx');$report=Join-Path $run ($name+'.json');$log=Join-Path $run ($name+'.log')
            $arguments=@($(if($direction -eq 'reverse'){'--worker-reverse'}else{'--worker-convert'}),$inputFile,$output,$report,$(if($direction -eq 'reverse'){'bracket'}else{'paragraph'}),(Join-Path $run 'stop.flag'))
            if($fast){$arguments+='--experimental-fast'}
            $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
            try{if(-not $process.WaitForExit(30000) -or $process.ExitCode -eq 0){throw ('应拒绝的输入未被拒绝：'+$name)}}finally{$process.Dispose()}
            $reason=if($kind -eq 'tracking'){'Disable revision tracking in source before batch conversion'}else{'Protected document cannot be converted'}
            $text=[IO.File]::ReadAllText($log)
            if(-not $text.Contains($reason) -or $text.Contains('UI_PHASE INITIALIZING') -or (Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath $report) -or (Get-FileHash -LiteralPath $inputFile).Hash -ne $hash){throw ('入口拒绝原因、时机或原件保护不符：'+$name+'；日志：'+$log)}
            $evidence.Cases+=@{Name=$name;Rejected=$true;InputUnchanged=$true;MathTypeNotInitialized=$true}
            Write-Output ('INPUT_REJECTION_PASSED: '+$name)
        }}}
    }
    $evidence.Status='PASSED'
}finally{
    $evidence|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $run 'fast-experimental-results.json') -Encoding utf8
    Write-Output ('FAST_EXPERIMENTAL_EVIDENCE: '+$run)
}
