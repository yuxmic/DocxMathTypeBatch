#requires -Version 7.0
# 高级设置只测真实效果，不将自定义行距和编号计入默认性能测量。
param([switch]$SkipBuild,[ValidateSet('All','Omml')][string]$Case='All')
$ErrorActionPreference='Stop'
$selectedCase=$Case
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
$run=Join-Path $ProjectRoot ('tests\output\高级设置-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
$evidence=[ordered]@{Status='FAILED';Case=$selectedCase;Directory=$run;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@()}
if($Case -eq 'All'){
    & python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_advanced.py') fixture $run
    if($LASTEXITCODE -ne 0){throw '高级样本生成失败。'}
    $source=Join-Path $run 'advanced-input.docx';$hash=(Get-FileHash -LiteralPath $source).Hash
}
try{
    if($Case -eq 'All'){
    foreach($case in @(
        @{Name='follow';Inline='follow';Display='follow';Extra=@()},
        @{Name='single';Inline='single';Display='double';Extra=@()},
        @{Name='onehalf';Inline='onehalf';Display='single';Extra=@()},
        @{Name='double';Inline='double';Display='onehalf';Extra=@()},
        @{Name='exact';Inline='exact:24';Display='atleast:28';Extra=@()},
        @{Name='atleast';Inline='atleast:28';Display='multiple:1.25';Extra=@()},
        @{Name='multiple';Inline='multiple:1.25';Display='exact:24';Extra=@()},
        @{Name='alphabetic';Inline='follow';Display='follow';Extra=@('--numbers','--number-format=[#CI.#Sa-#EA]','--number-chapter=4','--number-section=2','--number-start=27')},
        @{Name='roman';Inline='follow';Display='follow';Extra=@('--numbers','--number-format=#Ei','--number-start=9')},
        @{Name='numeric';Inline='follow';Display='follow';Extra=@('--numbers','--number-format=（#C1-#S1-#E1）','--number-chapter=3','--number-section=2','--number-start=9')}
    )){
        $name=$case.Name;$arguments=@('--worker-convert',$source,(Join-Path $run ($name+'.docx')),(Join-Path $run ($name+'.json')),'paragraph',(Join-Path $run 'stop.flag'),('--inline-spacing='+$case.Inline),('--display-spacing='+$case.Display))+$case.Extra
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{ '"'+$_+'"' }) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $run ($name+'.log')) -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
        try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('高级转换失败：'+$name)}}finally{$process.Dispose()}
        $check=(& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_advanced.py') $name $run)|ConvertFrom-Json
        if($LASTEXITCODE -ne 0 -or $check.Status -ne 'PASSED'){throw ('高级设置未正确生效：'+$name)}
        if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw '高级设置修改了原件。'}
        $evidence.Cases+=@{Name=$name;Check=$check;Output=(Join-Path $run ($name+'.docx'))}
        Write-Output ('ADVANCED_CASE_PASSED: '+$name)
    }
    # 无效配置必须在写输出前失败，不能留下部分文档。
    foreach($case in @(@{Name='invalid-spacing';Argument='--inline-spacing=multiple:0'},@{Name='invalid-format';Argument='--number-format=(#E1.#E1)'})){
        $arguments=@('--worker-convert',$source,(Join-Path $run ($case.Name+'.docx')),(Join-Path $run ($case.Name+'.json')),'paragraph',(Join-Path $run 'stop.flag'),'--numbers',$case.Argument)
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{ '"'+$_+'"' }) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $run ($case.Name+'.log')) -PassThru
        try{if(-not $process.WaitForExit(30000) -or $process.ExitCode -eq 0 -or (Test-Path -LiteralPath (Join-Path $run ($case.Name+'.docx')))){throw '无效配置未安全拒绝。'}}finally{$process.Dispose()}
        $evidence.Cases+=@{Name=$case.Name;RejectedWithoutOutput=$true}
    }
    }
    # 只增加一组开关回归；反向复用关闭后的输出，默认开启沿用同一份区域混合样本。
    $ommlRun=Join-Path $run 'omml-switch';New-Item -ItemType Directory -Path $ommlRun|Out-Null
    $helper=Join-Path $PSScriptRoot 'helpers/test_advanced.py'
    & python -B -X utf8 $helper omml-switch-fixture $ommlRun
    if($LASTEXITCODE -ne 0){throw 'OMML 开关样本生成失败。'}
    foreach($name in @('default-on','off-normal','off-fast','off-reverse','off-no-targets')){
        $reverse=$name -eq 'off-reverse';$inputFile=Join-Path $ommlRun $(if($reverse){'off-normal.docx'}elseif($name -eq 'off-no-targets'){'omml-only-input.docx'}else{'input.docx'})
        $inputHash=(Get-FileHash -LiteralPath $inputFile).Hash
        $output=Join-Path $ommlRun ($name+'.docx');$report=Join-Path $ommlRun ($name+'.json');$log=Join-Path $ommlRun ($name+'.log')
        $arguments=@($(if($reverse){'--worker-reverse'}else{'--worker-convert'}),$inputFile,$output,$report,$(if($reverse){'bracket'}else{'paragraph'}),(Join-Path $ommlRun ($name+'.flag')))
        $arguments+=@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        if($name -ne 'default-on'){$arguments+='--skip-omml'}
        if($name -eq 'off-fast'){$arguments+='--experimental-fast'}
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $ommlRun ($name+'.errors')) -PassThru
        try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('OMML 开关转换失败：'+$name+'；'+$log)}}finally{$process.Dispose()}
        if((Get-FileHash -LiteralPath $inputFile).Hash -ne $inputHash){throw ('OMML 开关测试修改了原件：'+$name)}
        $check=(& python -B -X utf8 $helper omml-switch-check $ommlRun $name)|ConvertFrom-Json
        if($LASTEXITCODE -ne 0 -or $check.Status -ne 'PASSED'){throw ('OMML 开关独立检查失败：'+$name)}
        $evidence.Cases+=@{Name=$name;Check=$check;InputSHA256=$inputHash;Report=$report}
        Write-Output ('OMML_SWITCH_CASE_PASSED: '+$name)
    }
    $evidence.Status='PASSED'
}finally{
    $evidence.SourceSHA256=Get-TestSourceHashes
    $evidence|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $run 'advanced-results.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'advanced-results.json') -Encoding utf8
}
