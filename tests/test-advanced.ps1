#requires -Version 7.0
# 高级设置只测真实效果，不将自定义行距和编号计入默认性能测量。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
$run=Join-Path $ProjectRoot ('tests\output\高级设置-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
$evidence=[ordered]@{Status='FAILED';Directory=$run;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@()}
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_advanced.py') fixture $run
if($LASTEXITCODE -ne 0){throw '高级样本生成失败。'}
$source=Join-Path $run 'advanced-input.docx';$hash=(Get-FileHash -LiteralPath $source).Hash
try{
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
    $evidence.Status='PASSED'
}finally{
    $evidence.SourceSHA256=@{};foreach($file in @('gui.cpp','convert_docx.cpp','layout_options.h','native_numbering.h','native_number_text.py','test-advanced.ps1','test_advanced.py','test_native_display.py')){$evidence.SourceSHA256[$file]=(Get-FileHash -LiteralPath (Get-TestSourcePath $file)).Hash}
    $evidence|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $run 'advanced-results.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'advanced-results.json') -Encoding utf8
}
