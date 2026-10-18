#requires -Version 7.0
# 先不编号转换，再补编号、重复更新及混入新公式；所有步骤另存副本，不覆盖原件。
param([switch]$SkipBuild,[string]$RunDirectory)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target All}
if(-not $RunDirectory){$RunDirectory=Join-Path $TestOutputRoot ('原生显示-'+[Guid]::NewGuid().ToString('N'))}
$run=[IO.Path]::GetFullPath($RunDirectory);New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') fixture $run
if($LASTEXITCODE -ne 0){throw '显示布局样本生成失败。'}
function Invoke-DisplayCase([string]$Name,[string]$InputFile,[int]$Targets,[int]$Added,[int]$Updated,[int]$Displays,[int]$Start=0,[int]$CheckStart=$Start){
    $output=Join-Path $run ($Name+'.docx');$report=Join-Path $run ($Name+'.json');$log=Join-Path $run ($Name+'.log')
    if(Test-Path -LiteralPath $output){throw ('测试输出已存在：'+$output)}
    $hash=(Get-FileHash -LiteralPath $InputFile).Hash
    $arguments=@('--worker-convert',$InputFile,$output,$report,'paragraph',(Join-Path $run 'stop.flag'))
    if($Start){$arguments+=@('--numbers',('--number-start='+$Start))}
    $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($Name+'.errors')) -PassThru
    try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('显示布局转换失败：'+$Name+'；'+$log)}}finally{$process.Dispose()}
    $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
    if($result.Status -ne 'COMPLETE' -or $result.Targets -ne $Targets -or $result.Succeeded -ne $Targets -or $result.NumbersAdded -ne $Added -or $result.NumbersUpdated -ne $Updated){throw ('显示布局数量错误：'+$Name+'；'+$report)}
    & python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') check $InputFile $output $Displays $CheckStart|ForEach-Object{Write-Host $_}
    if($LASTEXITCODE -ne 0 -or (Get-FileHash -LiteralPath $InputFile).Hash -ne $hash){throw ('显示布局或原件核验失败：'+$Name)}
    if($Targets -eq 0 -and (Get-Content -LiteralPath $log -Raw).Contains('DEFAULT_PREFERENCES_LENGTH')){throw '纯编号任务不应初始化 MathType。'}
    Write-Host ('NATIVE_DISPLAY_CASE_PASSED: '+$Name)
    return $output
}
$unnumbered=Invoke-DisplayCase 'unnumbered' (Join-Path $run 'input.docx') 6 0 0 5
$legacy=Join-Path $run 'legacy-input.docx'
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') legacy $unnumbered $legacy
if($LASTEXITCODE -ne 0){throw '旧版显示格式样本生成失败。'}
$legacyOutput=Invoke-DisplayCase 'legacy-unnumbered' $legacy 0 0 0 5
$numbered=Invoke-DisplayCase 'numbered-later' $unnumbered 0 5 0 5 1
$preserved=Invoke-DisplayCase 'preserve-numbered' $numbered 0 0 0 5 0 1
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') preserved-numbers $numbered $preserved
if($LASTEXITCODE -ne 0){throw '默认转换改变了原有编号。'}
$again=Invoke-DisplayCase 'numbered-again' $numbered 0 0 5 5 1
$mixed=Join-Path $run 'mixed-input.docx'
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') mixed $unnumbered $mixed
if($LASTEXITCODE -ne 0){throw '混合新旧公式样本生成失败。'}
$mixedOutput=Invoke-DisplayCase 'mixed' $mixed 1 6 0 6 1
$reset=Invoke-DisplayCase 'reset' $mixedOutput 0 0 6 6 7
$resetDefault=Invoke-DisplayCase 'reset-default' $reset 0 0 6 6 1
$referenceInput=Join-Path $run 'reference-input.docx'
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') reference $numbered $referenceInput
if($LASTEXITCODE -ne 0){throw '编号引用样本生成失败。'}
$referenceOutput=Invoke-DisplayCase 'reference' $referenceInput 1 1 5 6 1
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') check-reference $referenceOutput
if($LASTEXITCODE -ne 0){throw '原编号引用没有同步更新。'}
& (Join-Path $ProjectRoot 'build\bin\native-api-probe.exe') --display-self-test $unnumbered $run|ForEach-Object{Write-Host $_}
if($LASTEXITCODE -ne 0){throw 'MathType 原生手动插入编号失败。'}
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') manual (Join-Path $run 'manual-number.docx')
if($LASTEXITCODE -ne 0){throw '手动编号实际 PDF 右对齐核验失败。'}
$interop=Join-Path $run 'native-commands';New-Item -ItemType Directory -Path $interop|Out-Null
Copy-Item -LiteralPath $numbered -Destination (Join-Path $interop 'converted.docx')
& (Join-Path $ProjectRoot 'build\bin\native-api-probe.exe') --numbering-self-test (Join-Path $interop 'converted.docx') $interop|ForEach-Object{Write-Host $_}
if($LASTEXITCODE -ne 0){throw '原生编号更新、格式化或追加检查失败。'}
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_numbering.py') check $interop 5 new
if($LASTEXITCODE -ne 0){throw '原生更新、格式化后的可见 PDF 编号错误。'}
$evidence=[ordered]@{Status='PASSED';Directory=$run;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Unnumbered=$unnumbered;LegacyNormalized=$legacyOutput;NumberedLater=$numbered;ExistingNumbersPreserved=$preserved;Repeated=$again;Mixed=$mixedOutput;Reset=$reset;ReferenceUpdated=$referenceOutput;NoDuplicateNumbers=$true;OriginalsUnchanged=$true;OldEquationDataUnchanged=$true;NumberOnlyNoMathTypeActivation=$true;NativeManualInsertRightAligned=$true;NativeUpdateFormatAppend=$true;SourceSHA256=[ordered]@{}}
foreach($file in @('gui.cpp','convert_docx.cpp','docx_scan.h','native_numbering.h','native_api_probe.cpp','test_native_numbering.h','test-native-display.ps1','test_native_display.py')){$evidence.SourceSHA256[$file]=(Get-FileHash -LiteralPath (Get-TestSourcePath $file)).Hash}
$evidence|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $run 'native-display-results.json') -Encoding utf8
$evidence|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'native-display-results.json') -Encoding utf8
Write-Output ('NATIVE_DISPLAY_PASSED: '+$run)
