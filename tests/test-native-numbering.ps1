#requires -Version 7.0
# 功能实测不计入性能数据；编号更改必须经 MathType 原生过程及 PDF 独立核对。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target All}
$run=Join-Path $ProjectRoot ('tests\output\原生编号-'+[Guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe';$probe=Join-Path $ProjectRoot 'build\bin\native-api-probe.exe'
$evidence=[ordered]@{Status='FAILED';Directory=$run;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;ProbeExecutableSHA256=(Get-FileHash -LiteralPath $probe).Hash;Cases=@()}
function Invoke-NativeCheck([string]$Program,[string[]]$Arguments,[string]$Name){
    $log=Join-Path $run ($Name+'.log');$process=Start-Process -FilePath $Program -ArgumentList @($Arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($Name+'.errors')) -PassThru
    try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('原生编号检查失败：'+$Name+'；'+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
}
try{
    & python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_numbering.py') fixture $run
    if($LASTEXITCODE -ne 0){throw '原生编号样本生成失败。'}
    $source=Join-Path $run 'input.docx'
    foreach($kind in @('new','mixed','custom')){
        $caseDir=Join-Path $run $kind;New-Item -ItemType Directory -Path $caseDir|Out-Null
        $input=$source
        if($kind -eq 'mixed'){
            Invoke-NativeCheck $probe @('--numbering-fixture',$source,$caseDir) 'create-vendor-fixture'
            $input=Join-Path $caseDir 'native-fixture.docx'
        }
        $hash=(Get-FileHash -LiteralPath $input).Hash;$output=Join-Path $caseDir 'converted.docx';$report=Join-Path $caseDir 'converted.json'
        $arguments=@('--worker-convert',$input,$output,$report,'paragraph',(Join-Path $caseDir 'stop.flag'),'--numbers')
        if($kind -eq 'custom'){$arguments+=@('--number-custom','--number-format=[#CI.#Sa-#EA]','--number-start=27','--number-chapter=4','--number-section=2')}
        Invoke-NativeCheck $exe $arguments ($kind+'-convert')
        $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
        if($result.Status -ne 'COMPLETE' -or $result.NumbersAdded -ne 5 -or $result.Succeeded -ne 6){throw '新增编号或公式数量不符。'}
        Invoke-NativeCheck $probe @('--numbering-self-test',$output,$caseDir) ($kind+'-native-commands')
        $count=if($kind -eq 'mixed'){6}else{5}
        & python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_numbering.py') check $caseDir $count $kind |Set-Content -LiteralPath (Join-Path $caseDir 'pdf-check.jsonl') -Encoding utf8
        if($LASTEXITCODE -ne 0){throw '原生编号 PDF 实际显示未通过。'}
        if((Get-FileHash -LiteralPath $input).Hash -ne $hash){throw '原生编号验证修改了输入。'}
        $evidence.Cases+=@{Name=$kind;NoNewEquationBreaks=$true;NativeUpdate=$true;NativeFormat=$true;NativeAppend=$true;SavedReopen=$true;EquationAssetsUnchanged=$true;PdfNumbersVerified=$true;Count=$count;Directory=$caseDir}
        Write-Output ('NATIVE_NUMBERING_PASSED: '+$kind)
    }
    $evidence.Status='PASSED'
}finally{
    $evidence.SourceSHA256=Get-TestSourceHashes
    $evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $run 'native-numbering-results.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'native-numbering-results.json') -Encoding utf8
}
