#requires -Version 7.0
# 版式提醒定向验证；复用一份合成文档，按需选择模式，不启动全套回归。
param([switch]$SkipBuild,[string]$RunDirectory,[ValidateSet('normal','fast','skip-omml','reverse','recovery','geometry')][string[]]$Case=@('normal','fast','skip-omml','reverse','recovery','geometry'))
$ErrorActionPreference='Stop'
$selectedCase=@($Case)
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target All}
$run=if($RunDirectory){[IO.Path]::GetFullPath($RunDirectory)}else{Join-Path $TestOutputRoot ('版式提醒-'+[Guid]::NewGuid().ToString('N'))}
New-Item -ItemType Directory -Path $run -Force|Out-Null
$helper=Join-Path $PSScriptRoot 'helpers/test_layout_warnings.py'
if(-not (Test-Path -LiteralPath (Join-Path $run 'layout-input.docx'))){
    & python -B -X utf8 $helper fixture $run
    if($LASTEXITCODE -ne 0){throw '版式样本生成失败。'}
}
$checks=@()
foreach($name in $selectedCase){
    if($name -eq 'geometry'){
        # 独立隐藏 Word 只校准栏宽；不打开用户文档，不保存新文档。
        $word=$null;$document=$null
        try{
            $word=New-Object -ComObject Word.Application
            $word.Visible=$false;$word.DisplayAlerts=0
            $document=$word.Documents.Add();$page=$document.Sections.Item(1).PageSetup
            $page.PageWidth=612;$page.LeftMargin=72;$page.RightMargin=72;$page.Gutter=24
            $single=[double]$page.TextColumns.Item(1).Width
            $page.MirrorMargins=-1;$mirror=[double]$page.TextColumns.Item(1).Width
            $page.TextColumns.SetCount(2);$page.TextColumns.Spacing=18
            $equal=[double]$page.TextColumns.Item(1).Width
            if([Math]::Abs($single-444) -gt 1 -or [Math]::Abs($mirror-444) -gt 1 -or [Math]::Abs($equal-213) -gt 1){throw "实际栏宽与装订线计算不符：$single / $mirror / $equal"}
            $checks+=@{Case=$name;Status='PASSED';Single=$single;Mirror=$mirror;Equal=$equal}
        }finally{if($document){$document.Close(0)};if($word){$word.Quit()}}
        continue
    }
    $inputFile=Join-Path $run $(if($name -eq 'reverse'){'normal.docx'}elseif($name -eq 'recovery'){'rollback-input.docx'}else{'layout-input.docx'})
    $inputHash=(Get-FileHash -LiteralPath $inputFile).Hash
    $output=Join-Path $run ($name+'.docx');$report=Join-Path $run ($name+'.json');$stop=Join-Path $run ($name+'.flag')
    if((Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath $report)){throw ('本轮结果已存在：'+$name)}
    $exe=Join-Path $ProjectRoot ('build/bin/'+$(if($name -eq 'recovery'){'docx-math-convert.exe'}else{'DocxMathTypeBatch.exe'}))
    $arguments=@($(if($name -eq 'reverse'){'--worker-reverse'}elseif($name -eq 'recovery'){'--self-test'}else{'--worker-convert'}),$inputFile,$output,$report,$(if($name -eq 'reverse'){'bracket'}else{'paragraph'}))
    if($name -ne 'recovery'){$arguments+=$stop}
    $arguments+=if($name -eq 'recovery'){@('--experimental-tables','--skip-omml')}else{@('tables','textboxes','headers','notes'|ForEach-Object {'--experimental-'+$_})}
    if($name -eq 'fast'){$arguments+='--experimental-fast'}
    if($name -eq 'skip-omml'){$arguments+='--skip-omml'}
    $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object {'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $run ($name+'.log')) -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
    try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('版式转换失败：'+$name)}}finally{
        if(-not $process.HasExited -and $name -ne 'recovery'){Set-Content -LiteralPath $stop -Value 'stop';$null=$process.WaitForExit(30000)}
        $process.Dispose()
    }
    if((Get-FileHash -LiteralPath $inputFile).Hash -ne $inputHash){throw ('验证修改了原件：'+$name)}
    if($name -eq 'recovery'){
        $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
        $log=Get-Content -LiteralPath (Join-Path $run ($name+'.log')) -Raw
        if($result.Status -ne 'PARTIAL' -or $result.FailedPreserved -ne 1 -or $result.Succeeded -ne 5 -or $log -notmatch 'EXPERIMENTAL_RECOVERY text=1 layout=1 source=1'){throw '局部恢复未保留源式或未继续处理。'}
        $checks+=@{Case=$name;Status='PASSED';Succeeded=$result.Succeeded;FailedPreserved=$result.FailedPreserved}
    }else{
        $check=(& python -B -X utf8 $helper $name $run)|ConvertFrom-Json
        if($LASTEXITCODE -ne 0){throw ('版式提醒检查失败：'+$name)}
        $checks+=$check
    }
    Write-Output ('LAYOUT_CASE_PASSED: '+$name)
}
[ordered]@{Status='PASSED';ExecutableSHA256=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot 'build/bin/DocxMathTypeBatch.exe')).Hash;Cases=$checks}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $run 'layout-results.json') -Encoding utf8
