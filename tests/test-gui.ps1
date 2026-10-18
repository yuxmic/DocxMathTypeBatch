#requires -Version 7.0
# 最小真实 GUI 验收：同一 EXE 的扫描/转换、设置检查、过期清单、防覆盖和安全停止。
param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$guiTempRoot = [IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output\work'))
$guiIdentity = [Guid]::NewGuid().ToString('N')
$guiTask = Join-Path $guiTempRoot ('GUI-' + $guiIdentity)
$guiPublished = Join-Path $ProjectRoot ('tests\output\GUI-' + $guiIdentity)
New-Item -ItemType Directory -Path $guiTask | Out-Null
$guiEvidence = [ordered]@{CheckedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');Status='FAILED';Scope='阶段4 B双栏GUI正向连接';Executable=(Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe');Cases=@();Screenshots=@();Checks=@{}}
$guiChildActive = $false

. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
try {
    if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
    $baseline = New-TestFixtures -Native
    foreach($case in @(
        @{Name='complete';Input=(Join-Path $baseline.Directory 'mixed.docx');Count=5;Status='COMPLETE'},
        @{Name='partial';Input=(Join-Path $baseline.Directory 'boundaries.docx');Count=2;Status='PARTIAL'}
    )){
        $folder=Join-Path $guiTask $case.Name;New-Item -ItemType Directory -Path $folder|Out-Null
        $inputFile=Join-Path $folder 'input.docx';Copy-Item -LiteralPath $case.Input -Destination $inputFile
        $hash=(Get-FileHash -LiteralPath $inputFile).Hash
        $ready=Join-Path $folder '准备就绪.png';$resultImage=Join-Path $folder '转换结果.png'
        $occupied=Join-Path $folder 'input_normal.docx';Copy-Item -LiteralPath $inputFile -Destination $occupied
        $process=Start-Process -FilePath $guiEvidence.Executable -ArgumentList @('--ui-test',('"'+$inputFile+'"'),('"'+$folder+'"'),'normal') -WindowStyle Hidden -PassThru
        $readySaved=$false;$resultSaved=$false;$deadline=[DateTime]::UtcNow.AddSeconds(120)
        try{
            while(-not $process.WaitForExit(100)){
                if(-not $readySaved -and (Test-Path -LiteralPath (Join-Path $folder 'gui-ready.json'))){$readySaved=Save-GuiWindow $process.Id $ready}
                if(-not $resultSaved -and (Test-Path -LiteralPath (Join-Path $folder 'gui-check.json'))){$resultSaved=Save-GuiWindow $process.Id $resultImage}
                if([DateTime]::UtcNow -gt $deadline){throw 'GUI 回归超时。'}
            }
            if($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath (Join-Path $folder 'gui-check.json'))){throw ('GUI 控件或工作进程失败：'+$case.Name+'；'+$folder)}
        }finally{$process.Dispose()}
        $check=Get-Content -LiteralPath (Join-Path $folder 'gui-check.json') -Raw|ConvertFrom-Json
        $report=Get-Content -LiteralPath $check.Report -Raw|ConvertFrom-Json
        if($report.Status -ne $case.Status -or $report.Succeeded -ne $case.Count -or $report.NumbersAdded -ne 0 -or $report.Preferences -notlike 'size=18*'){throw 'GUI 参数未正确传入转换核心。'}
        foreach($flag in @('ControlValidation','SizeArrows','ParameterRepaint','KeyboardTabNavigation','Layout','NoPreScan','ScanProgress','SavedProgress','ReopenMarquee')){if(-not $check.$flag){throw ('GUI 检查未通过：'+$flag)}}
        if(-not $readySaved -or -not $resultSaved -or (Get-FileHash -LiteralPath $inputFile).Hash -ne $hash -or (Get-FileHash -LiteralPath $occupied).Hash -ne $hash -or [IO.Path]::GetFileName($check.Output) -ne 'input_normal (2).docx'){throw 'GUI 截图、原件或防覆盖检查失败。'}
        if(Test-Path -LiteralPath (Join-Path $folder 'gui-style.eqp')){throw 'GUI 仍生成旧字体样式文件。'}
        $guiEvidence.Cases+=@{Name=$case.Name;Outcome=$report.Status;Succeeded=$report.Succeeded;Input=(Join-Path $guiPublished ($case.Name+'\input.docx'));Preferences='paragraph';Check=(Join-Path $guiPublished ($case.Name+'\gui-check.json'));Report=(Join-Path $guiPublished ($case.Name+'\report.json'));ReadyScreenshot=(Join-Path $guiPublished ($case.Name+'\准备就绪.png'));ResultScreenshot=(Join-Path $guiPublished ($case.Name+'\转换结果.png'))}
        Write-Output ('GUI_CASE_PASSED: '+$case.Name)
    }
    $guiEvidence.Status='PASSED';$guiEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $guiEvidence.Executable).Hash
    $guiEvidence.SourceSHA256=@{}
    foreach($file in @('gui.cpp','convert_docx.cpp','docx_scan.h','test-gui.ps1')){$guiEvidence.SourceSHA256[$file]=(Get-FileHash -LiteralPath (Get-TestSourcePath $file)).Hash}
}finally{
    # 两个绝对路径均由项目的固定临时根与验证根构造，仅移动本轮证据。
    if(-not [IO.Path]::GetFullPath($guiTask).StartsWith($guiTempRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'GUI 证据来源超出临时目录。'}
    if(-not [IO.Path]::GetFullPath($guiPublished).StartsWith([IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output'))+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'GUI 证据目标超出验证目录。'}
    Move-Item -LiteralPath $guiTask -Destination $guiPublished
    $guiEvidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'gui-results.json') -Encoding utf8
}
