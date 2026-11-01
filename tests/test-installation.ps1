#requires -Version 7.0
# 安装路径专项：使用 D 盘中文/空格目录的组件副本，不移动实际安装或修改真实 COM 注册。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$installRoot=[IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output\work'))
$installIdentity=[Guid]::NewGuid().ToString('N')
$installTask=Join-Path $installRoot ('安装路径-'+$installIdentity)
$installPublished=Join-Path $ProjectRoot ('tests\output\安装路径-'+$installIdentity)
$installAlternate=Join-Path $installTask '非默认目录 with spaces\MathType'
$installExe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
$installProbe=Join-Path $ProjectRoot 'build\bin\native-api-probe.exe'
$installEvidence=[ordered]@{Status='FAILED';CheckedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');Cases=@();Checks=@{};SampleDirectory=$installPublished}
New-Item -ItemType Directory -Path $installAlternate -Force|Out-Null

function Invoke-InstallationTask {
    param([string]$Executable,[string[]]$Arguments,[string]$Name,[int]$ExpectedExit=0)
    $installLog=Join-Path $installTask ($Name+'.log')
    $installErrors=Join-Path $installTask ($Name+'.errors')
    $installProcess=Start-Process -FilePath $Executable -ArgumentList @($Arguments|ForEach-Object {'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $installLog -RedirectStandardError $installErrors -PassThru
    try {
        $installDeadline=[DateTime]::UtcNow.AddSeconds(180)
        while(-not $installProcess.WaitForExit(500)) {
            if([DateTime]::UtcNow -gt $installDeadline){$installProcess.Kill();$installProcess.WaitForExit();throw ('安装路径专项超时：'+$Name)}
        }
        if($installProcess.ExitCode -ne $ExpectedExit){throw ('安装路径专项退出码不符：'+$Name+'，详见 '+$installLog)}
        $installEvidence.Cases+=@{Name=$Name;ExitCode=$installProcess.ExitCode;Log=($Name+'.log')}
        return [IO.File]::ReadAllText($installLog)
    }finally{$installProcess.Dispose()}
}

try {
    if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target All}
    $installClsid=(Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4\CLSID').'(default)'
    $installServer=(Get-ItemProperty -LiteralPath ('Registry::HKEY_CLASSES_ROOT\CLSID\'+$installClsid+'\LocalServer32')).'(default)'
    if($installServer -notmatch '^(?:"(?<exe>[^"]+)"|(?<exe>.+?MathType\.exe))(?:\s|$)'){throw '测试无法读取实际 MathType 注册路径。'}
    $installOriginal=Split-Path -Parent $Matches.exe
    $installComponents=@('MathType.exe','MathPage\64\MathPage.wll','Office Support\BlankEqn.doc','Preferences\Times+Symbol 12.eqp','Translators\AMS LaTeX.tdl')
    $installHashes=@{}
    foreach($installComponent in $installComponents){
        $installSource=Join-Path $installOriginal $installComponent;$installTarget=Join-Path $installAlternate $installComponent
        $installHashes[$installComponent]=(Get-FileHash -LiteralPath $installSource).Hash
        New-Item -ItemType Directory -Path (Split-Path -Parent $installTarget) -Force|Out-Null
        Copy-Item -LiteralPath $installSource -Destination $installTarget
    }
    $installDiscovery=Invoke-InstallationTask $installProbe @('--installation-self-test',$installAlternate) 'registry-discovery'
    if($installDiscovery -notmatch 'INSTALLATION_DISCOVERY_PASSED' -or @([regex]::Matches($installDiscovery,'INSTALLATION_CASE .* PASSED')).Count -ne 11){throw '隔离注册表发现路径未通过。'}
    foreach($installComponent in $installComponents){
        # 只暂时移动本轮组件副本，逐项验证错误明确指出缺失文件，然后恢复。
        $installTarget=[IO.Path]::GetFullPath((Join-Path $installAlternate $installComponent))
        if(-not $installTarget.StartsWith($installTask+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '组件副本路径越界。'}
        $installRemoved=$installTarget+'.removed';Move-Item -LiteralPath $installTarget -Destination $installRemoved
        try {
            $installMissing=Invoke-InstallationTask $installProbe @('--installation-check',$installAlternate) ('missing-'+[IO.Path]::GetFileName($installComponent)) 1
            if(-not $installMissing.Contains($installComponent)){throw ('缺失组件错误未指出文件：'+$installComponent)}
        }finally{Move-Item -LiteralPath $installRemoved -Destination $installTarget}
    }
    $installInput = (New-TestFixtures).Input
    $installInputHash=(Get-FileHash -LiteralPath $installInput).Hash
    $installEqp=Join-Path $installAlternate 'Preferences\Times+Symbol 12.eqp'
    $installOutput=Join-Path $installTask 'forward.docx';$installReport=Join-Path $installTask 'forward.json'
    $installForward=Invoke-InstallationTask $installExe @('--worker-convert',$installInput,$installOutput,$installReport,$installEqp,(Join-Path $installTask 'stop.flag'),$installAlternate) 'forward-alternate'
    $installResult=Get-Content -LiteralPath $installReport -Raw|ConvertFrom-Json
    if($installResult.Status -ne 'COMPLETE' -or $installResult.Succeeded -ne 5 -or $installForward -notmatch 'OWNED_WORD_QUIT_PASSED'){throw '非默认目录正向转换失败。'}
    $installAudit=(& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/audit_docx.py') $installOutput)|ConvertFrom-Json
    if($LASTEXITCODE -ne 0 -or $installAudit.MathType -ne 5){throw '非默认目录结果独立检查失败。'}
    $installReverseOutput=Join-Path $installTask 'reverse.docx';$installReverseReport=Join-Path $installTask 'reverse.json'
    $null=Invoke-InstallationTask $installExe @('--worker-reverse',$installOutput,$installReverseOutput,$installReverseReport,'bracket',(Join-Path $installTask 'stop.flag'),$installAlternate) 'reverse-alternate'
    if((Get-Content -LiteralPath $installReverseReport -Raw|ConvertFrom-Json).Status -ne 'COMPLETE'){throw '非默认目录反向转换失败。'}
    $installCliOutput=Join-Path $installTask 'cli-auto.docx';$installCliReport=Join-Path $installTask 'cli-auto.json'
    $null=Invoke-InstallationTask (Join-Path $ProjectRoot 'build\bin\docx-math-convert.exe') @('--convert',$installInput,$installCliOutput,$installCliReport,$installEqp) 'cli-auto-discovery'
    if((Get-Content -LiteralPath $installCliReport -Raw|ConvertFrom-Json).Status -ne 'COMPLETE'){throw '无目录参数的命令行自动发现失败。'}
    # 损坏私有模板必须导致失败，用以确认没有偷偷使用原安装的固定模板路径。
    $installTemplate=Join-Path $installAlternate 'Office Support\BlankEqn.doc'
    [IO.File]::WriteAllBytes($installTemplate,[byte[]]@())
    $installBadOutput=Join-Path $installTask 'bad-template.docx';$installBadReport=Join-Path $installTask 'bad-template.json'
    $null=Invoke-InstallationTask $installExe @('--worker-convert',$installInput,$installBadOutput,$installBadReport,$installEqp,(Join-Path $installTask 'stop.flag'),$installAlternate) 'alternate-template-required' 1
    if((Test-Path -LiteralPath $installBadOutput) -or (Test-Path -LiteralPath $installBadReport) -or @(Get-ChildItem -LiteralPath $installTask -Filter '*.work-*').Count){throw '损坏模板任务发布了未知结果或遗留工作副本。'}
    foreach($installComponent in $installComponents){if((Get-FileHash -LiteralPath (Join-Path $installOriginal $installComponent)).Hash -ne $installHashes[$installComponent]){throw '真实 MathType 安装被改变。'}}
    if((Get-FileHash -LiteralPath $installInput).Hash -ne $installInputHash){throw '原始测试文档被改变。'}
    $installEvidence.Checks=@{IsolatedRegistry=$true;DiscoveryCases=11;MissingComponents=5;UnicodeAndSpaces=$true;DifferentDrive=([IO.Path]::GetPathRoot($installOriginal) -ne [IO.Path]::GetPathRoot($installAlternate));ExplicitForwardAndReverse=$true;DefaultCliDiscovery=$true;AlternateTemplateRequired=$true;OriginalInstallationUnchanged=$true;SourceDocumentUnchanged=$true}
    $installEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $installExe).Hash
    $installEvidence.CliExecutableSHA256=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot 'build\bin\docx-math-convert.exe')).Hash
    $installEvidence.ProbeExecutableSHA256=(Get-FileHash -LiteralPath $installProbe).Hash
    $installEvidence.SourceSHA256=Get-TestSourceHashes
    $installEvidence.Status='PASSED';Write-Output 'INSTALLATION_ACCEPTANCE_PASSED'
}finally{
    # 组件副本属于本轮临时文件；日志和 DOCX 保留为验证证据。
    $installResolved=[IO.Path]::GetFullPath($installAlternate)
    if(-not $installResolved.StartsWith($installRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '拒绝清理临时目录外的组件副本。'}
    Remove-Item -LiteralPath $installResolved -Recurse -Force
    $installParent=Split-Path -Parent $installResolved
    if(@(Get-ChildItem -LiteralPath $installParent).Count -eq 0){Remove-Item -LiteralPath $installParent}
    $installEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $installTask 'installation-results.json') -Encoding utf8
    $installDestination=[IO.Path]::GetFullPath($installPublished)
    if(-not $installDestination.StartsWith([IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output'))+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '验证材料目标越界。'}
    Move-Item -LiteralPath $installTask -Destination $installDestination
    foreach($installJson in Get-ChildItem -LiteralPath $installDestination -Filter '*.json'){
        $installData=Get-Content -LiteralPath $installJson.FullName -Raw|ConvertFrom-Json
        foreach($installField in @('Input','Output','Preferences')){if($installData.PSObject.Properties.Name -contains $installField){$installData.$installField=$installData.$installField.Replace($installTask,$installDestination)}}
        $installData|ConvertTo-Json -Depth 12|Set-Content -LiteralPath $installJson.FullName -Encoding utf8
    }
    $installEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'installation-results.json') -Encoding utf8
}
