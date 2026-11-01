#requires -Version 7.0
# 真实 Word 外链样本：私有更新方式可恢复，正常转换不更新预览或修改应用全局选项。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$linkRoot=[IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output\work'))
$linkIdentity=[Guid]::NewGuid().ToString('N');$linkTask=Join-Path $linkRoot ('外链验收-'+$linkIdentity)
$linkPublished=Join-Path $ProjectRoot ('tests\output\外链验收-'+$linkIdentity)
New-Item -ItemType Directory -Path $linkTask|Out-Null
$linkEvidence=[ordered]@{Status='FAILED';CheckedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');Checks=@{};SampleDirectory=$linkPublished}
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class LinkWordIdentity{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);}
'@
function Invoke-LinkTask {
    param([string]$Executable,[string[]]$Arguments,[string]$Name)
    $linkProcess=Start-Process -FilePath $Executable -ArgumentList @($Arguments|ForEach-Object {'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $linkTask ($Name+'.log')) -RedirectStandardError (Join-Path $linkTask ($Name+'.errors')) -PassThru
    try{if(-not $linkProcess.WaitForExit(90000) -or $linkProcess.ExitCode -ne 0){throw ('外链验收失败：'+$Name)}}finally{$linkProcess.Dispose()}
}
function Read-LinkModes {
    param([string]$Path)
    $linkZip=[IO.Compression.ZipFile]::OpenRead($Path)
    try{
        $linkEntry=$linkZip.GetEntry('word/document.xml');$linkReader=[IO.StreamReader]::new($linkEntry.Open())
        try{$linkXml=[xml]$linkReader.ReadToEnd()}finally{$linkReader.Dispose()}
        return @($linkXml.SelectNodes('//*[local-name()="OLEObject" and @Type="Link"]')|ForEach-Object {$_.GetAttribute('UpdateMode')})
    }finally{$linkZip.Dispose()}
}
$linkWord=$null;$linkDocument=$null;$linkSourceDocument=$null;$linkOwned=$false;$linkOriginalOption=$null
try{
    if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target All}
    $linkPriorWord=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
    $linkWord=New-Object -ComObject Word.Application;$linkSourceDocument=$linkWord.Documents.Add()
    [uint32]$linkPid=0;[LinkWordIdentity]::GetWindowThreadProcessId([IntPtr]$linkSourceDocument.ActiveWindow.Hwnd,[ref]$linkPid)|Out-Null
    $linkOwned=$linkPid -ne 0 -and $linkPid -notin $linkPriorWord;if(-not $linkOwned){throw '拒绝使用既有 Word 构建外链样本。'}
    $linkWord.Visible=$false;$linkWord.DisplayAlerts=0;$linkWord.AutomationSecurity=3
    $linkOriginalOption=$linkWord.Options.UpdateLinksAtOpen;$linkWord.Options.UpdateLinksAtOpen=$true
    $linkSource=Join-Path $linkTask 'linked-source.docx';$linkInput=Join-Path $linkTask 'input.docx'
    $linkSourceDocument.Content.Text='原始外链内容';$linkSourceDocument.SaveAs2($linkSource,16);$linkSourceDocument.Close(0);$linkSourceDocument=$null
    $linkDocument=$linkWord.Documents.Add();$linkDocument.Content.Text='正文 $x_1+2$ 保留。' + "`r"+'外链对象：' + "`r"
    $linkRange=$linkDocument.Range($linkDocument.Content.End-1,$linkDocument.Content.End-1)
    $linkShape=$linkDocument.InlineShapes.AddOLEObject([Type]::Missing,$linkSource,$true,$false,[Type]::Missing,[Type]::Missing,[Type]::Missing,$linkRange)
    $linkShape.LinkFormat.AutoUpdate=$true;$linkDocument.SaveAs2($linkInput,16);$linkDocument.Close(0);$linkDocument=$null
    # 修改外链源，若转换错误触发自动更新，缓存预览便会变化。
    $linkSourceDocument=$linkWord.Documents.Open($linkSource,$false,$false,$false);$linkSourceDocument.Content.Text='修改后绝不应自动更新的外链内容';$linkSourceDocument.Save();$linkSourceDocument.Close(0);$linkSourceDocument=$null
    $linkOriginalModes=@(Read-LinkModes $linkInput);if($linkOriginalModes.Count -ne 1 -or $linkOriginalModes[0] -ne 'Always'){throw '真实自动更新 OLE 样本未生成。'}
    $linkInputHash=(Get-FileHash -LiteralPath $linkInput).Hash;$linkSourceHash=(Get-FileHash -LiteralPath $linkSource).Hash
    Invoke-LinkTask (Join-Path $ProjectRoot 'build\bin\native-api-probe.exe') @('--link-settings-self-test',$linkInput,(Join-Path $linkTask 'settings-roundtrip.docx')) 'settings-test'
    if(@(Read-LinkModes (Join-Path $linkTask 'settings-roundtrip.docx'))[0] -ne 'Always'){throw '私有外链设置未还原。'}
    $linkStyle = 'paragraph'
    $linkOutput=Join-Path $linkTask 'output.docx';$linkReport=Join-Path $linkTask 'report.json'
    Invoke-LinkTask (Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe') @('--worker-convert',$linkInput,$linkOutput,$linkReport,$linkStyle,(Join-Path $linkTask 'stop.flag')) 'conversion'
    $linkResult=Get-Content -LiteralPath $linkReport -Raw|ConvertFrom-Json
    if($linkResult.Succeeded -ne 1 -or -not $linkResult.ExistingMediaBytesPreserved -or @(Read-LinkModes $linkOutput)[0] -ne 'Always'){throw '外链缓存、转换结果或更新方式未保持。'}
    if(-not $linkWord.Options.UpdateLinksAtOpen -or (Get-FileHash -LiteralPath $linkInput).Hash -ne $linkInputHash -or (Get-FileHash -LiteralPath $linkSource).Hash -ne $linkSourceHash){throw 'Word 全局选项、输入或外链源被改变。'}
    & python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/audit_docx.py') $linkOutput | Set-Content -LiteralPath (Join-Path $linkTask 'audit.json') -Encoding utf8
    if($LASTEXITCODE -ne 0){throw '外链结果独立包检查失败。'}
    $linkEvidence.Checks=@{RealAutomaticWordLink=$true;PrivateOnCallAndRestoration=$true;CachedPreviewUnchanged=$true;FinalUpdateModeRestored=$true;WordGlobalOptionUnchanged=$true;InputAndLinkedSourceUnchanged=$true;IndependentAudit=$true}
    $linkEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe')).Hash
    $linkEvidence.ProbeExecutableSHA256=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot 'build\bin\native-api-probe.exe')).Hash
    $linkEvidence.SourceSHA256=Get-TestSourceHashes
    $linkEvidence.Status='PASSED';Write-Output 'LINK_UPDATES_ACCEPTANCE_PASSED'
}finally{
    if($linkOwned){if($linkDocument){$linkDocument.Close(0)};if($linkSourceDocument){$linkSourceDocument.Close(0)};if($null -ne $linkOriginalOption){$linkWord.Options.UpdateLinksAtOpen=$linkOriginalOption};$linkWord.Quit(0)}
    if($linkWord){[Runtime.InteropServices.Marshal]::FinalReleaseComObject($linkWord)|Out-Null}
    $linkEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $linkTask 'link-updates-results.json') -Encoding utf8
    $linkResolved=[IO.Path]::GetFullPath($linkTask);$linkDestination=[IO.Path]::GetFullPath($linkPublished)
    if(-not $linkResolved.StartsWith($linkRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or -not $linkDestination.StartsWith([IO.Path]::GetFullPath((Join-Path $ProjectRoot 'tests\output'))+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '外链验收材料搬迁越界。'}
    Move-Item -LiteralPath $linkResolved -Destination $linkDestination
    $linkEvidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'link-updates-results.json') -Encoding utf8
}
