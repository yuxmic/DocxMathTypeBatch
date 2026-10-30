#requires -Version 7.0
# GUI 包与包含 EXE 的 Skill 包分别打包；不收集本机配置、样本或日志。
param(
    [ValidatePattern('^v[0-9]+\.[0-9]+(?:\.[0-9]+)?$')][string]$Version='v0.1.3',
    [ValidatePattern('^v[0-9]+\.[0-9]+(?:\.[0-9]+)?$')][string]$SkillVersion='v0.2',
    [ValidateSet('All','Exe','Skill')][string]$Target='All'
)
$ErrorActionPreference='Stop'
& (Join-Path $PSScriptRoot 'check-upload.ps1')
# 两类包共用程序身份、构建哈希、依赖和隐私检查，Skill 单独打包也必须执行。
$exe=Join-Path $PSScriptRoot 'build/bin/DocxMathTypeBatch.exe'
$guide=Join-Path $PSScriptRoot 'docs/简要使用说明.txt'
$license=Join-Path $PSScriptRoot 'LICENSE'
$build=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'build/bin/gui-build-info.json') -Raw|ConvertFrom-Json
$exeHash=(Get-FileHash -LiteralPath $exe).Hash
if($build.ExecutableSHA256 -ne $exeHash){throw '请先构建当前 GUI 成品。'}
$product=[Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
if($product.ProductName -ne 'DocxMathTypeBatch' -or $product.ProductVersion -ne $Version.Substring(1)){throw '程序资源版本与发布版本不一致，请核对 app_version.h 后重新构建。'}
# 发布二进制和说明也不得带入当前机器的路径或身份。
foreach($file in @($exe,$guide,$license)){
    $bytes=[IO.File]::ReadAllBytes($file)
    foreach($encoding in @([Text.Encoding]::UTF8,[Text.Encoding]::Unicode)){
        $text=$encoding.GetString($bytes)
        foreach($privateValue in @($PSScriptRoot,$env:USERPROFILE,$env:USERNAME)){
            if($privateValue -and $text.Contains($privateValue)){throw '发布文件含本机路径或身份。'}
        }
    }
}
$allowed=@('ole32.dll','oleaut32.dll','user32.dll','xmllite.dll','crypt32.dll','comctl32.dll','comdlg32.dll','shell32.dll','gdi32.dll','gdiplus.dll','d2d1.dll','dwrite.dll','advapi32.dll','uxtheme.dll','kernel32.dll')
# 功能扩展后仍限制为小型原生成品，不引入额外运行库。
if($build.ClrDirectory -ne 0 -or @($build.ImportedDlls|Where-Object {$_ -notin $allowed}).Count -or (Get-Item -LiteralPath $exe).Length -gt 2MB){throw 'PE 依赖或大小不符合发布约定。'}
if($Target -in @('All','Exe')){
$name='DocxMathTypeBatch-'+$Version+'-win-x64'
$folder=Join-Path $PSScriptRoot ('dist/'+$name)
$zip=$folder+'.zip'
# 已打好的版本不静默覆盖，避免同一版本对应不同成品。
if((Test-Path -LiteralPath $folder) -or (Test-Path -LiteralPath $zip) -or (Test-Path -LiteralPath ($folder+'.json'))){throw '该版本包已存在，请使用新版本号，或先将旧包移到本机临时目录。'}
New-Item -ItemType Directory -Path $folder -Force|Out-Null
$files=@('DocxMathTypeBatch.exe','简要使用说明.txt','LICENSE')
Copy-Item -LiteralPath $exe,$guide,$license -Destination $folder
[IO.Compression.ZipFile]::CreateFromDirectory($folder,$zip)
# ZIP 回读校验只允许程序、说明和许可文本。
$archive=[IO.Compression.ZipFile]::OpenRead($zip)
try{
    if($archive.Entries.Count -ne $files.Count){throw '发布包内容数量错误。'}
    foreach($entry in $archive.Entries){
        if($entry.FullName -cnotin $files){throw '发布包出现非预期文件。'}
        $stream=$entry.Open();$memory=[IO.MemoryStream]::new()
        try{$stream.CopyTo($memory);$hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($memory.ToArray()))}finally{$stream.Dispose();$memory.Dispose()}
        if($hash -ne (Get-FileHash -LiteralPath (Join-Path $folder $entry.FullName)).Hash){throw 'ZIP 内容哈希不一致。'}
    }
}finally{$archive.Dispose()}
$sourceHashes=[ordered]@{}
foreach($sourceName in Get-Content -LiteralPath (Join-Path $PSScriptRoot 'upload-files.txt')|Where-Object {$_ -and -not $_.StartsWith('#')}){$sourceHashes[$sourceName]=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot $sourceName)).Hash}
[ordered]@{Status='PACKAGED';Version=$Version;License='MIT';Folder=('dist/'+$name);Zip=('dist/'+$name+'.zip');ExeSHA256=$exeHash;ZipSHA256=(Get-FileHash -LiteralPath $zip).Hash;SourceSHA256=$sourceHashes;Files=$files;ArchiveReadBackVerified=$true;SampleTestsRun=$false}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath ($folder+'.json') -Encoding utf8
Write-Output ('PACKAGE_CREATED: dist/'+$name+'.zip')
}
if($Target -in @('All','Skill')){
    $skillName='docx-latex-to-mathtype'
    $prefix='skills/'+$skillName+'/'
    $skillSource=Join-Path $PSScriptRoot ('skills/'+$skillName)
    $metadata=[IO.File]::ReadAllText((Join-Path $skillSource 'SKILL.md'))
    $declared=[regex]::Match($metadata,'(?m)^  version: "([^"]+)"\s*$')
    if(-not $declared.Success -or $declared.Groups[1].Value -ne $SkillVersion.Substring(1)){throw 'Skill 声明与发布版本不一致。'}
    $skillFiles=@(Get-Content -LiteralPath (Join-Path $PSScriptRoot 'upload-files.txt')|Where-Object{$_ -and $_.StartsWith($prefix)}|ForEach-Object{$_.Substring($prefix.Length)})
    if(-not $skillFiles -or 'SKILL.md' -notin $skillFiles){throw '公开清单没有完整 Skill。'}
    $skillPackageName=$skillName+'-skill-'+$SkillVersion
    $skillFolder=Join-Path $PSScriptRoot ('dist/'+$skillPackageName)
    $skillZip=$skillFolder+'.zip'
    foreach($path in @($skillFolder,$skillZip,($skillFolder+'.json'))){if(Test-Path -LiteralPath $path){throw '该 Skill 版本包已存在，请选择新版本或先由用户归档旧包。'}}
    $contentRoot=Join-Path $skillFolder $skillName
    $expected=[ordered]@{}
    foreach($name in $skillFiles){
        $source=Join-Path $skillSource $name;$destination=Join-Path $contentRoot $name
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
        Copy-Item -LiteralPath $source -Destination $destination
        $expected[$skillName+'/'+$name]=(Get-FileHash -LiteralPath $source).Hash
    }
    # 始终注入当前已验证的构建结果，不打包源码目录中可能过时的程序副本。
    $runtimePath=Join-Path $contentRoot 'bin/DocxMathTypeBatch.exe'
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($runtimePath))
    Copy-Item -LiteralPath $exe -Destination $runtimePath
    $expected[$skillName+'/bin/DocxMathTypeBatch.exe']=$exeHash
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'LICENSE') -Destination (Join-Path $contentRoot 'LICENSE')
    $expected[$skillName+'/LICENSE']=(Get-FileHash -LiteralPath (Join-Path $contentRoot 'LICENSE')).Hash
    [IO.Compression.ZipFile]::CreateFromDirectory($skillFolder,$skillZip)
    # 归档回读只允许明确的公开 Skill 文件与许可，不收集运行生成物。
    $archive=[IO.Compression.ZipFile]::OpenRead($skillZip)
    try{
        if($archive.Entries.Count -ne $expected.Count){throw 'Skill 包文件数量不符。'}
        foreach($entry in $archive.Entries){
            if(-not $expected.Contains($entry.FullName)){throw 'Skill 包包含清单以外文件。'}
            $stream=$entry.Open();$memory=[IO.MemoryStream]::new()
            try{$stream.CopyTo($memory);$hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($memory.ToArray()))}finally{$stream.Dispose();$memory.Dispose()}
            if($hash -ne $expected[$entry.FullName]){throw 'Skill ZIP 回读字节不符。'}
        }
    }finally{$archive.Dispose()}
    [ordered]@{Status='PACKAGED';Version=$SkillVersion;BundledApplicationVersion=$Version;BundledExeSHA256=$exeHash;CliContractVersion=1;License='MIT';Zip=('dist/'+$skillPackageName+'.zip');ZipSHA256=(Get-FileHash -LiteralPath $skillZip).Hash;Files=@($expected.Keys);SourceSHA256=$expected;ArchiveReadBackVerified=$true}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath ($skillFolder+'.json') -Encoding utf8
    Write-Output ('SKILL_PACKAGE_CREATED: dist/'+$skillPackageName+'.zip')
}
