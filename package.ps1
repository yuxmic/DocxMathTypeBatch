#requires -Version 7.0
# 发布只读取当前构建和说明原稿；不打包源码、本地样本或测试报告。
$ErrorActionPreference='Stop'
& (Join-Path $PSScriptRoot 'check-upload.ps1')
$exe=Join-Path $PSScriptRoot 'build/bin/DocxMathTypeBatch.exe'
$guide=Join-Path $PSScriptRoot 'docs/简要使用说明.txt'
$build=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'build/bin/gui-build-info.json') -Raw|ConvertFrom-Json
$exeHash=(Get-FileHash -LiteralPath $exe).Hash
if($build.ExecutableSHA256 -ne $exeHash){throw '请先构建当前 GUI 成品。'}
# 发布二进制和说明也不得带入当前机器的路径或身份。
foreach($file in @($exe,$guide)){
    $bytes=[IO.File]::ReadAllBytes($file)
    foreach($encoding in @([Text.Encoding]::UTF8,[Text.Encoding]::Unicode)){
        $text=$encoding.GetString($bytes)
        foreach($privateValue in @($PSScriptRoot,$env:USERPROFILE,$env:USERNAME)){
            if($privateValue -and $text.Contains($privateValue)){throw '发布文件含本机路径或身份。'}
        }
    }
}
$allowed=@('ole32.dll','oleaut32.dll','user32.dll','xmllite.dll','crypt32.dll','comctl32.dll','comdlg32.dll','shell32.dll','gdi32.dll','gdiplus.dll','d2d1.dll','dwrite.dll','advapi32.dll','uxtheme.dll','kernel32.dll')
if($build.ClrDirectory -ne 0 -or @($build.ImportedDlls|Where-Object {$_ -notin $allowed}).Count -or (Get-Item -LiteralPath $exe).Length -gt 1MB){throw 'PE 依赖或大小不符合发布约定。'}
$name='DocxMathTypeBatch-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,6)
$folder=Join-Path $PSScriptRoot ('dist/'+$name)
New-Item -ItemType Directory -Path $folder -Force|Out-Null
$files=@('DocxMathTypeBatch.exe','简要使用说明.txt')
Copy-Item -LiteralPath $exe,$guide -Destination $folder
$zip=$folder+'.zip'
[IO.Compression.ZipFile]::CreateFromDirectory($folder,$zip)
# ZIP 回读校验只允许两个明确的发布文件。
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
[ordered]@{Status='PACKAGED';Folder=('dist/'+$name);Zip=('dist/'+$name+'.zip');ExeSHA256=$exeHash;ZipSHA256=(Get-FileHash -LiteralPath $zip).Hash;SourceSHA256=$sourceHashes;Files=$files;ArchiveReadBackVerified=$true;SampleTestsRun=$false}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath ($folder+'.json') -Encoding utf8
Write-Output ('PACKAGE_CREATED: dist/'+$name+'.zip')
