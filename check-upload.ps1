#requires -Version 7.0
# 上传检查只读取明确的公开清单；私有样本和运行报告不作为检查或导出的候选文件。
param([string]$ProjectDirectory=$PSScriptRoot,[switch]$Export,[string]$PrivacyDirectory=$ProjectDirectory)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath($ProjectDirectory)
$manifest=Join-Path $root 'upload-files.txt'
$names=@(Get-Content -LiteralPath $manifest -Encoding utf8 | Where-Object {$_ -and -not $_.StartsWith('#')})
if($names.Count -ne @($names|Sort-Object -Unique).Count){throw '公开清单包含重复路径。'}
$issues=[Collections.Generic.List[string]]::new()
$binaryNames=@('assets/app_icon.png','assets/app_icon.ico')
$privateValues=@($env:USERPROFILE,$env:USERNAME,[IO.Path]::GetFullPath($PrivacyDirectory))|Where-Object {$_ -and $_.Length -ge 4}
$privateFolder=Join-Path $PrivacyDirectory '.local/samples'
if(Test-Path -LiteralPath $privateFolder){$privateValues+=@(Get-ChildItem -LiteralPath $privateFolder -File|ForEach-Object {$_.Name})}
$hashes=[ordered]@{}
foreach($name in $names){
    if($name -match '(^/|\\|(^|/)\.\.(/|$)|^\.|^(build|dist|临时文件夹)/|^tests/output/)' -and $name -notin @('.gitignore','.gitattributes')){throw '公开清单包含禁止路径。'}
    $path=[IO.Path]::GetFullPath((Join-Path $root $name))
    if(-not $path.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '公开文件路径越界。'}
    $entry=Get-Item -LiteralPath $path
    # 检查沿途目录，避免符号链接把本地文件带入公开导出。
    $ancestor=$entry
    while($ancestor.FullName -ne $root){
        if($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint){throw ('公开文件经过链接：'+$name)}
        $ancestor=Get-Item -LiteralPath (Split-Path -Parent $ancestor.FullName)
    }
    if($entry.PSIsContainer){throw ('清单项必须是文件：'+$name)}
    $bytes=[IO.File]::ReadAllBytes($path)
    $hashes[$name]=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    if($name -notin $binaryNames -and $entry.Extension -notin @('.ps1','.cmd','.py','.cpp','.c','.h','.rc','.md','.txt','.json','.yaml') -and $name -notin @('.gitignore','.gitattributes','LICENSE')){throw ('禁止上传该类型：'+$name)}
    $text=[Text.Encoding]::UTF8.GetString($bytes)
    if($name -notin $binaryNames){
        # 公开文本统一 UTF-8；拒绝空字节，避免其他编码绕过路径检查。
        $text=[Text.UTF8Encoding]::new($false,$true).GetString($bytes)
        if($text.Contains([char]0)){throw ('公开文本必须为 UTF-8：'+$name)}
    }
    foreach($value in $privateValues){
        # 用户名按独立词匹配；路径和私人文件名按完整字面量匹配。
        $pattern=[regex]::Escape($value)
        if($value -eq $env:USERNAME){$pattern='(?<![\w])'+$pattern+'(?![\w])'}
        if([regex]::IsMatch($text,$pattern,'IgnoreCase')){$issues.Add($name+': 含本机身份、路径或私人样本名称')}
        if([Text.Encoding]::Unicode.GetString($bytes).Contains($value)){$issues.Add($name+': 二进制中含本机信息')}
    }
    if($name -notin $binaryNames){
        $lineNumber=0
        foreach($line in ($text -split '\r?\n')){
            ++$lineNumber
            if($line -match '(?<![A-Za-z0-9_])[A-Za-z]:[\\/]' -or $line -match '(?:/Users/|/home/)[A-Za-z0-9_.-]+/'){$issues.Add($name+':'+$lineNumber+': 含绝对文件路径')}
            if($line -match '(?:gh[pousr]_[A-Za-z0-9_]{20,}|github_pat_[A-Za-z0-9_]{20,}|sk-[A-Za-z0-9_-]{20,}|AKIA[0-9A-Z]{16}|-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----)'){$issues.Add($name+':'+$lineNumber+': 疑似密钥')}
            foreach($email in [regex]::Matches($line,'[A-Za-z0-9_.+%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')){
                if($email.Value -notmatch '@(?:users\.)?noreply\.github\.com$'){$issues.Add($name+':'+$lineNumber+': 含非 GitHub 隐私邮箱')}
            }
        }
    }
}
if(Test-Path -LiteralPath (Join-Path $root '.git')){
    $tracked=@(& git -C $root -c core.quotePath=false ls-files)
    if($LASTEXITCODE -ne 0){throw '无法读取 Git 已跟踪文件。'}
    foreach($trackedName in $tracked){if($trackedName -cnotin $names){$issues.Add('Git 已跟踪公开清单之外的文件')}}
}
if($issues.Count){throw ('上传检查未通过：'+[Environment]::NewLine+($issues -join [Environment]::NewLine))}
Write-Output ('UPLOAD_CHECK_PASSED: '+$names.Count+' 个公开文件')
if($Export){
    $dist=Join-Path $root 'dist';New-Item -ItemType Directory -Path $dist -Force|Out-Null
    $zipPath=Join-Path $dist ('DocxMathTypeBatch-source-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,6)+'.zip')
    $zip=[IO.Compression.ZipFile]::Open($zipPath,[IO.Compression.ZipArchiveMode]::Create)
    try{foreach($name in $names){[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,(Join-Path $root $name),$name)|Out-Null}}finally{$zip.Dispose()}
    # 从归档回读文件名和字节，保证实际导出与已检查的公开清单一致。
    $zip=[IO.Compression.ZipFile]::OpenRead($zipPath)
    try{
        if($zip.Entries.Count -ne $names.Count){throw '源码包文件数错误。'}
        foreach($entry in $zip.Entries){
            if($entry.FullName -cnotin $names){throw '源码包包含清单之外的文件。'}
            $stream=$entry.Open();$memory=[IO.MemoryStream]::new()
            try{$stream.CopyTo($memory);$hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($memory.ToArray()))}finally{$stream.Dispose();$memory.Dispose()}
            if($hash -ne $hashes[$entry.FullName]){throw '源码包与检查时文件不同。'}
        }
    }finally{$zip.Dispose()}
    Write-Output ('SOURCE_EXPORT_CREATED: dist/'+[IO.Path]::GetFileName($zipPath))
}
