#requires -Version 7.0
# 只推送白名单源码；发布程序包由维护者在 GitHub 网页操作。
param([string]$Message=('更新源码 '+(Get-Date -Format 'yyyy-MM-dd HH:mm')),[switch]$CheckOnly)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$repository='https://github.com/yuxmic/DocxMathTypeBatch.git'
$publicName='yuxmic'
$publicEmail='159784959+yuxmic@users.noreply.github.com'

function Invoke-ProjectGit {
    & git -C $root @args
    if($LASTEXITCODE -ne 0){throw ('Git 操作失败：'+$args[0]+'。未继续推送。')}
}

if((Invoke-ProjectGit rev-parse --show-toplevel).Replace('/','\') -ne $root.Replace('/','\')){throw '请将脚本放在 DocxMathTypeBatch 仓库根目录。'}
if((Invoke-ProjectGit branch --show-current) -ne 'main'){throw '请先切回 main 分支。'}
foreach($direction in @(@('remote','get-url','--all','origin'),@('remote','get-url','--push','--all','origin'))){
    $urls=@(Invoke-ProjectGit @direction)
    if($urls.Count -ne 1 -or $urls[0] -cne $repository){throw 'origin 必须只指向公开的 DocxMathTypeBatch 仓库。'}
}
& git -C $root diff --cached --quiet
if($LASTEXITCODE -eq 1){throw '暂存区已有改动，请先取消暂存后再使用一键更新；文件内容不会丢失。'}
if($LASTEXITCODE -ne 0){throw '无法检查暂存区。'}
if(-not $CheckOnly){Invoke-ProjectGit fetch --no-tags origin main}
$head=Invoke-ProjectGit rev-parse HEAD
$baseline=Invoke-ProjectGit rev-parse refs/remotes/origin/main
if($head -ne $baseline){throw '本地与 origin/main 不一致。请先同步远程改动，或处理未推送提交；脚本不会自动合并、重置或强推。'}
$names=@(Get-Content -LiteralPath (Join-Path $root 'upload-files.txt') -Encoding utf8|Where-Object {$_ -and -not $_.StartsWith('#')})
$previous=@(Invoke-ProjectGit show HEAD:upload-files.txt|Where-Object {$_ -and -not $_.StartsWith('#')})
$removed=@($previous|Where-Object {$_ -cnotin $names})

$sandbox=Join-Path $root ('临时文件夹/github-update-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $sandbox|Out-Null
$originalIndex=$env:GIT_INDEX_FILE
try{
    # 独立索引只收集清单文件，不改用户暂存区；移出清单的旧文件仅取消跟踪。
    $env:GIT_INDEX_FILE=Join-Path $sandbox 'index'
    Invoke-ProjectGit read-tree HEAD
    if($removed.Count){Invoke-ProjectGit rm --cached --ignore-unmatch -- @removed|Out-Null}
    & (Join-Path $root 'check-upload.ps1')
    Invoke-ProjectGit add -- @names
    $tree=Invoke-ProjectGit write-tree
    $committedNames=@(Invoke-ProjectGit -c core.quotePath=false ls-tree -r --name-only $tree)
    if(Compare-Object -ReferenceObject $names -DifferenceObject $committedNames -CaseSensitive){throw '实际提交文件与白名单不一致。'}
    # 从 Git 对象回读实际上传字节，过滤器或暂存内容也必须通过同一隐私检查。
    $archive=Join-Path $sandbox 'source.zip'
    $snapshot=Join-Path $sandbox 'source'
    Invoke-ProjectGit archive --format=zip "--output=$archive" $tree
    [IO.Compression.ZipFile]::ExtractToDirectory($archive,$snapshot)
    $snapshotNames=@(Get-Content -LiteralPath (Join-Path $snapshot 'upload-files.txt') -Encoding utf8|Where-Object {$_ -and -not $_.StartsWith('#')})
    if(Compare-Object -ReferenceObject $names -DifferenceObject $snapshotNames -CaseSensitive){throw 'Git 中的白名单与已检查清单不一致。'}
    & (Join-Path $root 'check-upload.ps1') -ProjectDirectory $snapshot -PrivacyDirectory $root
    $metadata=Join-Path $sandbox 'metadata'
    New-Item -ItemType Directory -Path $metadata|Out-Null
    'commit.txt'|Set-Content -LiteralPath (Join-Path $metadata 'upload-files.txt') -Encoding utf8
    $Message|Set-Content -LiteralPath (Join-Path $metadata 'commit.txt') -Encoding utf8
    & (Join-Path $root 'check-upload.ps1') -ProjectDirectory $metadata -PrivacyDirectory $root
    Write-Output ('公开文件：'+$names.Count+' 个；移出白名单：'+$removed.Count+' 个；目标：'+$repository)
    Invoke-ProjectGit diff --cached --stat
    if($CheckOnly){Write-Output 'CHECK_ONLY_PASSED：未联网、未提交、未推送；远程状态以本地缓存为准。';return}
    if($tree -eq (Invoke-ProjectGit rev-parse 'HEAD^{tree}')){Write-Output '源码没有变化，无需推送。';return}
    # 直接用已检查的树创建提交，避免提交钩子再次修改文件；固定公开身份。
    $identityVariables=@('GIT_AUTHOR_NAME','GIT_AUTHOR_EMAIL','GIT_COMMITTER_NAME','GIT_COMMITTER_EMAIL')
    $savedIdentity=@{}
    try{
        foreach($variable in $identityVariables){
            $savedIdentity[$variable]=[Environment]::GetEnvironmentVariable($variable,'Process')
            $value=if($variable.EndsWith('_EMAIL')){$publicEmail}else{$publicName}
            [Environment]::SetEnvironmentVariable($variable,$value,'Process')
        }
        $commit=Invoke-ProjectGit commit-tree $tree -p $head -m $Message
    }finally{
        foreach($variable in $identityVariables){
            if($null -eq $savedIdentity[$variable]){Remove-Item -LiteralPath ('Env:'+$variable)}
            else{[Environment]::SetEnvironmentVariable($variable,$savedIdentity[$variable],'Process')}
        }
    }
    # 只发送本次已检查的提交与 main，不附带标签；失败时本地分支保持原位，修复后可重试。
    Invoke-ProjectGit push --no-follow-tags --recurse-submodules=no $repository ($commit+':refs/heads/main')
    Invoke-ProjectGit update-ref refs/heads/main $commit $head
    $env:GIT_INDEX_FILE=$originalIndex
    Invoke-ProjectGit read-tree $commit
    Invoke-ProjectGit update-ref refs/remotes/origin/main $commit $baseline
    Write-Output ('GITHUB_SOURCE_UPDATED：'+$commit)
    Write-Output '源码更新完成。发布程序请打开 https://github.com/yuxmic/DocxMathTypeBatch/releases/new'
}finally{
    $env:GIT_INDEX_FILE=$originalIndex
    # 只清理本次创建的临时副本，删除前确认目标位于项目临时目录。
    $resolved=[IO.Path]::GetFullPath($sandbox)
    $temporaryRoot=[IO.Path]::GetFullPath((Join-Path $root '临时文件夹'))+[IO.Path]::DirectorySeparatorChar
    if(-not $resolved.StartsWith($temporaryRoot,[StringComparison]::OrdinalIgnoreCase)){throw '临时目录清理越界。'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
