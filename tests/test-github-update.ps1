#requires -Version 7.0
# 在临时本地裸仓库验证推送与隐私闸门，全部操作都不连接 GitHub。
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$sandbox=Join-Path $TestOutputRoot ('github-update-'+[Guid]::NewGuid().ToString('N'))
$project=Join-Path $sandbox 'project'
$remote=Join-Path $sandbox 'remote.git'
$gitExecutable=(Get-Command git -CommandType Application|Select-Object -First 1).Source
$updateTestState=@{PushCount=0;RejectPush=$false}
function Invoke-FixtureGit {
    & $gitExecutable @args
    if($LASTEXITCODE -ne 0){throw '本地 Git 测试准备或断言失败。'}
}
function Assert-Update([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
function Assert-Rejected([scriptblock]$Action){
    $rejected=$false
    try{& $Action|Out-Null}catch{$rejected=$true}
    Assert-Update $rejected '本应拒绝的更新未被拒绝。'
}
# 仅在测试作用域将固定公开地址映射到本地裸仓库，其余调用均执行真正的 Git。
function git {
    $Arguments=@($args)
    if('get-url' -in $Arguments){$global:LASTEXITCODE=0;return 'https://github.com/yuxmic/DocxMathTypeBatch.git'}
    if('push' -in $Arguments){
        ++$updateTestState.PushCount
        if($updateTestState.RejectPush){throw '模拟网络推送失败。'}
    }
    $Arguments=@($Arguments|ForEach-Object {if($_ -eq 'https://github.com/yuxmic/DocxMathTypeBatch.git'){$remote}else{$_}})
    & $gitExecutable @Arguments
    $global:LASTEXITCODE=$LASTEXITCODE
}
New-Item -ItemType Directory -Path $project|Out-Null
try{
    Invoke-FixtureGit init --bare --quiet $remote
    Invoke-FixtureGit -C $project init --quiet --initial-branch=main
    Invoke-FixtureGit -C $project config core.autocrlf false
    Invoke-FixtureGit -C $project config user.name fixture
    Invoke-FixtureGit -C $project config user.email ('private'+'@'+'example.invalid')
    $names=@('.gitignore','upload-files.txt','check-upload.ps1','update-github.ps1','public.txt')
    Copy-Item -LiteralPath (Join-Path $ProjectRoot 'check-upload.ps1'),(Join-Path $ProjectRoot 'update-github.ps1') -Destination $project
    $names|Set-Content -LiteralPath (Join-Path $project 'upload-files.txt') -Encoding utf8
    @('*')+@($names|ForEach-Object {'!/'+$_})|Set-Content -LiteralPath (Join-Path $project '.gitignore') -Encoding utf8
    'baseline'|Set-Content -LiteralPath (Join-Path $project 'public.txt') -Encoding utf8
    Invoke-FixtureGit -C $project add -- @names
    Invoke-FixtureGit -C $project commit --quiet -m baseline
    Invoke-FixtureGit -C $project remote add origin $remote
    Invoke-FixtureGit -C $project push --quiet -u origin main
    $baseline=Invoke-FixtureGit -C $project rev-parse HEAD
    $entry=Join-Path $project 'update-github.ps1'
    New-Item -ItemType Directory -Path (Join-Path $project '.local/samples') -Force|Out-Null
    'private fixture'|Set-Content -LiteralPath (Join-Path $project '.local/samples/internal-paper.docx')
    'private fixture'|Set-Content -LiteralPath (Join-Path $project 'private.txt')
    'public change'|Set-Content -LiteralPath (Join-Path $project 'public.txt') -Encoding utf8
    & $entry -CheckOnly|Out-Null
    Assert-Update ((Invoke-FixtureGit -C $project rev-parse HEAD) -eq $baseline) '预检改变了本地提交。'
    Assert-Update ((Invoke-FixtureGit -C $project write-tree) -eq (Invoke-FixtureGit -C $project rev-parse 'HEAD^{tree}')) '预检改变了暂存区。'
    Assert-Update ($updateTestState.PushCount -eq 0) '预检执行了推送。'
    foreach($value in @(('ghp_'+('a'*30)),('Q'+':\'+'private\paper.docx'),('personal'+'@'+'example.invalid'),'internal-paper.docx')){
        $value|Set-Content -LiteralPath (Join-Path $project 'public.txt') -Encoding utf8
        Assert-Rejected {& $entry -CheckOnly}
    }
    'public change'|Set-Content -LiteralPath (Join-Path $project 'public.txt') -Encoding utf8
    Assert-Rejected {& $entry -Message ('ghp_'+('b'*30)) -CheckOnly}
    # 合成 clean 过滤器向 Git 对象注入密钥，验证扫描的是实际提交字节。
    $filterPath=Join-Path $sandbox 'filter.ps1'
    @'
# 消费输入后只输出合成密钥，不读取任何本机配置。
$null=[Console]::In.ReadToEnd()
[Console]::Write('ghp_'+('c'*30))
'@|Set-Content -LiteralPath $filterPath -Encoding utf8
    'public.txt filter=fixture'|Set-Content -LiteralPath (Join-Path $project '.git/info/attributes')
    Invoke-FixtureGit -C $project config filter.fixture.clean ('pwsh -NoProfile -File "'+$filterPath+'"')
    Assert-Rejected {& $entry -CheckOnly}
    ''|Set-Content -LiteralPath (Join-Path $project '.git/info/attributes')
    Invoke-FixtureGit -C $project config --remove-section filter.fixture
    Invoke-FixtureGit -C $project add --force -- private.txt
    Assert-Rejected {& $entry -CheckOnly}
    Invoke-FixtureGit -C $project reset --quiet HEAD -- private.txt
    $updateTestState.RejectPush=$true
    Assert-Rejected {& $entry -Message 'public update'}
    Assert-Update ((Invoke-FixtureGit -C $project rev-parse HEAD) -eq $baseline) '推送失败改变了本地分支。'
    Assert-Update ((Invoke-FixtureGit -C $project write-tree) -eq (Invoke-FixtureGit -C $project rev-parse 'HEAD^{tree}')) '推送失败改变了暂存区。'
    $updateTestState.RejectPush=$false
    & $entry -Message 'public update'|Out-Null
    $updated=Invoke-FixtureGit -C $project rev-parse HEAD
    Assert-Update ($updated -ne $baseline -and $updated -eq (Invoke-FixtureGit --git-dir=$remote rev-parse refs/heads/main)) '本地与远程更新结果不同。'
    $actual=@(Invoke-FixtureGit --git-dir=$remote ls-tree -r --name-only main)
    Assert-Update (-not (Compare-Object $names $actual)) '推送包含白名单之外的文件。'
    $identity=@(Invoke-FixtureGit -C $project log -1 '--format=%an%n%ae%n%cn%n%ce')
    Assert-Update (($identity -join '|') -eq 'yuxmic|159784959+yuxmic@users.noreply.github.com|yuxmic|159784959+yuxmic@users.noreply.github.com') '提交泄露了本地配置的身份。'
    $pushes=$updateTestState.PushCount
    & $entry|Out-Null
    Assert-Update ($updateTestState.PushCount -eq $pushes) '无改动时仍执行推送。'
    # 移出清单只从远程取消跟踪，保留本机文件。
    $names=@($names|Where-Object {$_ -ne 'public.txt'})
    $names|Set-Content -LiteralPath (Join-Path $project 'upload-files.txt') -Encoding utf8
    & $entry -Message 'remove public fixture'|Out-Null
    Assert-Update (Test-Path -LiteralPath (Join-Path $project 'public.txt')) '取消跟踪删除了本机文件。'
    $actual=@(Invoke-FixtureGit --git-dir=$remote ls-tree -r --name-only main)
    Assert-Update (-not (Compare-Object $names $actual)) '移出白名单的文件仍在远程。'
    Invoke-FixtureGit -C $project commit --quiet --allow-empty -m 'unreviewed local history'
    Assert-Rejected {& $entry -CheckOnly}
}finally{
    # 本次测试数据均为临时生成，清理前核实绝对路径。
    $resolved=[IO.Path]::GetFullPath($sandbox)
    if(-not $resolved.StartsWith($TestOutputRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '测试清理越界。'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
Write-Output 'GITHUB_UPDATE_TESTS_PASSED'
