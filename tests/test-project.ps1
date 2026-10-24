#requires -Version 7.0
# 检查公开脚本语法、Python 模块和隐私检查的拒绝行为，不启动 Office。
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
# 模拟中文系统的 GBK 控制台，公共入口须恢复正确的原生 UTF-8 JSON 管道。
[Console]::OutputEncoding=[Text.Encoding]::GetEncoding(936)
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$nativeJson=& python -B -X utf8 -c 'import json; print(json.dumps(dict(text=chr(0x4e2d)+chr(0x6587)+chr(34)), ensure_ascii=False))'
if($LASTEXITCODE -ne 0 -or ($nativeJson|ConvertFrom-Json).text -cne ('中文'+[char]34)){throw '原生 UTF-8 JSON 的中文或转义引号解码错误。'}
$names=Get-Content -LiteralPath (Join-Path $ProjectRoot 'upload-files.txt')|Where-Object {$_ -and -not $_.StartsWith('#')}
foreach($name in $names|Where-Object {$_ -like '*.ps1'}){
    $tokens=$null;$parseErrors=$null
    $null=[Management.Automation.Language.Parser]::ParseFile((Join-Path $ProjectRoot $name),[ref]$tokens,[ref]$parseErrors)
    if($parseErrors){throw ($name+': '+($parseErrors.Message -join '; '))}
}
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_project.py')
if($LASTEXITCODE -ne 0){throw 'Python 或合成样本检查失败。'}
$sandbox=Join-Path $TestOutputRoot ('upload-check-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $sandbox|Out-Null
try{
    # 用临时 Git 元数据核对真正的待上传文件，不在项目中初始化仓库。
    $gitDirectory=Join-Path $sandbox 'git-metadata'
    & git init --bare --quiet $gitDirectory
    if($LASTEXITCODE -ne 0){throw '临时 Git 初始化失败。'}
    $candidates=@(& git "--git-dir=$gitDirectory" "--work-tree=$ProjectRoot" -c core.quotePath=false ls-files --others --exclude-standard)
    if($LASTEXITCODE -ne 0 -or (Compare-Object -ReferenceObject @($names) -DifferenceObject $candidates)){throw 'Git 放行规则与公开清单不一致。'}
    'example.txt'|Set-Content -LiteralPath (Join-Path $sandbox 'upload-files.txt') -Encoding utf8
    $example=Join-Path $sandbox 'example.txt'
    'public fixture'|Set-Content -LiteralPath $example -Encoding utf8
    & (Join-Path $ProjectRoot 'check-upload.ps1') -ProjectDirectory $sandbox
    # 删除中的已跟踪文件允许退出清单，仍存在的清单外文件必须拒绝。
    & git -C $sandbox init --quiet
    $retired=Join-Path $sandbox 'retired.txt';'retired fixture'|Set-Content -LiteralPath $retired
    & git -C $sandbox add -- retired.txt
    if($LASTEXITCODE -ne 0){throw '清单删除回归准备失败。'}
    $rejected=$false
    try{& (Join-Path $ProjectRoot 'check-upload.ps1') -ProjectDirectory $sandbox|Out-Null}catch{$rejected=$true}
    if(-not $rejected){throw '上传检查未拒绝仍存在的清单外已跟踪文件。'}
    Remove-Item -LiteralPath $retired
    & (Join-Path $ProjectRoot 'check-upload.ps1') -ProjectDirectory $sandbox
    foreach($value in @(('Z'+':\'+'private\sample.docx'),('ghp_'+('a'*30)),'.local/paper.docx')){
        if($value -eq '.local/paper.docx'){$value|Set-Content -LiteralPath (Join-Path $sandbox 'upload-files.txt')}else{$value|Set-Content -LiteralPath $example}
        $rejected=$false
        try{& (Join-Path $ProjectRoot 'check-upload.ps1') -ProjectDirectory $sandbox|Out-Null}catch{$rejected=$true}
        if(-not $rejected){throw '上传检查未拒绝测试中的私人路径或密钥。'}
    }
    'example.txt'|Set-Content -LiteralPath (Join-Path $sandbox 'upload-files.txt') -Encoding utf8
    [IO.File]::WriteAllBytes($example,[Text.Encoding]::Unicode.GetBytes(('Z'+':\'+'private\sample.docx')))
    $rejected=$false
    try{& (Join-Path $ProjectRoot 'check-upload.ps1') -ProjectDirectory $sandbox|Out-Null}catch{$rejected=$true}
    if(-not $rejected){throw '上传检查未拒绝其他编码的私人路径。'}
}finally{
    # 本轮安全测试沙箱可删除，必须确认目标仍在测试输出目录内。
    $resolved=[IO.Path]::GetFullPath($sandbox)
    if(-not $resolved.StartsWith($TestOutputRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '测试清理越界。'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
& (Join-Path $PSScriptRoot 'test-github-update.ps1')
Write-Output 'PROJECT_STATIC_CHECK_PASSED'
