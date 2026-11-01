# 测试只从当前项目获取输入；结果和进程身份均属于本轮运行。
# Python 使用 UTF-8 输出 JSON，统一测试控制台编码，避免中文系统按 GBK 解码原生管道。
[Console]::InputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$OutputEncoding=[Console]::OutputEncoding
$ProjectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$TestOutputRoot = Join-Path $ProjectRoot 'tests/output'
New-Item -ItemType Directory -Path $TestOutputRoot,(Join-Path $TestOutputRoot 'work') -Force | Out-Null

function Get-TestSourceHashes {
    # 公开清单同时覆盖程序、专项脚本和测试辅助文件，模块迁移后无需维护第二份列表。
    $hashes=[ordered]@{}
    foreach($relative in Get-Content -LiteralPath (Join-Path $ProjectRoot 'upload-files.txt') | Where-Object {$_ -and -not $_.StartsWith('#')}) {
        $hashes[$relative]=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot $relative)).Hash
    }
    return $hashes
}

function New-TestFixtures([switch]$Native) {
    $folder = Join-Path $TestOutputRoot ('fixtures-' + [Guid]::NewGuid().ToString('N'))
    & python -B -X utf8 (Join-Path $PSScriptRoot 'test_lightweight_fixtures.py') $folder | Out-Null
    if ($LASTEXITCODE -ne 0) { throw '合成样本生成失败。' }
    $inputFile = Join-Path $folder 'mixed.docx'
    Copy-Item -LiteralPath $inputFile -Destination (Join-Path $folder 'input.docx')
    $outputFile = Join-Path $folder 'forward.docx'
    $ownedMath = @()
    if ($Native) {
        $log = Join-Path $folder 'prepare.log'
        $report = Join-Path $folder 'prepare.json'
        # 清除调用者的故障注入参数，准备过程不可停在正式用例的故障边界。
        $previousBoundary=$env:DOCX_MATH_FAULT_BOUNDARY; $previousGate=$env:DOCX_MATH_FAULT_GATE
        $prepareStarted=[DateTime]::UtcNow.ToFileTimeUtc()
        try {
            $env:DOCX_MATH_FAULT_BOUNDARY=$null; $env:DOCX_MATH_FAULT_GATE=$null
            $arguments=@('--worker-convert',$inputFile,$outputFile,$report,'paragraph',(Join-Path $folder 'stop.flag'))
            $process=Start-Process -FilePath (Join-Path $ProjectRoot 'build/bin/DocxMathTypeBatch.exe') -ArgumentList @($arguments|ForEach-Object {'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $folder 'prepare.errors') -PassThru
            try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw '本轮原生样本准备失败，请查看本地 prepare.log。'}}finally{$process.Dispose()}
        } finally { $env:DOCX_MATH_FAULT_BOUNDARY=$previousBoundary; $env:DOCX_MATH_FAULT_GATE=$previousGate }
        $result = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json
        if ($result.Status -ne 'COMPLETE' -or $result.Succeeded -ne 5) { throw '合成原生样本数量错误。' }
        Copy-Item -LiteralPath $outputFile -Destination (Join-Path $folder 'all.docx')
        foreach ($match in [regex]::Matches([IO.File]::ReadAllText($log),'SHARED_MATHTYPE_PID=(\d+) STARTED=(\d+)')) {
            $started=[long]$match.Groups[2].Value
            if($started -ge $prepareStarted){$ownedMath += @{Id=[int]$match.Groups[1].Value; Started=$started}}
        }
    }
    return [pscustomobject]@{Directory=$folder;SampleDirectory=$folder;Input=$inputFile;Output=$outputFile;OwnedMath=$ownedMath}
}
