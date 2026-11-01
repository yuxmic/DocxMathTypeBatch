# 连接与转换共用的协议检查；只使用 Windows 自带的 PowerShell/.NET。
function Throw-ConverterError([string]$Code,[string]$Message) {
    $failure=[InvalidOperationException]::new($Message)
    $failure.Data['Code']=$Code
    throw $failure
}

function Write-ConverterJson($Value) {
    $Value['SkillVersion']='0.2.1'
    $Value|ConvertTo-Json -Depth 8 -Compress|Write-Output
}

function Write-ConverterFailure($Record) {
    $code=$Record.Exception.Data['Code']
    if(-not $code){$code='CALL_FAILED'}
    Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='FINISHED';Status='FAILED';VerificationMode='NOT_COMPLETED';ErrorCode=$code;Reason=$code;Stage='WRAPPER';Message=$Record.Exception.Message})
}

function Get-ConverterExcerpt([string]$Source) {
    # 摘要只保留一行短片段；截断不拆 UTF-16 代理对，连同省略号最多 160 个字符。
    $text=$Source -replace '[\r\n\v\t]',' '
    if($text.Length -le 160){return $text}
    $length=159
    if([char]::IsHighSurrogate($text[$length-1])){$length--}
    return $text.Substring(0,$length)+'…'
}

function ConvertTo-ConverterProblem($Item,[string]$Status,[string]$Stage) {
    # 页码只采用转换前报告值；缺少阶段或页码时保留 null，不从当前文档推算。
    if($Item.Status){$Status=$Item.Status}
    if($Item.Stage){$Stage=$Item.Stage}
    return [ordered]@{Id=$Item.Id;Region=$Item.Region;Status=$Status;Reason=$Item.Reason;Stage=$(if($Stage){$Stage}else{$null});SourceLocation=$Item.SourceLocation;Page=$Item.Page;SourceExcerpt=(Get-ConverterExcerpt $Item.Source)}
}

function ConvertTo-ConverterArgument([string]$Value) {
    # 按 Windows 命令行规则处理引号和尾反斜杠，参数内容不会被当作 PowerShell 代码。
    $quoted=[Text.StringBuilder]::new();[void]$quoted.Append([char]34);$slashes=0
    foreach($character in $Value.ToCharArray()){
        if($character -eq [char]92){$slashes++;continue}
        $count=if($character -eq [char]34){2*$slashes+1}else{$slashes}
        [void]$quoted.Append([char]92,[int]$count);[void]$quoted.Append($character);$slashes=0
    }
    [void]$quoted.Append([char]92,[int](2*$slashes));[void]$quoted.Append([char]34)
    return $quoted.ToString()
}

function Get-ConverterConnection([string]$ExePath) {
    # 直接使用随 Skill 分发的程序，不读取注册表或搜索外部安装位置。
    if(-not $ExePath){$ExePath=Join-Path $PSScriptRoot '../bin/DocxMathTypeBatch.exe'}
    $path=[IO.Path]::GetFullPath($ExePath)
    if(-not [IO.File]::Exists($path)){Throw-ConverterError 'EXECUTABLE_NOT_FOUND' '程序文件缺失。请重新解压包含 bin/DocxMathTypeBatch.exe 的完整 Skill 包。'}
    # 旧版没有产品资源；先读文件身份，避免向旧 GUI 发送不认识的参数而弹窗。
    $identity=[Diagnostics.FileVersionInfo]::GetVersionInfo($path)
    if($identity.ProductName -ne 'DocxMathTypeBatch'){Throw-ConverterError 'UNSUPPORTED_EXECUTABLE' '包内程序身份不符，请重新解压完整 Skill 包。'}
    $start=[Diagnostics.ProcessStartInfo]::new();$start.FileName=$path;$start.Arguments='--describe --json'
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    $start.StandardOutputEncoding=[Text.UTF8Encoding]::new($false);$start.StandardErrorEncoding=[Text.UTF8Encoding]::new($false)
    $process=[Diagnostics.Process]::new();$process.StartInfo=$start
    try{
        [void]$process.Start();$stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
        if(-not $process.WaitForExit(5000)){$process.Kill();$process.WaitForExit();Throw-ConverterError 'QUERY_TIMEOUT' '程序能力查询超时，未开始转换。'}
        $json=$stdout.GetAwaiter().GetResult();$null=$stderr.GetAwaiter().GetResult()
        if($process.ExitCode -ne 0){Throw-ConverterError 'QUERY_FAILED' '包内程序未能返回调用能力，请更新完整 Skill 包后重试。'}
        try{$description=$json|ConvertFrom-Json}catch{Throw-ConverterError 'INVALID_DESCRIPTION' '程序未返回有效的 JSON 能力说明。'}
        if($description.Application -ne 'DocxMathTypeBatch' -or $description.CliContractVersion -ne 1 -or $description.ReportSchemaVersion -ne 1){Throw-ConverterError 'INCOMPATIBLE_CONTRACT' 'EXE 与 Skill 的调用接口不兼容，请安装匹配版本。'}
        return [pscustomobject]@{Path=$path;Description=$description}
    }finally{$process.Dispose()}
}

function Get-ConverterWorkerError([int]$ExitCode,[string]$LogPath) {
    $code=switch($ExitCode){124{'TIMEOUT'}125{'SERVICE_DISCONNECTED'}126{'RECOVERY_REQUIRED'}127{'CANCELLED'}default{'ENGINE_ERROR'}}
    $message='转换未完成，请查看本机任务日志。';$reason=$code;$stage=$null
    if([IO.File]::Exists($LogPath)){
        $log=[IO.File]::ReadAllText($LogPath,[Text.Encoding]::UTF8)
        $phases=[regex]::Matches($log,'(?m)^UI_(?:EXPERIMENTAL_)?PHASE ([A-Z_]+)\s*$')
        if($phases.Count){$stage=$phases[$phases.Count-1].Groups[1].Value}
        $matches=[regex]::Matches($log,'(?m)^UI_ERROR ([0-9A-F]+)\s*$')
        if($matches.Count){
            $hex=$matches[$matches.Count-1].Groups[1].Value;$bytes=[byte[]]::new($hex.Length/2)
            for($i=0;$i -lt $bytes.Length;$i++){$bytes[$i]=[Convert]::ToByte($hex.Substring(2*$i,2),16)}
            $message=[Text.Encoding]::UTF8.GetString($bytes)
            $reason=$message
            if($message -eq 'MATHTYPE_TASK_BUSY'){$code='BUSY';$message='另一个转换正在进行，请等待它结束后再调用。'}
            elseif($message -eq 'Disable revision tracking in source before batch conversion'){$code='TRACK_CHANGES_ENABLED';$message='文件开启修订，请先关闭修订并保存后重试。'}
            elseif($message -eq 'Protected document cannot be converted'){$code='DOCUMENT_PROTECTED';$message='文件已限制编辑，请解除编辑限制后重试。'}
        }
    }
    return [pscustomobject]@{Code=$code;Reason=$reason;Stage=$stage;Message=$message}
}
