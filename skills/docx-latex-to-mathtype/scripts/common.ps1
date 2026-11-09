# 连接与转换共用的协议检查；只使用 Windows 自带的 PowerShell/.NET。
function Throw-ConverterError([string]$Code,[string]$Message) {
    $failure=[InvalidOperationException]::new($Message)
    $failure.Data['Code']=$Code
    throw $failure
}

function Write-ConverterJson($Value) {
    $Value['SkillVersion']='0.2.3'
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

function New-ConverterTaskDirectory([string]$TaskDirectory) {
    if(-not $TaskDirectory){$TaskDirectory=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('DocxMathTypeBatch/agent-tasks/'+[Guid]::NewGuid().ToString('N'))}
    $task=[IO.Path]::GetFullPath($TaskDirectory)
    if(Test-Path -LiteralPath $task){Throw-ConverterError 'TASK_DIRECTORY_EXISTS' '每次调用必须使用全新的任务目录，以免混用旧报告或停止标记。'}
    [void][IO.Directory]::CreateDirectory($task)
    return $task
}

function Invoke-ConverterWorker([string]$ExePath,[string[]]$Arguments,[int]$TimeoutSeconds,$Started,[ref]$ExitCode) {
    $start=[Diagnostics.ProcessStartInfo]::new();$start.FileName=$ExePath
    $start.Arguments=(@($Arguments|ForEach-Object{ConvertTo-ConverterArgument $_}) -join ' ')
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    $start.StandardOutputEncoding=[Text.UTF8Encoding]::new($false);$start.StandardErrorEncoding=[Text.UTF8Encoding]::new($false)
    $start.EnvironmentVariables['DOCX_MATH_TIMEOUT_SECONDS']=$TimeoutSeconds.ToString()
    [void]$start.EnvironmentVariables.Remove('DOCX_MATH_GUI_COMMIT')
    $process=[Diagnostics.Process]::new();$process.StartInfo=$start
    $processStarted=$false;$logStream=$null;$errorsStream=$null
    try{
        [void]$process.Start();$processStarted=$true
        # 异步拷贝到本机文件，不把整份逐式日志驻留内存或发送给模型。
        $logStream=[IO.FileStream]::new($Started.LogPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,[IO.FileOptions]::Asynchronous)
        $errorsPath=Join-Path $Started.TaskDirectory 'worker.errors'
        $errorsStream=[IO.FileStream]::new($errorsPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,[IO.FileOptions]::Asynchronous)
        $logCopy=$process.StandardOutput.BaseStream.CopyToAsync($logStream);$errorsCopy=$process.StandardError.BaseStream.CopyToAsync($errorsStream)
        Write-ConverterJson $Started
        while(-not $process.WaitForExit(250)){}
        $process.WaitForExit();[void]$logCopy.GetAwaiter().GetResult();[void]$errorsCopy.GetAwaiter().GetResult()
        $ExitCode.Value=$process.ExitCode
    }finally{
        # 脚本被中断时只请求本任务监管进程清理，不结束其他 Word 或 MathType。
        if($processStarted -and -not $process.HasExited){
            [IO.File]::WriteAllText($Started.StopFile,'stop')
            [void]$process.WaitForExit(30000)
        }
        $process.Dispose()
        if($logStream){$logStream.Dispose()}
        if($errorsStream){$errorsStream.Dispose()}
    }
}

function Read-ConverterScanReport([string]$Path,[string]$ErrorCode) {
    try{$report=[IO.File]::ReadAllText($Path,[Text.Encoding]::UTF8)|ConvertFrom-Json}
    catch{Throw-ConverterError $ErrorCode '无法读取扫描清单，请使用 scan.ps1 重新扫描。'}
    if($report.SchemaVersion -ne 1 -or $report.ReadOnly -ne $true -or $report.ExactTargets -ne $true -or
       $report.Scope -cne 'MAIN_STORY_BODY_ONLY' -or $report.Direction -ceq 'REVERSE' -or
       $report.SourceSHA256 -notmatch '\A[0-9A-Fa-f]{64}\z'){
        Throw-ConverterError $ErrorCode '扫描清单格式不符，请使用 scan.ps1 重新扫描。'
    }
    return $report
}

function Get-ConverterTargetLines($Report,[int[]]$Ids) {
    $requested=[Collections.Generic.HashSet[int]]::new()
    foreach($id in $Ids){
        if($id -lt 1 -or -not $requested.Add($id)){Throw-ConverterError 'TARGET_SELECTION_INVALID' '公式 ID 必须是扫描清单中互不重复的正整数。'}
    }
    $matched=[Collections.Generic.HashSet[int]]::new()
    $lines=[Collections.Generic.List[string]]::new()
    $lines.Add($Report.SourceSHA256.ToUpperInvariant())
    $lastEnd=0
    # 依扫描顺序传递所选范围，并将原式编码成 ASCII 任务文件，避免命令行转义或泄漏公式。
    foreach($item in $Report.Items){
        $id=0
        if(-not [int]::TryParse([string]$item.Id,[ref]$id) -or -not $requested.Contains($id)){continue}
        $start=0;$end=0
        if(-not $matched.Add($id) -or $item.Status -cne 'READY' -or $item.Region -cne 'BODY' -or
           $item.Kind -cnotin @('LATEX','OMML') -or -not $item.Delimiter -or -not $item.Source -or
           -not [int]::TryParse([string]$item.FormulaStart,[ref]$start) -or -not [int]::TryParse([string]$item.FormulaEnd,[ref]$end) -or
           $start -lt $lastEnd -or $end -le $start){
            Throw-ConverterError 'TARGET_SELECTION_INVALID' '所选 ID 无法唯一定位到正文中可转换的公式。'
        }
        $delimiter=[BitConverter]::ToString([Text.Encoding]::UTF8.GetBytes([string]$item.Delimiter)).Replace('-','')
        $source=[BitConverter]::ToString([Text.Encoding]::UTF8.GetBytes([string]$item.Source)).Replace('-','')
        $lines.Add("$start`t$end`t$delimiter`t$source")
        $lastEnd=$end
    }
    if($matched.Count -ne $requested.Count){Throw-ConverterError 'TARGET_SELECTION_INVALID' '扫描清单里找不到所选的公式 ID。'}
    return $lines
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
            elseif($message -eq 'TARGET_SOURCE_CHANGED'){$code=$message;$message='扫描后原件已变化，请重新扫描并选择公式。'}
            elseif($message -eq 'TARGET_SELECTION_MISMATCH'){$code=$message;$message='复扫时未找到所选公式，请重新扫描并选择公式。'}
            elseif($message -eq 'TARGET_SELECTION_INVALID'){$code=$message;$message='公式选择清单无效，请重新扫描并选择公式。'}
            elseif($message -eq 'TARGET_SELECTION_UNSUPPORTED'){$code=$message;$message='逐式选择仅支持标准模式正向转换正文。'}
        }
    }
    return [pscustomobject]@{Code=$code;Reason=$reason;Stage=$stage;Message=$message}
}
