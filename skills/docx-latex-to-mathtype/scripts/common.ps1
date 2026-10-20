# 连接与转换共用的协议检查；只使用 Windows 自带的 PowerShell/.NET。
function Throw-ConverterError([string]$Code,[string]$Message) {
    $failure=[InvalidOperationException]::new($Message)
    $failure.Data['Code']=$Code
    throw $failure
}

function Write-ConverterJson($Value) {
    $Value|ConvertTo-Json -Depth 8 -Compress|Write-Output
}

function Write-ConverterFailure($Record) {
    $code=$Record.Exception.Data['Code']
    if(-not $code){$code='CALL_FAILED'}
    Write-ConverterJson ([ordered]@{SchemaVersion=1;Event='FINISHED';Status='FAILED';ErrorCode=$code;Message=$Record.Exception.Message})
}

function Get-RegisteredConverter {
    $key=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\DocxMathTypeBatch')
    try{if($key){return [string]$key.GetValue('AgentExecutable','')}}finally{if($key){$key.Dispose()}}
    return ''
}

function Set-RegisteredConverter([string]$Path) {
    $key=[Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\DocxMathTypeBatch')
    try{$key.SetValue('AgentExecutable',$Path,[Microsoft.Win32.RegistryValueKind]::String)}finally{$key.Dispose()}
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
    if(-not $ExePath){$ExePath=Get-RegisteredConverter}
    if(-not $ExePath){Throw-ConverterError 'NEEDS_CONNECTION' '尚未连接转换器。请正常打开 EXE 一次，或用 connect.ps1 -ExePath 指定位置。'}
    $path=[IO.Path]::GetFullPath($ExePath)
    if(-not [IO.File]::Exists($path)){Throw-ConverterError 'EXECUTABLE_NOT_FOUND' '已登记的 EXE 不存在，请打开新位置的程序或重新连接。'}
    # 旧版没有产品资源；先读文件身份，避免向旧 GUI 发送不认识的参数而弹窗。
    $identity=[Diagnostics.FileVersionInfo]::GetVersionInfo($path)
    if($identity.ProductName -ne 'DocxMathTypeBatch'){Throw-ConverterError 'UNSUPPORTED_EXECUTABLE' '此文件不是支持连接协议的 DocxMathTypeBatch，请选择新版单文件 EXE。'}
    $start=[Diagnostics.ProcessStartInfo]::new();$start.FileName=$path;$start.Arguments='--describe --json'
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    $start.StandardOutputEncoding=[Text.UTF8Encoding]::new($false);$start.StandardErrorEncoding=[Text.UTF8Encoding]::new($false)
    $process=[Diagnostics.Process]::new();$process.StartInfo=$start
    try{
        [void]$process.Start();$stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
        if(-not $process.WaitForExit(5000)){$process.Kill();$process.WaitForExit();Throw-ConverterError 'QUERY_TIMEOUT' '程序能力查询超时，未开始转换。'}
        $json=$stdout.GetAwaiter().GetResult();$null=$stderr.GetAwaiter().GetResult()
        if($process.ExitCode -ne 0){Throw-ConverterError 'QUERY_FAILED' '程序未能返回调用能力，请更新 EXE 后重新连接。'}
        try{$description=$json|ConvertFrom-Json}catch{Throw-ConverterError 'INVALID_DESCRIPTION' '程序未返回有效的 JSON 能力说明。'}
        if($description.Application -ne 'DocxMathTypeBatch' -or $description.CliContractVersion -ne 1 -or $description.ReportSchemaVersion -ne 1){Throw-ConverterError 'INCOMPATIBLE_CONTRACT' 'EXE 与 Skill 的调用接口不兼容，请安装匹配版本。'}
        return [pscustomobject]@{Path=$path;Description=$description}
    }finally{$process.Dispose()}
}

function Get-ConverterWorkerError([int]$ExitCode,[string]$LogPath) {
    $code=switch($ExitCode){124{'TIMEOUT'}125{'SERVICE_DISCONNECTED'}126{'RECOVERY_REQUIRED'}127{'CANCELLED'}default{'ENGINE_ERROR'}}
    $message='转换未完成，请查看本机任务日志。'
    if([IO.File]::Exists($LogPath)){
        $matches=[regex]::Matches([IO.File]::ReadAllText($LogPath,[Text.Encoding]::UTF8),'(?m)^UI_ERROR ([0-9A-F]+)\s*$')
        if($matches.Count){
            $hex=$matches[$matches.Count-1].Groups[1].Value;$bytes=[byte[]]::new($hex.Length/2)
            for($i=0;$i -lt $bytes.Length;$i++){$bytes[$i]=[Convert]::ToByte($hex.Substring(2*$i,2),16)}
            $message=[Text.Encoding]::UTF8.GetString($bytes)
            if($message -eq 'MATHTYPE_TASK_BUSY'){$code='BUSY';$message='另一个转换正在进行，请等待它结束后再调用。'}
        }
    }
    return [pscustomobject]@{Code=$code;Message=$message}
}
