#requires -Version 5.1
# 指定路径时先核对身份与接口再保存；不传参数时只检查现有绑定。
param([string]$ExePath)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'common.ps1')
try{
    $connection=Get-ConverterConnection $ExePath
    if($ExePath){Set-RegisteredConverter $connection.Path}
    Write-ConverterJson ([ordered]@{SchemaVersion=1;Status='CONNECTED';ExecutablePath=$connection.Path;ApplicationVersion=$connection.Description.Version;CliContractVersion=$connection.Description.CliContractVersion})
    exit 0
}catch{Write-ConverterFailure $_;exit 1}
