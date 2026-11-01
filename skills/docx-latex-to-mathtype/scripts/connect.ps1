#requires -Version 5.1
# 只诊断同一完整 Skill 包内的程序，不接受外部路径，也不写本机绑定。
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'common.ps1')
try{
    $connection=Get-ConverterConnection
    Write-ConverterJson ([ordered]@{SchemaVersion=1;Status='CONNECTED';ExecutablePath=$connection.Path;ApplicationVersion=$connection.Description.Version;CliContractVersion=$connection.Description.CliContractVersion})
    exit 0
}catch{Write-ConverterFailure $_;exit 1}
