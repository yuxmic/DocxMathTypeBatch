#requires -Version 5.1
# 默认检查包内程序；可指定路径作临时诊断，不写入本机绑定。
param([string]$ExePath)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'common.ps1')
try{
    $connection=Get-ConverterConnection $ExePath
    Write-ConverterJson ([ordered]@{SchemaVersion=1;Status='CONNECTED';ExecutablePath=$connection.Path;ApplicationVersion=$connection.Description.Version;CliContractVersion=$connection.Description.CliContractVersion})
    exit 0
}catch{Write-ConverterFailure $_;exit 1}
