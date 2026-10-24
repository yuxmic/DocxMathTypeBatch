#requires -Version 7.0
# 测试入口：每个专项使用独立 PowerShell 进程，避免 COM 和截图类型相互影响。
param([ValidateSet('Static','Core','Safety','UI','All')][string]$Suite='Core',[switch]$SkipBuild)
$ErrorActionPreference='Stop'
& (Join-Path $PSScriptRoot 'check-upload.ps1')
& pwsh -NoProfile -File (Join-Path $PSScriptRoot 'tests/test-project.ps1')
if($LASTEXITCODE -ne 0){throw '项目静态检查失败。'}
if($Suite -eq 'Static'){return}
if(-not $SkipBuild){& (Join-Path $PSScriptRoot 'build.ps1') -Target All}
$groups=[ordered]@{
    Core=@('lightweight','advanced','native-display','native-numbering','experimental','agent')
    Safety=@('cancel','faults','timeouts','installation','link-updates')
    UI=@('gui','diagnostics','batch','ui-style','ui-experience','foreground')
}
$selected=if($Suite -eq 'All'){@('Core','Safety','UI')}else{@($Suite)}
foreach($group in $selected){
    foreach($name in $groups[$group]){
        $arguments=@('-NoProfile','-File',(Join-Path $PSScriptRoot ('tests/test-'+$name+'.ps1')),'-SkipBuild')
        if($name -eq 'lightweight'){$arguments+='-SkipPerformance'}
        & pwsh @arguments
        if($LASTEXITCODE -ne 0){throw ('测试失败：'+$name)}
    }
}
Write-Output ('TEST_SUITE_PASSED: '+$Suite)
