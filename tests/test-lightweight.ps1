#requires -Version 7.0
# 当前轻量化契约回归：正文、模式、字号、多行源码、编号及保存后的独立包审计。
param([switch]$SkipBuild,[switch]$SkipPerformance,[string]$BaselineExecutable,[ValidateRange(1,9)][int]$Repetitions=3)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$project=[IO.Path]::GetFullPath($ProjectRoot)
$run=Join-Path $project ('tests\output\轻量化-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run | Out-Null
$exe=Join-Path $project 'build\bin\DocxMathTypeBatch.exe'
if(-not $SkipBuild){& (Join-Path $project 'build.ps1') -Target All}
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_lightweight_fixtures.py') $run
if($LASTEXITCODE -ne 0){throw '生成样本失败。'}
$evidence=[ordered]@{Status='FAILED';CheckedAt=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss');Directory=$run;Executable=$exe;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@();Performance=@()}

function Invoke-LightweightTask([string]$Name,[string[]]$Arguments,[string]$Program=$exe,[int]$ExpectedExit=0){
    $log=Join-Path $run ($Name+'.log');$errors=Join-Path $run ($Name+'.errors')
    $clock=[Diagnostics.Stopwatch]::StartNew()
    $process=Start-Process -FilePath $Program -ArgumentList @($Arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError $errors -PassThru
    try {
        while(-not $process.WaitForExit(500)){if($clock.Elapsed.TotalSeconds -gt 300){throw ('测试任务超时，请按登记 PID 检查：'+$Name)}}
        if($process.ExitCode -ne $ExpectedExit){throw ('任务退出码错误：'+$Name+'；'+[IO.File]::ReadAllText($log))}
    }finally{$process.Dispose()}
    [pscustomobject]@{Seconds=$clock.Elapsed.TotalSeconds;Log=$log}
}
function Read-LightweightPackage([string]$Path){
    $audit=(& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/audit_docx.py') $Path)|ConvertFrom-Json
    if($LASTEXITCODE -ne 0){throw ('独立包审计失败：'+$Path)}
    $audit
}

$mixed=Join-Path $run 'mixed.docx';$inputHash=(Get-FileHash -LiteralPath $mixed).Hash
$inputPackage=Read-LightweightPackage $mixed
foreach($case in @(
    @{Name='all';Mode='all';Style='paragraph';Count=5;Numbers=0},
    @{Name='inline';Mode='inline';Style='paragraph';Count=1;Numbers=0},
    @{Name='display';Mode='display';Style='paragraph';Count=4;Numbers=0},
    @{Name='fixed';Mode='all';Style='size=15.5';Count=5;Numbers=0},
    @{Name='numbered';Mode='all';Style='paragraph';Count=5;Numbers=4}
)){
    $output=Join-Path $run ($case.Name+'.docx');$report=Join-Path $run ($case.Name+'.json')
    $arguments=@('--worker-convert',$mixed,$output,$report,$case.Style,(Join-Path $run 'stop.flag'),('--scope='+$case.Mode))
    if($case.Numbers){$arguments+='--numbers'}
    $timing=Invoke-LightweightTask $case.Name $arguments
    $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
    if($result.Status -ne 'COMPLETE' -or $result.Targets -ne $case.Count -or $result.Succeeded -ne $case.Count -or $result.NumbersAdded -ne $case.Numbers){throw ('转换数量或范围不符：'+$case.Name)}
    if(-not $result.SavedNativeBatch){throw ('正文资产未使用一次读取核验：'+$case.Name)}
    $package=Read-LightweightPackage $output
    if($package.MathType -ne $case.Count -or $package.TableTopologySHA256 -ne $inputPackage.TableTopologySHA256 -or -not $package.Text.Contains('表格 $t_1$ 与 \[t_2\] 保留。')){throw ('表格或包结构不符：'+$case.Name)}
    if($case.Name -eq 'inline' -and (-not $package.Text.Contains('\[\frac{a+1}{b}\]') -or -not $package.Text.Contains('$$'))){throw '未选中的行外原文发生变化。'}
    if($case.Name -eq 'display' -and -not $package.Text.Contains('$x_1+1$')){throw '未选中的行内原文发生变化。'}
    $sizes=@($result.Items | ForEach-Object{[double]$_.NativeFullSize})
    if($case.Name -eq 'fixed'){if(@($sizes|Where-Object{$_ -ne 15.5}).Count){throw '指定字号未写入 MathType 原生内容。'}}
    elseif($case.Mode -ne 'inline'){if(@($sizes|Where-Object{$_ -eq 16}).Count -ne 1 -or @($sizes|Where-Object{$_ -ne 12 -and $_ -ne 16}).Count){throw '原文字号未正确跟随。'}}
    if((Get-FileHash -LiteralPath $mixed).Hash -ne $inputHash){throw '原件被修改。'}
    $evidence.Cases+=@{Name=$case.Name;Report=$report;Output=$output;Seconds=$timing.Seconds;Succeeded=$result.Succeeded;Numbers=$result.NumbersAdded;NativeFullSizes=$sizes;PackageAudit=$true}
    Write-Output ('LIGHTWEIGHT_CASE_PASSED: '+$case.Name)
}

# 反向同时核对分段函数各分支，防止只生成首行公式却通过数量检查。
$reverse=Join-Path $run 'reverse.docx';$reverseReport=Join-Path $run 'reverse.json'
$null=Invoke-LightweightTask 'reverse' @('--worker-reverse',(Join-Path $run 'all.docx'),$reverse,$reverseReport,'bracket',(Join-Path $run 'stop.flag'))
$reverseResult=Get-Content -LiteralPath $reverseReport -Raw|ConvertFrom-Json
$reversePackage=Read-LightweightPackage $reverse
if($reverseResult.Succeeded -ne 5 -or $reversePackage.MathType -ne 0){throw '基本反向导出失败。'}
$exports=($reverseResult.Items.OutputText -join '') -replace '\s',''
foreach($part in @('x_1','a+1','-x','c','d','\leq')){if(-not $exports.Contains($part)){throw ('多行公式导出丢失分支：'+$part)}}
$evidence.Cases+=@{Name='reverse';Report=$reverseReport;Output=$reverse;AllBranchesPresent=$true}

# 生产反向恢复在已有后方成功结果时复用正文快照，不能只覆盖开发自检分支。
$previousBodyFault=$env:DOCX_MATH_BODY_FAULT
try{
    $env:DOCX_MATH_BODY_FAULT='1|RECOVERY'
    $null=Invoke-LightweightTask 'reverse-snapshot-recovery' @('--worker-reverse',(Join-Path $run 'all.docx'),(Join-Path $run 'reverse-snapshot-recovery.docx'),(Join-Path $run 'reverse-snapshot-recovery.json'),'bracket',(Join-Path $run 'stop.flag'))
}finally{$env:DOCX_MATH_BODY_FAULT=$previousBodyFault}
$recovered=Get-Content -LiteralPath (Join-Path $run 'reverse-snapshot-recovery.json') -Raw|ConvertFrom-Json
if($recovered.Status -ne 'PARTIAL' -or $recovered.Succeeded -ne 4 -or $recovered.NotAttempted -ne 1 -or (Read-LightweightPackage (Join-Path $run 'reverse-snapshot-recovery.docx')).MathType -ne 1){throw '反向正文快照恢复回归失败。'}
$evidence.Cases+=@{Name='reverse-snapshot-recovery';Report=(Join-Path $run 'reverse-snapshot-recovery.json');SourcePreserved=$true}

$againReport=Join-Path $run 'again.json'
$null=Invoke-LightweightTask 'again' @('--worker-convert',(Join-Path $run 'all.docx'),(Join-Path $run 'again.docx'),$againReport,'paragraph',(Join-Path $run 'stop.flag'))
$again=Get-Content -LiteralPath $againReport -Raw|ConvertFrom-Json
if($again.Targets -ne 0 -or $again.ExistingMathType -ne 5 -or -not $again.ExistingNativeStreamsUnchanged){throw '已有公式再次转换时发生改变。'}
$evidence.Cases+=@{Name='again';ExistingMathTypeUnchanged=$true;Report=$againReport}

# 无所选公式的任务不初始化 MathType，范围外公式也不算失败。
$emptyReport=Join-Path $run 'empty.json'
$empty=Invoke-LightweightTask 'empty' @('--worker-convert',(Join-Path $run 'layout.docx'),(Join-Path $run 'empty.docx'),$emptyReport,'paragraph',(Join-Path $run 'stop.flag'),'--scope=inline')
if((Get-Content -LiteralPath $emptyReport -Raw|ConvertFrom-Json).Targets -ne 0 -or [IO.File]::ReadAllText($empty.Log).Contains('DEFAULT_PREFERENCES_LENGTH')){throw '无目标任务仍初始化 MathType。'}
$evidence.Cases+=@{Name='empty';MathTypeInitializationSkipped=$true}
$boundaryReport=Join-Path $run 'boundary.json'
$null=Invoke-LightweightTask 'boundary' @('--worker-convert',(Join-Path $run 'boundaries.docx'),(Join-Path $run 'boundary.docx'),$boundaryReport,'paragraph',(Join-Path $run 'stop.flag'))
$boundary=Get-Content -LiteralPath $boundaryReport -Raw|ConvertFrom-Json
if($boundary.Succeeded -ne 2 -or $boundary.RejectedPreserved -lt 3){throw '破损环境、货币或下一公式边界错误。'}
$evidence.Cases+=@{Name='boundary';Report=$boundaryReport;MalformedSourcesPreserved=$true}

# 恢复注入使用自建正文样本，不再依赖历史表格转换契约。
$recoveryInput=Join-Path $run 'recovery-input.docx'
Copy-Item -LiteralPath (Join-Path $run 'layout.docx') -Destination $recoveryInput
& python -B -X utf8 -c 'import sys;from docx import Document;d=Document(sys.argv[1]);d.add_paragraph("$beta_2$");d.add_paragraph("$split_1+split_2$");d.save(sys.argv[1])' $recoveryInput
$recoveryReport=Join-Path $run 'recovery.json'
$null=Invoke-LightweightTask 'recovery' @('--self-test',$recoveryInput,(Join-Path $run 'recovery.docx'),$recoveryReport,'paragraph') (Join-Path $project 'build\bin\docx-math-convert.exe')
$recovery=Get-Content -LiteralPath $recoveryReport -Raw|ConvertFrom-Json
if($recovery.Succeeded -ne 3 -or $recovery.FailedPreserved -ne 2){throw '准备失败或创建失败的原文恢复未通过。'}
$evidence.Cases+=@{Name='recovery';Report=$recoveryReport;ExactSourceRecovery=$true}

# 用同一原生自检入口导出实际页面，核对原生制表位布局及原段距保留。
$layoutInput=Join-Path $run 'layout.docx';$layoutOutput=Join-Path $run 'layout-output.docx'
$null=Invoke-LightweightTask 'layout' @('--self-test',$layoutInput,$layoutOutput,(Join-Path $run 'layout-output.json'),'paragraph') (Join-Path $project 'build\bin\docx-math-convert.exe')
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_native_display.py') simple-layout $layoutInput $layoutOutput
if($LASTEXITCODE -ne 0){throw '行外居中、首行缩进或段距保持未通过。'}
$numberedPreview=Join-Path $run 'numbered-preview.docx'
$null=Invoke-LightweightTask 'numbered-preview' @('--self-test',$mixed,$numberedPreview,(Join-Path $run 'numbered-preview.json'),'paragraph','--numbers') (Join-Path $project 'build\bin\docx-math-convert.exe')
$alignmentInput=Join-Path $run 'alignment.docx'
foreach($numbered in @($false,$true)){
    $name=if($numbered){'alignment-numbered'}else{'alignment-plain'}
    $aligned=Join-Path $run ($name+'.docx')
    $arguments=@('--self-test',$alignmentInput,$aligned,(Join-Path $run ($name+'.json')),'paragraph')
    if($numbered){$arguments+='--numbers'}
    $null=Invoke-LightweightTask $name $arguments (Join-Path $project 'build\bin\docx-math-convert.exe')
    $checks=@($alignmentInput,$aligned);if($numbered){$checks+='--numbers'}
    & python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_display_layout.py') @checks | Set-Content -LiteralPath (Join-Path $run ($name+'-check.json')) -Encoding utf8
    if($LASTEXITCODE -ne 0){throw ('实际页面居中验证失败：'+$name)}
}
$evidence.Cases+=@{Name='layout';MinimalCenteringVerified=$true;OriginalParagraphSpacingPreserved=$true;PlainPdf=($layoutOutput+'.pdf');NumberedPdf=($numberedPreview+'.pdf');AlignmentPlainPdf=(Join-Path $run 'alignment-plain.docx.pdf');AlignmentNumberedPdf=(Join-Path $run 'alignment-numbered.docx.pdf')}

# 实测 MathType 的 aligned 兼容，不能把原生 API 成功误当作无红字。
$alignedOutput=Join-Path $run 'aligned-output.docx'
$null=Invoke-LightweightTask 'aligned' @('--self-test',(Join-Path $run 'aligned.docx'),$alignedOutput,(Join-Path $run 'aligned-output.json'),'paragraph') (Join-Path $project 'build\bin\docx-math-convert.exe')
$null=Invoke-LightweightTask 'aligned-reverse' @('--worker-reverse',$alignedOutput,(Join-Path $run 'aligned-reverse.docx'),(Join-Path $run 'aligned-reverse.json'),'bracket',(Join-Path $run 'stop.flag'))
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_aligned.py') $alignedOutput | Set-Content -LiteralPath (Join-Path $run 'aligned-check.json') -Encoding utf8
if($LASTEXITCODE -ne 0){throw 'aligned 内容、行列或红色残留检查失败。'}
$evidence.Cases+=@{Name='aligned';AllRowsAndColumnsVerified=$true;NoRedResiduals=$true;Pdf=($alignedOutput+'.pdf');Report=(Join-Path $run 'aligned-output.json')}

# 此处验证真实导入；兼容改写的纯规则由 test.ps1 在 Office 专项前检查。
$compatibilityHelper=Join-Path $PSScriptRoot 'helpers/test_tex_compatibility.py'
& python -B -X utf8 $compatibilityHelper fixture $run
if($LASTEXITCODE -ne 0){throw '兼容改写样本生成失败。'}
$compatibilityInput=Join-Path $run 'tex-compatibility-input.docx';$compatibilityHash=(Get-FileHash -LiteralPath $compatibilityInput).Hash
$null=Invoke-LightweightTask 'tex-compatibility' @('--self-test',$compatibilityInput,(Join-Path $run 'tex-compatibility.docx'),(Join-Path $run 'tex-compatibility.json'),'paragraph','--experimental-tables') (Join-Path $project 'build/bin/docx-math-convert.exe')
$null=Invoke-LightweightTask 'tex-compatibility-reverse' @('--worker-reverse',(Join-Path $run 'tex-compatibility.docx'),(Join-Path $run 'tex-compatibility-reverse.docx'),(Join-Path $run 'tex-compatibility-reverse.json'),'bracket',(Join-Path $run 'stop.flag'),'--experimental-tables')
& python -B -X utf8 $compatibilityHelper check $run
if($LASTEXITCODE -ne 0 -or (Get-FileHash -LiteralPath $compatibilityInput).Hash -ne $compatibilityHash){throw '兼容改写页面、内容或原件检查失败。'}
$evidence.Cases+=@{Name='tex-compatibility';NoRedResiduals=$true;OriginalSyntaxPreserved=$true;NestedAndExperimental=$true}

# 顶饰只增加一组正文/表格样本；开发核对结构，正式转换仍沿用原判定。
& python -B -X utf8 $compatibilityHelper accent-fixture $run
if($LASTEXITCODE -ne 0){throw '顶饰样本生成失败。'}
$null=Invoke-LightweightTask 'accents' @('--self-test',(Join-Path $run 'accents-input.docx'),(Join-Path $run 'accents.docx'),(Join-Path $run 'accents.json'),'paragraph','--experimental-tables') (Join-Path $project 'build/bin/docx-math-convert.exe')
$null=Invoke-LightweightTask 'accents-native' @('--accent-export',(Join-Path $run 'accents.docx'),(Join-Path $run 'accents-native.txt')) (Join-Path $project 'build/bin/native-api-probe.exe')
& python -B -X utf8 $compatibilityHelper accent-check $run
if($LASTEXITCODE -ne 0){throw '顶饰转换回归失败。'}
$evidence.Cases+=@{Name='accents';UnicodeAndOmml=$true;BodyAndTable=$true;UnmappedPreserved=$true}
$null=Invoke-LightweightTask 'accents-reverse' @('--worker-reverse',(Join-Path $run 'accents.docx'),(Join-Path $run 'accents-reverse.docx'),(Join-Path $run 'accents-reverse.json'),'bracket',(Join-Path $run 'stop.flag'),'--experimental-tables')
& python -B -X utf8 $compatibilityHelper accent-reverse-check $run
if($LASTEXITCODE -ne 0){throw '顶饰反向导出未能逐式保留异常对象。'}
$evidence.Cases+=@{Name='accents-reverse';NestedDelimitersPreserved=$true;OtherFormulasExported=$true}

# 合成定界符样本覆盖行内/行外上下文，输入不含论文内容。
$source=Join-Path $run 'delimiters.docx';$output=Join-Path $run 'delimiters-output.docx'
$null=Invoke-LightweightTask 'delimiters' @('--self-test',$source,$output,(Join-Path $run 'delimiters-output.json'),'paragraph') (Join-Path $project 'build\bin\docx-math-convert.exe')
& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/test_delimiters.py') $source $output
if($LASTEXITCODE -ne 0){throw '合成定界符回归失败。'}
$evidence.Cases+=@{Name='delimiters';GeneratedFixture=$true}

# 非对称定界符与嵌入工作簿不能影响正常公式；正文和实验区都覆盖百分号往返。
$inputForms=Join-Path $PSScriptRoot 'helpers/test_input_forms.py'
& python -B -X utf8 $inputForms fixture $run
if($LASTEXITCODE -ne 0){throw '输入形式样本生成失败。'}
foreach($delimiter in @('bracket','dollar')){
    $name='reverse-forms-'+$delimiter
    $null=Invoke-LightweightTask $name @('--worker-reverse',(Join-Path $run 'reverse-input-forms.docx'),(Join-Path $run ($name+'.docx')),(Join-Path $run ($name+'.json')),$delimiter,(Join-Path $run 'stop.flag'))
}
foreach($extra in @($false,$true)){
    $name=if($extra){'embedded-forms-extra'}else{'embedded-forms'}
    $arguments=@('--worker-convert',(Join-Path $run 'embedded-input-forms.docx'),(Join-Path $run ($name+'.docx')),(Join-Path $run ($name+'.json')),'paragraph',(Join-Path $run 'stop.flag'))
    if($extra){$arguments+='--experimental-tables'}
    $null=Invoke-LightweightTask $name $arguments
}
$null=Invoke-LightweightTask 'embedded-forms-reverse' @('--worker-reverse',(Join-Path $run 'embedded-forms-extra.docx'),(Join-Path $run 'embedded-forms-reverse.docx'),(Join-Path $run 'embedded-forms-reverse.json'),'bracket',(Join-Path $run 'stop.flag'),'--experimental-tables')
& python -B -X utf8 $inputForms check $run
if($LASTEXITCODE -ne 0){throw '定界符、百分号或嵌入资产保留回归失败。'}
$evidence.Cases+=@{Name='input-forms';AsymmetricDelimiters=$true;PercentRoundtrip=$true;OtherEmbeddingsUnchanged=$true;PartialOutputPreserved=$true}

# 同一物理资产被多个对象引用时，分别核对全导出和仍有表格引用的结果。
& python -B -X utf8 $inputForms shared-fixture $run
if($LASTEXITCODE -ne 0){throw '共享公式资产样本生成失败。'}
foreach($name in @('shared','shared-table')){
    $null=Invoke-LightweightTask $name @('--worker-reverse',(Join-Path $run ($name+'-input.docx')),(Join-Path $run ($name+'.docx')),(Join-Path $run ($name+'.json')),'bracket',(Join-Path $run 'stop.flag'))
}
& python -B -X utf8 $inputForms shared-check $run
if($LASTEXITCODE -ne 0){throw '共享公式资产反向转换回归失败。'}
$evidence.Cases+=@{Name='shared-reverse';AllReferencesExported=$true;TableReferencePreserved=$true}

# 行外公式旁的空白控件及空结果域属于排除边界，不能因整段空白扩展而被删除。
& python -B -X utf8 $inputForms boundary-fixture $run
if($LASTEXITCODE -ne 0){throw '公式外侧边界样本生成失败。'}
$null=Invoke-LightweightTask 'formula-boundaries' @('--worker-convert',(Join-Path $run 'formula-boundaries-input.docx'),(Join-Path $run 'formula-boundaries.docx'),(Join-Path $run 'formula-boundaries.json'),'paragraph',(Join-Path $run 'stop.flag'),'--experimental-tables')
& python -B -X utf8 $inputForms boundary-check $run
if($LASTEXITCODE -ne 0){throw '公式外侧控件或普通域未完整保留。'}
$evidence.Cases+=@{Name='formula-boundaries';LatexAndOmml=$true;ControlsAndFieldsPreserved=$true}

# 快速模式一次保存，正反向及错误继续输出均明确报告未经核验。
$null=Invoke-LightweightTask 'formula-boundaries-fast' @('--worker-convert',(Join-Path $run 'formula-boundaries-input.docx'),(Join-Path $run 'formula-boundaries-fast.docx'),(Join-Path $run 'formula-boundaries-fast.json'),'paragraph',(Join-Path $run 'stop.flag'),'--experimental-tables','--experimental-fast','--numbers')
$null=Invoke-LightweightTask 'formula-boundaries-fast-reverse' @('--worker-reverse',(Join-Path $run 'formula-boundaries-fast.docx'),(Join-Path $run 'formula-boundaries-fast-reverse.docx'),(Join-Path $run 'formula-boundaries-fast-reverse.json'),'bracket',(Join-Path $run 'stop.flag'),'--experimental-tables','--experimental-fast')
$previousFastFault=$env:DOCX_MATH_BODY_FAULT
try{
    $env:DOCX_MATH_BODY_FAULT='1|RECOVERY'
    $null=Invoke-LightweightTask 'formula-boundaries-fast-failure' @('--worker-convert',(Join-Path $run 'formula-boundaries-input.docx'),(Join-Path $run 'formula-boundaries-fast-failure.docx'),(Join-Path $run 'formula-boundaries-fast-failure.json'),'paragraph',(Join-Path $run 'stop.flag'),'--experimental-tables','--experimental-fast')
}finally{$env:DOCX_MATH_BODY_FAULT=$previousFastFault}
& python -B -X utf8 $inputForms fast-check $run
if($LASTEXITCODE -ne 0){throw '实验快速模式的输出或核验标记不符。'}
$evidence.Cases+=@{Name='experimental-fast';ForwardAndReverse=$true;FailureUnverified=$true;PostChecksSkipped=$true}

# 原生行内公式的独立编号行应清掉空格和手动编号，并按行外布局居中。
& python -B -X utf8 $inputForms manual-omml-fixture $run
if($LASTEXITCODE -ne 0){throw '手动编号原生公式样本生成失败。'}
foreach($manualCase in @(@{Name='manual-omml';Flags=@()},@{Name='manual-omml-fast';Flags=@('--experimental-fast')},@{Name='manual-omml-numbered';Flags=@('--numbers')})){
    $null=Invoke-LightweightTask $manualCase.Name (@('--worker-convert',(Join-Path $run 'manual-omml-input.docx'),(Join-Path $run ($manualCase.Name+'.docx')),(Join-Path $run ($manualCase.Name+'.json')),'paragraph',(Join-Path $run 'stop.flag'))+$manualCase.Flags)
}
& python -B -X utf8 $inputForms manual-omml-check $run
if($LASTEXITCODE -ne 0){throw '原生公式手动编号清理或居中失败。'}
$evidence.Cases+=@{Name='manual-omml';ManualNumbersRemoved=$true;NativeDisplayCentered=$true;ControlsAndUnitsPreserved=$true}
$previousManualFault=$env:DOCX_MATH_BODY_FAULT
try{
    $env:DOCX_MATH_BODY_FAULT='1|RECOVERY'
    $null=Invoke-LightweightTask 'manual-omml-recovery' @('--worker-convert',(Join-Path $run 'manual-omml-input.docx'),(Join-Path $run 'manual-omml-recovery.docx'),(Join-Path $run 'manual-omml-recovery.json'),'paragraph',(Join-Path $run 'stop.flag'))
}finally{$env:DOCX_MATH_BODY_FAULT=$previousManualFault}
& python -B -X utf8 $inputForms manual-omml-recovery-check $run
if($LASTEXITCODE -ne 0){throw '原生公式失败后的手动编号恢复检查失败。'}

# LaTeX 四种定界符复用同一编号清理规则，字段和控件仍作为范围边界。
& python -B -X utf8 $inputForms manual-latex-fixture $run
if($LASTEXITCODE -ne 0){throw 'LaTeX 手动编号样本生成失败。'}
foreach($manualCase in @(@{Name='manual-latex';Flags=@()},@{Name='manual-latex-fast';Flags=@('--experimental-fast')},@{Name='manual-latex-numbered';Flags=@('--numbers')})){
    $null=Invoke-LightweightTask $manualCase.Name (@('--worker-convert',(Join-Path $run 'manual-latex-input.docx'),(Join-Path $run ($manualCase.Name+'.docx')),(Join-Path $run ($manualCase.Name+'.json')),'paragraph',(Join-Path $run 'stop.flag'))+$manualCase.Flags)
}
& python -B -X utf8 $inputForms manual-latex-check $run
if($LASTEXITCODE -ne 0){throw 'LaTeX 手动编号清理或居中失败。'}
$evidence.Cases+=@{Name='manual-latex';FourDelimiters=$true;ManualNumbersRemoved=$true;ControlsAndUnitsPreserved=$true}

# 交换保存包中的两条不同公式，数量和字节多重集均相同，仍必须拒绝错误位置。
$oldBoundary=$env:DOCX_MATH_FAULT_BOUNDARY;$oldGate=$env:DOCX_MATH_FAULT_GATE
$swapProcess=$null
try{
    $env:DOCX_MATH_FAULT_BOUNDARY='SAVED_DOCUMENT';$env:DOCX_MATH_FAULT_GATE=Join-Path $run 'swap-release.flag'
    $swapOutput=Join-Path $run 'swap.docx';$swapReport=Join-Path $run 'swap.json';$swapLog=Join-Path $run 'swap.log'
    $swapArguments=@('--worker-convert',$mixed,$swapOutput,$swapReport,'paragraph',(Join-Path $run 'stop.flag'))
    $swapProcess=Start-Process -FilePath $exe -ArgumentList @($swapArguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $swapLog -RedirectStandardError (Join-Path $run 'swap.errors') -PassThru
    $deadline=[DateTime]::UtcNow.AddSeconds(60)
    while(-not ((Get-Content -LiteralPath $swapLog -Raw -ErrorAction SilentlyContinue) -match 'FAULT_BOUNDARY=SAVED_DOCUMENT')){
        if($swapProcess.HasExited -or [DateTime]::UtcNow -gt $deadline){throw '未到达保存后交换边界。'}
        Start-Sleep -Milliseconds 100
    }
    $candidate=@(Get-ChildItem -LiteralPath $run -Filter 'swap.docx.work-*.docx')
    if($candidate.Count -ne 1){throw '无法唯一确定本轮私有副本。'}
    $zip=[IO.Compression.ZipFile]::Open($candidate[0].FullName,[IO.Compression.ZipArchiveMode]::Update)
    try{
        $entries=@($zip.Entries|Where-Object FullName -like 'word/embeddings/*.bin'|Select-Object -First 2)
        $bytes=@();$names=@()
        foreach($entry in $entries){$memory=[IO.MemoryStream]::new();$stream=$entry.Open();try{$stream.CopyTo($memory);$bytes+=,$memory.ToArray();$names+=$entry.FullName}finally{$stream.Dispose();$memory.Dispose()};$entry.Delete()}
        for($i=0;$i -lt 2;$i++){$stream=$zip.CreateEntry($names[$i]).Open();try{$stream.Write($bytes[1-$i])}finally{$stream.Dispose()}}
    }finally{$zip.Dispose()}
    [IO.File]::WriteAllText($env:DOCX_MATH_FAULT_GATE,'release')
    if(-not $swapProcess.WaitForExit(30000) -or $swapProcess.ExitCode -ne 1 -or (Test-Path -LiteralPath $swapOutput) -or (Test-Path -LiteralPath $swapReport) -or (Get-Content -LiteralPath $swapLog -Raw) -notmatch 'Reopened target native content and style identical at its original position'){throw '错误位置公式未被拒绝。'}
    if(@(Get-ChildItem -LiteralPath $run -Filter 'swap.docx.work-*').Count){throw '错误候选未清理。'}
    $evidence.Cases+=@{Name='saved-swap';WrongPositionsRejected=$true;UnknownResultNotPublished=$true}
}finally{
    if($swapProcess){if(-not $swapProcess.HasExited){[IO.File]::WriteAllText($env:DOCX_MATH_FAULT_GATE,'release');$null=$swapProcess.WaitForExit(30000)};$swapProcess.Dispose()}
    $env:DOCX_MATH_FAULT_BOUNDARY=$oldBoundary;$env:DOCX_MATH_FAULT_GATE=$oldGate
}

if(-not $SkipPerformance){
    if(-not $BaselineExecutable){throw '性能对照需要显式传入基线成品。'}
    $evidence.BaselineExecutable=[IO.Path]::GetFullPath($BaselineExecutable)
    $evidence.BaselineExecutableSHA256=(Get-FileHash -LiteralPath $BaselineExecutable).Hash
    $evidence.PerformanceSizeMode='paragraph'
    foreach($shape in @(@(20,80),@(100,400),@(10,2000))){
        $measurements=@()
        for($repeat=0;$repeat -lt $Repetitions;$repeat++){
            foreach($version in $(if($repeat%2){@('Current','Baseline')}else{@('Baseline','Current')})){
                $name='perf-'+($shape -join '-')+'-'+$repeat+'-'+$version
                $program=if($version -eq 'Baseline'){$BaselineExecutable}else{$exe}
                # 两边都使用默认跟随原文，避免将 EQP 准备路径差异算入本次修改。
                $style='paragraph'
                $measurement=Invoke-LightweightTask $name @('--worker-convert',(Join-Path $run ('performance-'+($shape -join '-')+'.docx')),(Join-Path $run ($name+'.docx')),(Join-Path $run ($name+'.json')),$style,(Join-Path $run 'stop.flag')) $program
                $result=Get-Content -LiteralPath (Join-Path $run ($name+'.json')) -Raw|ConvertFrom-Json
                if($result.Succeeded -ne $shape[0] -or $result.Status -ne 'COMPLETE'){throw '性能样本转换未完成。'}
                $measurements+=@{Version=$version;Seconds=$measurement.Seconds;Log=$measurement.Log}
                Write-Output ('LIGHTWEIGHT_TIMING: '+$name+' '+[Math]::Round($measurement.Seconds,3)+' s')
            }
        }
        $medians=@{}
        foreach($version in @('Baseline','Current')){$values=@($measurements|Where-Object Version -eq $version|ForEach-Object Seconds|Sort-Object);$medians[$version]=$values[[int][Math]::Floor($values.Count/2)]}
        $evidence.Performance+=@{Formulas=$shape[0];Paragraphs=$shape[1];Runs=$measurements;MedianSeconds=$medians;ReductionPercent=[Math]::Round(100*(1-$medians.Current/$medians.Baseline),1)}
    }
}
$evidence.Status='PASSED'
$evidence.SourceSHA256=Get-TestSourceHashes
$evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'lightweight-results.json') -Encoding utf8
Write-Output ('LIGHTWEIGHT_PASSED: '+$run)
