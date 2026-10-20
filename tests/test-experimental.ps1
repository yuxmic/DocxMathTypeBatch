# 实验分支使用合成的正文、表格、文本框、页眉页脚和注释；不读取私人样本。
param([switch]$SkipBuild,[ValidateSet('basic','body','header-only','failures','recovery','long','review','tables','textboxes','headers','notes','no-targets','no-changes','all')][string[]]$Case=@('basic','tables','textboxes','headers','notes','no-targets','no-changes','all'))
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
$run=Join-Path $TestOutputRoot ('实验分支-'+[Guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build/bin/DocxMathTypeBatch.exe'
$helper=Join-Path $PSScriptRoot 'helpers/test_experimental.py'
& python -B -X utf8 $helper fixture $run
if($LASTEXITCODE -ne 0){throw '实验样本生成失败'}
$source=Join-Path $run 'input.docx';$originalHash=(Get-FileHash -LiteralPath $source).Hash
$cases=[ordered]@{basic=@();tables=@('tables');textboxes=@('textboxes');headers=@('headers');notes=@('notes');all=@('tables','textboxes','headers','notes')}
$expected=@{basic=3;tables=11;textboxes=5;headers=7;notes=7;all=21}
$evidence=[ordered]@{Status='FAILED';ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@();Directory=$run}
function Invoke-ExtraConversion([string]$Name,[string]$InputFile,[string[]]$Options,[int]$Expected,[string]$Reference='basic',[string]$Style='paragraph',[switch]$Reverse,[switch]$SelfTest){
    $output=Join-Path $run ($Name+'.docx');$report=Join-Path $run ($Name+'.json');$log=Join-Path $run ($Name+'.log');$hash=(Get-FileHash -LiteralPath $InputFile).Hash
    $mode=if($SelfTest){'--self-test'}elseif($Reverse){'--worker-reverse'}else{'--worker-convert'}
    $arguments=@($mode,$InputFile,$output,$report,$Style);if(-not $SelfTest){$arguments+=Join-Path $run 'stop.flag'};$arguments+=$Options
    $program=if($SelfTest){Join-Path $ProjectRoot 'build/bin/docx-math-convert.exe'}else{$exe}
    $process=Start-Process -FilePath $program -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($Name+'.errors')) -PassThru
    try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ($Name+"`n"+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
    $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
    $status=if($Name -in @('nonascii','rollback')){'PARTIAL'}else{'COMPLETE'}
    if($result.Status -ne $status -or $result.Succeeded -ne $Expected -or (Get-FileHash -LiteralPath $InputFile).Hash -ne $hash){throw ('扩展组合核验失败：'+$Name)}
    if($SelfTest -and $result.ReopenedEquationActivations -ne 3){throw '开发回读次数错误计入了未激活的实验对象'}
    if($Name -eq 'noformat-advanced' -and ($result.NumbersAdded -ne 0 -or $result.NumbersUpdated -ne 0)){throw '实验分支错误应用正文右编号'}
    if($Reverse){$referenceSource=if($Name -eq 'complex-reverse'){Join-Path $run 'complex-input.docx'}else{$source};& python -B -X utf8 $helper check-reverse $referenceSource $output}else{& python -B -X utf8 $helper check $InputFile $output $Name (Join-Path $run ($Reference+'.docx'))}
    if($LASTEXITCODE -ne 0){throw ('扩展组合独立审计失败：'+$Name)}
    $evidence.Cases+=@{Name=$Name;Succeeded=$result.Succeeded;OriginalUnchanged=$true;FormattingApplied=$result.Experimental.FormattingApplied}
    Write-Output ('EXPERIMENTAL_CASE_PASSED: '+$Name)
}
function Invoke-PartialConversion([string]$Name,[string]$InputFile,[string[]]$Options,[string]$RegionFault='',[string]$BodyFault='',[switch]$Reverse){
    # 故障只传给本次子进程，原件和所有测试结果均位于本轮目录。
    $previousRegion=$env:DOCX_MATH_REGION_FAULT;$previousBody=$env:DOCX_MATH_BODY_FAULT
    $hash=(Get-FileHash -LiteralPath $InputFile).Hash;$log=Join-Path $run ($Name+'.log');$flag=Join-Path $run ($Name+'.flag');$process=$null
    try{
        $env:DOCX_MATH_REGION_FAULT=$RegionFault;$env:DOCX_MATH_BODY_FAULT=$BodyFault
        $arguments=@($(if($Reverse){'--worker-reverse'}else{'--worker-convert'}),$InputFile,(Join-Path $run ($Name+'.docx')),(Join-Path $run ($Name+'.json')),$(if($Reverse){'bracket'}else{'paragraph'}),$flag)+$Options
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($Name+'.errors')) -PassThru
        if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ($Name+"`n"+[IO.File]::ReadAllText($log))}
        $result=Get-Content -LiteralPath (Join-Path $run ($Name+'.json')) -Raw|ConvertFrom-Json
        $status=if($RegionFault -or $BodyFault){'PARTIAL'}else{'COMPLETE'}
        if($result.Status -ne $status -or (Get-FileHash -LiteralPath $InputFile).Hash -ne $hash -or @(Get-ChildItem -LiteralPath $run -Filter ($Name+'.docx.work-*')).Count){throw ('部分完成状态、原件或临时清理不符：'+$Name)}
        $evidence.Cases+=@{Name=$Name;Status=$result.Status;Succeeded=$result.Succeeded;FailedPreserved=$result.FailedPreserved;NotAttempted=$result.NotAttempted;Issues=$result.Issues.Count;OriginalUnchanged=$true}
        return $result
    }finally{
        $env:DOCX_MATH_REGION_FAULT=$previousRegion;$env:DOCX_MATH_BODY_FAULT=$previousBody
        if($process){if(-not $process.HasExited){[IO.File]::WriteAllText($flag,'stop');$null=$process.WaitForExit(30000)};$process.Dispose()}
    }
}
function Assert-ExperimentalPhases([string]$Log,[switch]$Converted,[switch]$Attempted){
    # 实验阶段只按准备、转换、核验推进；无目标不得进入转换，无改动不得进入保存重开核验。
    $text=[IO.File]::ReadAllText($Log)
    if($Converted){
        $phases=[string[]]@([regex]::Matches($text,'UI_EXPERIMENTAL_PHASE (\w+)')|ForEach-Object{$_.Groups[1].Value})
        if(($phases -join ',') -ne 'PREPARING,CONVERTING,VERIFYING'){throw ('实验阶段顺序错误：'+($phases -join ','))}
        if($text.IndexOf('EXPERIMENTAL_CONVERSION_VERIFIED') -lt $text.LastIndexOf('UI_EXPERIMENTAL_PHASE VERIFYING')){throw '实验完成标记早于核验阶段'}
    }elseif($Attempted){
        if($text.Contains('EXPERIMENTAL_CONVERSION_VERIFIED') -or $text.Contains('UI_EXPERIMENTAL_PHASE VERIFYING')){throw '有候选但无成功替换的实验阶段仍执行了保存重开核验'}
    }else{
        if($text.Contains('EXPERIMENTAL_CONVERSION_VERIFIED') -or $text.Contains('UI_EXPERIMENTAL_PHASE CONVERTING')){throw '短路的实验阶段仍执行了转换或保存重开'}
    }
}
try{
    foreach($name in @('basic')+@($Case|Where-Object {$_ -notin @('basic','body','header-only','failures','recovery','long','review','no-targets','no-changes')})){
        $output=Join-Path $run ($name+'.docx');$report=Join-Path $run ($name+'.json');$log=Join-Path $run ($name+'.log')
        $arguments=@('--worker-convert',$source,$output,$report,'paragraph',(Join-Path $run 'stop.flag'))+@($cases[$name]|ForEach-Object{'--experimental-'+$_})
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
        try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('实验转换失败：'+$name+"`n"+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
        $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
        if($result.Status -ne 'COMPLETE' -or $result.Succeeded -ne $expected[$name] -or $result.FailedPreserved -ne 0){throw ('实验转换数量或状态错误：'+$name+"`n"+($result|ConvertTo-Json -Depth 5))}
        if((Get-FileHash -LiteralPath $source).Hash -ne $originalHash){throw '实验转换修改了输入原件'}
        & python -B -X utf8 $helper check $source $output $name (Join-Path $run 'basic.docx')
        if($LASTEXITCODE -ne 0){throw ('实验输出独立检查失败：'+$name)}
        if($name -eq 'tables'){Assert-ExperimentalPhases $log -Converted}
        $evidence.Cases+=@{Name=$name;Succeeded=$result.Succeeded;OriginalUnchanged=$true;FormattingApplied=$result.Experimental.FormattingApplied}
        Write-Output ('EXPERIMENTAL_CASE_PASSED: '+$name)
    }
    if($Case -contains 'all' -or $Case -contains 'no-targets'){
        # 正文有公式而四个启用区域只有普通文字：实验阶段应在只读预检后直接短路。
        & python -B -X utf8 $helper no-targets-fixture $run
        if($LASTEXITCODE -ne 0){throw '无目标样本创建失败'}
        $inputFile=Join-Path $run 'no-targets-input.docx';$hash=(Get-FileHash -LiteralPath $inputFile).Hash
        $output=Join-Path $run 'no-targets.docx';$report=Join-Path $run 'no-targets.json';$log=Join-Path $run 'no-targets.log'
        $arguments=@('--worker-convert',$inputFile,$output,$report,'paragraph',(Join-Path $run 'stop.flag'))+@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run 'no-targets.errors') -PassThru
        try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('无目标实验转换失败：'+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
        if((Get-FileHash -LiteralPath $inputFile).Hash -ne $hash){throw '无目标实验转换修改了输入原件'}
        $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
        if($result.Status -ne 'COMPLETE' -or $result.Succeeded -ne 3 -or -not $result.Experimental.Enabled -or $result.Experimental.Skipped -ne 'NO_TARGETS' -or $result.Experimental.Succeeded -ne 0 -or $result.Experimental.Items.Count -ne 0 -or $result.Experimental.Issues.Count -ne 0){throw ('无目标实验报告不符合短路契约'+($result|ConvertTo-Json -Depth 6))}
        if(-not ([IO.File]::ReadAllText($log)).Contains('UI_EXPERIMENTAL_SKIPPED NO_TARGETS')){throw '无目标实验缺少短路日志'}
        Assert-ExperimentalPhases $log
        & python -B -X utf8 $helper check-no-targets $run
        if($LASTEXITCODE -ne 0){throw '无目标实验独立审计失败'}
        $evidence.Cases+=@{Name='no-targets';Succeeded=$result.Succeeded;Skipped=$result.Experimental.Skipped;OriginalUnchanged=$true}
        Write-Output 'EXPERIMENTAL_CASE_PASSED: no-targets'
    }
    if($Case -contains 'all' -or $Case -contains 'no-changes'){
        # 额外区域有候选但全部非 ASCII 失败：实验阶段应跳过保存重开，保留原式。
        & python -B -X utf8 $helper no-changes-fixture $run
        if($LASTEXITCODE -ne 0){throw '无改动样本创建失败'}
        $inputFile=Join-Path $run 'no-changes-input.docx';$hash=(Get-FileHash -LiteralPath $inputFile).Hash
        $output=Join-Path $run 'no-changes.docx';$report=Join-Path $run 'no-changes.json';$log=Join-Path $run 'no-changes.log'
        $arguments=@('--worker-convert',$inputFile,$output,$report,'paragraph',(Join-Path $run 'stop.flag'))+@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run 'no-changes.errors') -PassThru
        try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('无改动实验转换失败：'+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
        if((Get-FileHash -LiteralPath $inputFile).Hash -ne $hash){throw '无改动实验转换修改了输入原件'}
        $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
        if($result.Status -ne 'PARTIAL' -or $result.Succeeded -ne 3 -or $result.FailedPreserved -lt 1 -or -not $result.Experimental.Enabled -or $result.Experimental.Skipped -ne 'NO_CHANGES' -or $result.Experimental.Succeeded -ne 0 -or $result.Experimental.Items.Count -ne $result.FailedPreserved){throw ('无改动实验报告不符合短路契约'+($result|ConvertTo-Json -Depth 6))}
        $logText=[IO.File]::ReadAllText($log)
        if(-not $logText.Contains('EXPERIMENTAL_NO_CHANGES') -or -not $logText.Contains('UI_EXPERIMENTAL_SKIPPED NO_CHANGES')){throw '无改动实验缺少短路日志'}
        Assert-ExperimentalPhases $log -Attempted
        & python -B -X utf8 $helper check-no-changes $run
        if($LASTEXITCODE -ne 0){throw '无改动实验独立审计失败'}
        $evidence.Cases+=@{Name='no-changes';Status=$result.Status;Succeeded=$result.Succeeded;FailedPreserved=$result.FailedPreserved;Skipped=$result.Experimental.Skipped;OriginalUnchanged=$true}
        Write-Output 'EXPERIMENTAL_CASE_PASSED: no-changes'
    }
    if($Case -contains 'all' -or $Case -contains 'header-only'){
        & python -B -X utf8 $helper header-fixture $run
        if($LASTEXITCODE -ne 0){throw '单页眉样本创建失败'}
        foreach($name in @('header-only-first','header-only-renumber','header-only-reverse','header-history-first')){
            $inputFile=Join-Path $run $(if($name -eq 'header-only-first'){'header-only-input.docx'}elseif($name -eq 'header-history-first'){'header-history-input.docx'}else{'header-only-first.docx'})
            $hash=(Get-FileHash -LiteralPath $inputFile).Hash
            $arguments=@($(if($name -eq 'header-only-reverse'){'--worker-reverse'}else{'--worker-convert'}),$inputFile,(Join-Path $run ($name+'.docx')),(Join-Path $run ($name+'.json')),$(if($name -eq 'header-only-reverse'){'bracket'}else{'paragraph'}),(Join-Path $run 'stop.flag'),'--experimental-headers')
            if($name -ne 'header-only-reverse'){$arguments+='--numbers'}
            if($name -eq 'header-only-renumber'){$arguments+='--number-start=7'}
            $log=Join-Path $run ($name+'.log')
            $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
            try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ($name+"`n"+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
            if((Get-FileHash -LiteralPath $inputFile).Hash -ne $hash){throw '页眉复测修改了原件'}
            & python -B -X utf8 $helper check-header $run $name
            if($LASTEXITCODE -ne 0){throw ('单页眉审计失败：'+$name)}
            if($name -eq 'header-only-renumber'){
                & (Join-Path $ProjectRoot 'build/bin/native-api-probe.exe') --experimental-native-compare $inputFile (Join-Path $run ($name+'.docx'))
                if($LASTEXITCODE -ne 0){throw '更新编号改变了正文或页眉公式数据'}
                if([IO.File]::ReadAllText($log).Contains('DEFAULT_PREFERENCES_LENGTH')){throw '只更新编号仍初始化了 MathType 转换服务'}
                $headerLog=[IO.File]::ReadAllText($log)
                if($headerLog.Contains('EXPERIMENTAL_CONVERSION_VERIFIED') -or -not $headerLog.Contains('UI_EXPERIMENTAL_SKIPPED NO_TARGETS')){throw '页眉已无候选仍执行了实验保存重开'}
            }
            $evidence.Cases+=@{Name=$name;OriginalUnchanged=$true}
        }
    }
    if($Case -contains 'all' -or $Case -contains 'body'){
        & python -B -X utf8 $helper body-fixture $run
        if($LASTEXITCODE -ne 0){throw '正文混合样本生成失败'}
        # 两种来源按同一正文流程转换；已有公式与新公式统一编号。
        foreach($name in @('body-plain','body-numbered','body-repeat','body-mixed','body-inline','body-display','body-failure','body-inline-failure','body-extra')){
            if($name -eq 'body-mixed'){
                & python -B -X utf8 $helper body-mixed (Join-Path $run 'body-numbered.docx') (Join-Path $run 'body-mixed-input.docx')
                if($LASTEXITCODE -ne 0){throw '新旧混排样本创建失败'}
            }
            $inputName=switch($name){'body-repeat'{'body-numbered.docx'};'body-mixed'{'body-mixed-input.docx'};'body-failure'{'body-failure-input.docx'};'body-inline-failure'{'body-inline-failure-input.docx'};default{'body-input.docx'}}
            $inputFile=Join-Path $run $inputName;$inputHash=(Get-FileHash -LiteralPath $inputFile).Hash
            $selfTest=$name -in @('body-failure','body-inline-failure');$program=if($selfTest){Join-Path $ProjectRoot 'build/bin/docx-math-convert.exe'}else{$exe}
            $style=if($name -in @('body-numbered','body-extra')){'size=18'}else{'paragraph'}
            $arguments=@($(if($selfTest){'--self-test'}else{'--worker-convert'}),$inputFile,(Join-Path $run ($name+'.docx')),(Join-Path $run ($name+'.json')),$style)
            if(-not $selfTest){$arguments+=Join-Path $run 'stop.flag'}
            if($name -notin @('body-plain','body-inline')){$arguments+='--numbers'}
            if($name -in @('body-numbered','body-extra')){$arguments+=@('--inline-spacing=exact:30','--display-spacing=double')}
            if($name -eq 'body-inline'){$arguments+='--scope=inline'}
            if($name -eq 'body-display'){$arguments+='--scope=display'}
            if($name -eq 'body-extra'){$arguments+=@('--experimental-tables','--experimental-headers')}
            $log=Join-Path $run ($name+'.log')
            $process=Start-Process -FilePath $program -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
            try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ($name+"`n"+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
            if((Get-FileHash -LiteralPath $inputFile).Hash -ne $inputHash){throw '正文转换修改了原件'}
            & python -B -X utf8 $helper check-body $run $name
            if($LASTEXITCODE -ne 0){throw ('正文统一处理审计失败：'+$name)}
            $evidence.Cases+=@{Name=$name;OriginalUnchanged=$true}
        }
    }
    if($Case -contains 'all'){
        $all=@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        $locations=@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        Invoke-ExtraConversion 'reverse-all' (Join-Path $run 'all.docx') $locations 21 -Style bracket -Reverse
        $isolated=Join-Path $run 'experimental-only.docx'
        Invoke-ExtraConversion 'noformat-basic' $isolated @() 0 -Reference 'noformat-basic'
        Invoke-ExtraConversion 'noformat-default' $isolated $all 18 -Reference 'noformat-basic'
        Invoke-ExtraConversion 'noformat-advanced' $isolated ($all+@('--inline-spacing=exact:36','--display-spacing=double','--numbers')) 18 -Reference 'noformat-basic' -Style 'size=24'
        & (Join-Path $ProjectRoot 'build/bin/native-api-probe.exe') --experimental-native-compare (Join-Path $run 'noformat-default.docx') (Join-Path $run 'noformat-advanced.docx')
        if($LASTEXITCODE -ne 0){throw '正文格式设置影响了实验公式原生内容'}
        Invoke-ExtraConversion 'nonascii' (Join-Path $run 'nonascii-input.docx') @('--experimental-tables') 10
        Invoke-ExtraConversion 'rollback' (Join-Path $run 'rollback-input.docx') @('--experimental-tables') 10 -SelfTest
        $scopeInput=Join-Path $run 'scope-input.docx'
        Invoke-ExtraConversion 'scope-basic' $scopeInput @() 4 -Reference 'scope-basic'
        Invoke-ExtraConversion 'scope-inline' $scopeInput ($all+@('--scope=inline')) 21 -Reference 'scope-basic'
        Invoke-ExtraConversion 'scope-display' $scopeInput ($all+@('--scope=display')) 3 -Reference 'scope-basic'
        & (Join-Path $ProjectRoot 'build/bin/native-api-probe.exe') --experimental-fixture $source $run
        if($LASTEXITCODE -ne 0){throw '复杂区域样本生成失败'}
        $complexInput=Join-Path $run 'complex-input.docx'
        $complexCount=[int](& python -B -X utf8 $helper targets $complexInput)
        Invoke-ExtraConversion 'complex-basic' $complexInput @() 3 -Reference 'complex-basic'
        Invoke-ExtraConversion 'complex-notes' $complexInput @('--experimental-notes') 8 -Reference 'complex-basic'
        Invoke-ExtraConversion 'complex-all' $complexInput $all $complexCount -Reference 'complex-basic'
        Invoke-ExtraConversion 'complex-reverse' (Join-Path $run 'complex-all.docx') $locations $complexCount -Style bracket -Reverse
        # 在正文 OMML 或实验公式实际替换后取消，当前副本都必须完整撤回。
        foreach($cancel in @(
            @{Name='cancel';Input=$source;Boundary='EXPERIMENTAL_REPLACED';Option='--experimental-tables'},
            @{Name='body-cancel';Input=(Join-Path $run 'body-input.docx');Boundary='BODY_OMML_REPLACED';Option=''},
            @{Name='number-cancel';Input=(Join-Path $run 'body-input.docx');Boundary='NUMBERING_WRITTEN';Option='--numbers'},
            @{Name='rebuild-cancel';Input=(Join-Path $run 'body-input.docx');Boundary='BODY_COPY_REBUILT';Option='';BodyFault='2|RECOVERY'}
        )){
        $cancelName=$cancel.Name;$cancelHash=(Get-FileHash -LiteralPath $cancel.Input).Hash
        $previousBoundary=$env:DOCX_MATH_FAULT_BOUNDARY;$previousGate=$env:DOCX_MATH_FAULT_GATE;$previousBody=$env:DOCX_MATH_BODY_FAULT
        $cancelOutput=Join-Path $run ($cancelName+'.docx');$cancelReport=Join-Path $run ($cancelName+'.json');$cancelFlag=Join-Path $run ($cancelName+'.flag');$cancelLog=Join-Path $run ($cancelName+'.log');$gate=Join-Path $run ($cancelName+'.gate')
        $env:DOCX_MATH_FAULT_BOUNDARY=$cancel.Boundary;$env:DOCX_MATH_FAULT_GATE=$gate;$env:DOCX_MATH_BODY_FAULT=$cancel.BodyFault
        $process=$null
        try{
            $arguments=@('--worker-convert',$cancel.Input,$cancelOutput,$cancelReport,'paragraph',$cancelFlag);if($cancel.Option){$arguments+=$cancel.Option}
            $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $cancelLog -RedirectStandardError (Join-Path $run ($cancelName+'.errors')) -PassThru
            $deadline=[DateTime]::UtcNow.AddSeconds(45)
            do{
                $content=if(Test-Path -LiteralPath $cancelLog){[string](Get-Content -LiteralPath $cancelLog -Raw)}else{''}
                if($process.HasExited -or [DateTime]::UtcNow -gt $deadline){throw '未到达实验替换后的取消边界'}
                Start-Sleep -Milliseconds 50
            }until(([string]$content).Contains('FAULT_BOUNDARY='+$cancel.Boundary))
            [IO.File]::WriteAllText($cancelFlag,'stop')
            if(-not $process.WaitForExit(30000) -or $process.ExitCode -ne 127){throw ('实验取消失败：'+[IO.File]::ReadAllText($cancelLog))}
            if((Test-Path -LiteralPath $cancelOutput) -or (Test-Path -LiteralPath $cancelReport) -or @(Get-ChildItem -LiteralPath $run -Filter ($cancelName+'.docx.work-*')).Count -or (Get-FileHash -LiteralPath $cancel.Input).Hash -ne $cancelHash){throw '实验取消留下输出、工作副本或修改了原件'}
            $evidence.Cases+=@{Name=($cancelName+'-after-replacement');NoOutput=$true;OriginalUnchanged=$true};Write-Output ('EXPERIMENTAL_CANCEL_PASSED: '+$cancelName)
        }finally{
            $env:DOCX_MATH_FAULT_BOUNDARY=$previousBoundary;$env:DOCX_MATH_FAULT_GATE=$previousGate;$env:DOCX_MATH_BODY_FAULT=$previousBody
            if($process){if(-not $process.HasExited){[IO.File]::WriteAllText($cancelFlag,'stop');[IO.File]::WriteAllText($gate,'release');$null=$process.WaitForExit(30000)};$process.Dispose()}
        }
        }
    }
    if($Case -contains 'all' -or $Case -contains 'failures' -or $Case -contains 'recovery'){
        & python -B -X utf8 $helper failure-fixture $run
        if($LASTEXITCODE -ne 0){throw '隔离故障样本创建失败'}
        $locations=@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        $failureInput=Join-Path $run 'failure-input.docx'
        if($Case -notcontains 'recovery'){
        $null=Invoke-PartialConversion 'failure-basic' $failureInput @()
        $null=Invoke-PartialConversion 'failure-all' $failureInput $locations
        foreach($entry in @(@('table','body:table:1'),@('textbox','body:box:1'),@('header','Headers:1:1'),@('footer','Footers:1:1'),@('footnote','Footnotes:1'),@('endnote','Endnotes:1'))){
            foreach($stage in @('SCAN','RECOVERY','VERIFY','OMML')){
                $name='partial-'+$entry[0]+'-'+$stage.ToLowerInvariant()
                $null=Invoke-PartialConversion $name $failureInput $locations ($entry[1]+'|'+$stage)
                & python -B -X utf8 $helper check-failure $run $name $entry[1] $stage
                if($LASTEXITCODE -ne 0){throw ('区域故障独立审计失败：'+$name)}
            }
            $name='partial-'+$entry[0]+'-reverse'
            $null=Invoke-PartialConversion $name (Join-Path $run 'failure-all.docx') $locations ($entry[1]+'|ITEM') -Reverse
            & python -B -X utf8 $helper check-failure $run $name $entry[1] 'ITEM' reverse
            if($LASTEXITCODE -ne 0){throw ('反向保留独立审计失败：'+$name)}
        }
        $null=Invoke-PartialConversion 'failure-repeat' (Join-Path $run 'failure-all.docx') $locations
        }
        $null=Invoke-PartialConversion 'failure-numbered' (Join-Path $run 'header-only-input.docx') @('--numbers','--experimental-headers')
        foreach($entry in @(
            @{Name='partial-prepare';Input='body-input.docx';Fault='3|PREPARE';Options=@('--numbers')},
            @{Name='partial-body-recovery';Input='body-input.docx';Fault='3|RECOVERY';Options=@('--numbers')},
            @{Name='partial-omml-recovery';Input='body-input.docx';Fault='2|RECOVERY';Options=@('--numbers')},
            @{Name='partial-inline-recovery';Input='body-input.docx';Fault='6|RECOVERY';Options=@('--numbers')},
            @{Name='partial-body-reverse';Input='failure-numbered.docx';Fault='1|RECOVERY';Options=@();Reverse=$true},
            @{Name='partial-display-omml';Input='body-input.docx';RegionFault='body:table:1|OMML';Options=@('--numbers','--experimental-tables','--experimental-headers')},
            @{Name='partial-numbering';Input='body-input.docx';Fault='*|NUMBERING';Options=@('--numbers')},
            @{Name='partial-format';Input='failure-numbered.docx';Fault='1|FORMAT';Options=@('--experimental-headers')},
            @{Name='partial-renumber';Input='failure-numbered.docx';Fault='*|NUMBERING';Options=@('--numbers','--number-start=7','--experimental-headers')}
        )){
            $null=Invoke-PartialConversion $entry.Name (Join-Path $run $entry.Input) $entry.Options -BodyFault $entry.Fault -RegionFault $entry.RegionFault -Reverse:($entry.Reverse -eq $true)
            & python -B -X utf8 $helper check-body-failure $run $entry.Name
            if($LASTEXITCODE -ne 0){throw ('正文故障独立审计失败：'+$entry.Name)}
            if($entry.Name -in @('partial-format','partial-renumber')){
                & (Join-Path $ProjectRoot 'build/bin/native-api-probe.exe') --experimental-native-compare (Join-Path $run $entry.Input) (Join-Path $run ($entry.Name+'.docx'))
                if($LASTEXITCODE -ne 0){throw '失败恢复改变了原有公式数据'}
            }
        }
        . (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
        foreach($kind in @('body','region')){
            # 通过真实高级设置开启实验项，验证重做后的计数、文件行状态和详情。
            $folder=Join-Path $run ('partial-ui-'+$kind);New-Item -ItemType Directory -Path $folder|Out-Null
            $inputFile=Join-Path $folder 'input.docx';Copy-Item -LiteralPath (Join-Path $run $(if($kind -eq 'body'){'body-input.docx'}else{'failure-input.docx'})) -Destination $inputFile
            $hash=(Get-FileHash -LiteralPath $inputFile).Hash;$previousRegion=$env:DOCX_MATH_REGION_FAULT;$previousBody=$env:DOCX_MATH_BODY_FAULT;$process=$null
            try{
                $env:DOCX_MATH_REGION_FAULT=if($kind -eq 'region'){'Footnotes:1|VERIFY'}else{''};$env:DOCX_MATH_BODY_FAULT=if($kind -eq 'body'){'3|PREPARE'}else{''}
                $process=Start-Process -FilePath $exe -ArgumentList @('--ui-batch-test',('"'+$folder+'"'),'normal',('"'+$inputFile+'"')) -WindowStyle Hidden -PassThru
                $deadline=[DateTime]::UtcNow.AddSeconds(15)
                while(-not (Test-Path -LiteralPath (Join-Path $folder 'gui-ready.json'))){if($process.HasExited -or [DateTime]::UtcNow -gt $deadline){throw '部分完成界面未就绪'};Start-Sleep -Milliseconds 20}
                $window=[GuiBatchCheck]::Find($process.Id);[GuiBatchCheck]::KillTimer($window,[UIntPtr]2)|Out-Null
                if($kind -eq 'region'){
                    [GuiBatchCheck]::PostMessage([GuiBatchCheck]::GetDlgItem($window,143),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
                    do{$dialog=[GuiBatchCheck]::Dialog($process.Id,'高级设置');if([DateTime]::UtcNow -gt $deadline){throw '部分完成高级设置未打开'};Start-Sleep -Milliseconds 20}while($dialog -eq [IntPtr]::Zero)
                    foreach($id in @(241,242,243,244)){[GuiBatchCheck]::SendMessage([GuiBatchCheck]::GetDlgItem($dialog,$id),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null}
                    [GuiBatchCheck]::PostMessage($dialog,0x111,[IntPtr]1,[IntPtr]::Zero)|Out-Null
                    while([GuiBatchCheck]::Dialog($process.Id,'高级设置') -ne [IntPtr]::Zero){if([DateTime]::UtcNow -gt $deadline){throw '高级设置未提交'};Start-Sleep -Milliseconds 20}
                }
                [GuiBatchCheck]::PostMessage([GuiBatchCheck]::GetDlgItem($window,125),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
                $deadline=[DateTime]::UtcNow.AddSeconds(180);$saved=$false
                while(-not $process.WaitForExit(30)){
                    if([DateTime]::UtcNow -gt $deadline){throw '部分完成界面测试超时'}
                    if(-not $saved -and (Test-Path -LiteralPath (Join-Path $folder 'batch-check.json'))){$saved=Save-GuiWindow -PidValue $process.Id -Path (Join-Path $folder '部分完成.png')}
                }
                if($process.ExitCode -ne 0 -or -not $saved){throw '部分完成界面未正常结束或截图失败'}
                $ui=Get-Content -LiteralPath (Join-Path $folder 'batch-check.json') -Raw|ConvertFrom-Json
                $report=Get-Content -LiteralPath (Join-Path $folder 'report-0.json') -Raw|ConvertFrom-Json
                if($ui.Items.Count -ne 1 -or $ui.Items[0].State -ne 3 -or $ui.Items[0].ExitCode -ne 0 -or -not (Test-Path -LiteralPath $ui.Items[0].Output) -or $report.Status -ne 'PARTIAL' -or (Get-FileHash -LiteralPath $inputFile).Hash -ne $hash){throw '文件状态没有显示部分完成或没有保留可打开结果'}
                if($kind -eq 'region' -and ($report.Succeeded -ne 21 -or -not $ui.Items[0].Details.Contains('UI_EXPERIMENTAL_RESET') -or -not $ui.Items[0].Error.Contains('该区域未转换，原内容保留。'))){throw '实验重做计数或失败详情错误'}
                if($kind -eq 'body' -and ($report.Succeeded -ne 5 -or $report.FailedPreserved -ne 1)){throw '正文部分完成界面计数错误'}
                $evidence.Cases+=@{Name=('partial-ui-'+$kind);Status='PARTIAL';FileState='部分完成';OutputExists=$true;OriginalUnchanged=$true}
                Write-Output ('PARTIAL_UI_PASSED: '+$kind)
            }finally{
                $env:DOCX_MATH_REGION_FAULT=$previousRegion;$env:DOCX_MATH_BODY_FAULT=$previousBody
                if($process){if(-not $process.HasExited){[GuiBatchCheck]::PostMessage([GuiBatchCheck]::Find($process.Id),0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null;$null=$process.WaitForExit(30000)};$process.Dispose()}
            }
        }
    }
    if($Case -contains 'all' -or $Case -contains 'long'){
        & python -B -X utf8 $helper long-fixture $run
        if($LASTEXITCODE -ne 0){throw '长原生公式样本创建失败'}
        $locations=@('tables','textboxes','headers','notes'|ForEach-Object{'--experimental-'+$_})
        $null=Invoke-PartialConversion 'long-all' (Join-Path $run 'long-input.docx') $locations
        & python -B -X utf8 $helper check-long $run
        if($LASTEXITCODE -ne 0){throw '长原生公式诊断截断影响了转换'}
    }
    if($Case -contains 'all' -or $Case -contains 'review'){
        & python -B -X utf8 $helper review-fixture $run
        if($LASTEXITCODE -ne 0){throw '审阅回归样本创建失败'}
        $reviewCases=@(
            @{Name='review-first-header';Count=1},@{Name='review-even-header';Count=1},
            @{Name='review-first-footer';Count=1},@{Name='review-even-footer';Count=1},
            @{Name='review-header-control';Count=2},
            @{Name='review-first-header-reverse';Input='review-first-header.docx';Reverse=$true;Count=1},
            @{Name='review-even-header-reverse';Input='review-even-header.docx';Reverse=$true;Count=1},
            @{Name='review-positions';Positions=$true;Count=10},
            @{Name='review-positions-reverse';Input='review-positions.docx';Positions=$true;Reverse=$true;Count=10},
            @{Name='review-positions-skip';Input='review-positions-input.docx';Positions=$true;Fault='body:table:1|VERIFY';Count=7},
            @{Name='review-positions-reverse-skip';Input='review-positions.docx';Positions=$true;Reverse=$true;Fault='body:table:1|VERIFY';Count=7},
            @{Name='review-positions-restored';Input='review-positions-input.docx';Positions=$true;Fault='body|VERIFY';Count=6},
            @{Name='review-preserved';Positions=$true;Count=10}
        )
        foreach($entry in $reviewCases){
            $name=$entry.Name;$inputFile=Join-Path $run $(if($entry.Input){$entry.Input}else{$name+'-input.docx'})
            $output=Join-Path $run ($name+'.docx');$report=Join-Path $run ($name+'.json');$log=Join-Path $run ($name+'.log')
            $hash=(Get-FileHash -LiteralPath $inputFile).Hash;$previousRegion=$env:DOCX_MATH_REGION_FAULT
            try{
                $env:DOCX_MATH_REGION_FAULT=$entry.Fault
                $arguments=@($(if($entry.Reverse){'--worker-reverse'}else{'--worker-convert'}),$inputFile,$output,$report,$(if($entry.Reverse){'bracket'}else{'paragraph'}),(Join-Path $run ($name+'.flag')))
                if($entry.Positions){$arguments+='--experimental-tables';if(-not $entry.Reverse){$arguments+='--numbers'}}else{$arguments+='--experimental-headers'}
                $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError (Join-Path $run ($name+'.errors')) -PassThru
                try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ($name+"`n"+[IO.File]::ReadAllText($log))}}finally{$process.Dispose()}
            }finally{$env:DOCX_MATH_REGION_FAULT=$previousRegion}
            $result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
            $status=if($entry.Fault -or $name -eq 'review-preserved'){'PARTIAL'}else{'COMPLETE'}
            if($result.Status -ne $status -or $result.Succeeded -ne $entry.Count -or (Get-FileHash -LiteralPath $inputFile).Hash -ne $hash){throw ('审阅回归状态、数量或原件不符：'+$name)}
            if($entry.Positions){
                # 两个方向都用独立反向扫描读取混合的 MathType/LaTeX 实际坐标，不复用报告的偏移算法。
                $arguments=@('--worker-scan',$output,(Join-Path $run ($name+'-scan.json')),'reverse',(Join-Path $run ($name+'-scan.flag')))
                $process=Start-Process -FilePath $exe -ArgumentList @($arguments|ForEach-Object{'"'+$_+'"'}) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $run ($name+'-scan.log')) -RedirectStandardError (Join-Path $run ($name+'-scan.errors')) -PassThru
                try{if(-not $process.WaitForExit(180000) -or $process.ExitCode -ne 0){throw ('最终位置独立重扫失败：'+$name)}}finally{$process.Dispose()}
            }
            & python -B -X utf8 $helper check-review $run $name
            if($LASTEXITCODE -ne 0){throw ('审阅回归独立审计失败：'+$name)}
            $evidence.Cases+=@{Name=$name;Status=$result.Status;Succeeded=$result.Succeeded;OriginalUnchanged=$true;FinalPositionsVerified=[bool]$entry.Positions}
        }
    }
    $evidence.Status='PASSED'
}finally{
    $evidence.SourceSHA256=[ordered]@{}
    foreach($name in @('src/convert_docx.cpp','src/docx_scan.h','src/experimental_conversion.cpp','src/experimental_conversion.h','src/equation_objects.h','src/gui.cpp','tests/helpers/test_experimental.py','tests/test-experimental.ps1')){$evidence.SourceSHA256[$name]=(Get-FileHash -LiteralPath (Join-Path $ProjectRoot $name)).Hash}
    $evidence|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $run 'experimental-results.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'experimental-results.json') -Encoding utf8
}
