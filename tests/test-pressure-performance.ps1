#requires -Version 7.0
# 合成压力文档的单轮计时：读取原生阶段消息，计时范围包含监管进程退出，不包含独立审计。
param(
    [Parameter(Mandatory)][string]$RunDirectory,
    [Parameter(Mandatory)][ValidateSet('Baseline','Current')][string]$Version,
    [ValidateRange(1,9)][int]$Round=1,
    [switch]$NoNumbers
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
$project=[IO.Path]::GetFullPath($ProjectRoot)
$sample = Join-Path (New-TestFixtures).Directory 'performance-100-400.docx'
$run=[IO.Path]::GetFullPath($RunDirectory)
$exe=if($Version -eq 'Baseline'){Join-Path $run 'baseline\DocxMathTypeBatch.exe'}else{Join-Path $project 'build\bin\DocxMathTypeBatch.exe'}
$name=$Version.ToLowerInvariant()+$(if($NoNumbers){'-unnumbered'}else{''})+'-'+$Round
$output=Join-Path $run ($name+'.docx');$report=Join-Path $run ($name+'.report.json');$stop=Join-Path $run ($name+'.stop')
if(Test-Path -LiteralPath $output){throw '本轮结果已存在，请使用新的轮次或目录。'}
$originalHash=(Get-FileHash -LiteralPath $sample).Hash
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
public static class PressureTiming {
    public class Mark { public string Name; public double Milliseconds; }
    public class Result { public double TotalMilliseconds; public List<Mark> Marks=new List<Mark>(); }
    // 实时读管道并逐行落盘；程序本身不增加计时线程或等待。
    public static Result Run(string exe,string[] arguments,string log,string stop) {
        var info=new ProcessStartInfo(exe){UseShellExecute=false,CreateNoWindow=true,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true};
        foreach(var value in arguments)info.ArgumentList.Add(value);
        info.Environment.Remove("DOCX_MATH_GUI_COMMIT");
        using(var process=new Process{StartInfo=info})
        using(var writer=new StreamWriter(log){AutoFlush=true}) {
            var result=new Result();var timer=Stopwatch.StartNew();process.Start();process.StandardInput.Close();
            var errors=process.StandardError.ReadToEndAsync();
            while(true) {
                var pending=process.StandardOutput.ReadLineAsync();
                while(!pending.Wait(100))if(timer.Elapsed.TotalSeconds>600){File.WriteAllText(stop,"cancel");process.WaitForExit(15000);throw new Exception("压力计时超时，已向本任务请求停止："+log);}
                var line=pending.Result;if(line==null)break;writer.WriteLine(line);
                string stage=line.StartsWith("UI_PHASE ")?line.Substring(9):line.StartsWith("TIMING_STAGE ")?line.Substring(13):null;
                if(stage!=null){var elapsed=timer.Elapsed.TotalMilliseconds;result.Marks.Add(new Mark{Name=stage,Milliseconds=elapsed});Console.WriteLine("PERF {0:F3}s {1}",elapsed/1000,stage);}
            }
            process.WaitForExit();result.TotalMilliseconds=timer.Elapsed.TotalMilliseconds;File.WriteAllText(log+".errors",errors.Result);
            if(process.ExitCode!=0)throw new Exception("压力转换失败，退出码="+process.ExitCode+"；日志="+log);
            return result;
        }
    }
}
'@
$arguments=@('--worker-convert',$sample,$output,$report,'paragraph',$stop,'--scope=all');if(-not $NoNumbers){$arguments+='--numbers'}
$timing=[PressureTiming]::Run($exe,$arguments,(Join-Path $run ($name+'.log')),$stop)
$result=Get-Content -LiteralPath $report -Raw|ConvertFrom-Json
$audit=(& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/audit_docx.py') $output)|ConvertFrom-Json
if($LASTEXITCODE -ne 0 -or -not $result.WordReopenVerified -or -not $result.OriginalFileBytesUnchanged -or -not $result.ExistingNativeStreamsUnchanged -or -not $result.ExistingMediaBytesPreserved -or $result.FullDevelopmentReadback -or $result.ReopenedEquationActivations -ne 0){throw '压力转换未通过完整性检查。'}
if((Get-FileHash -LiteralPath $sample).Hash -ne $originalHash){throw '压力测试原件发生变化。'}
$signature=[ordered]@{
    Status=$result.Status;Targets=$result.Targets;Succeeded=$result.Succeeded;FailedPreserved=$result.FailedPreserved
    NotAttempted=$result.NotAttempted;RejectedPreserved=$result.RejectedPreserved;NumbersAdded=$result.NumbersAdded
    Items=@($result.Items|Select-Object @{Name='Source';Expression={$_.Source.Trim()}},Status,Reason,Display,NativeFullSize,Row,Column,AddedParagraphs)
    MathType=$audit.MathType;Text=$audit.Text;TableTopology=$audit.TableTopologySHA256;OMML=$audit.OMML
}
$baseline=Join-Path $run ($(if($NoNumbers){'baseline-unnumbered'}else{'baseline'})+'-1.timing.json')
if(Test-Path -LiteralPath $baseline){
    $previous=Get-Content -LiteralPath $baseline -Raw|ConvertFrom-Json -AsHashtable
    if(($signature|ConvertTo-Json -Depth 10 -Compress) -ne ($previous.Signature|ConvertTo-Json -Depth 10 -Compress)){throw '优化前后转换结果、编号、正文或表格不一致。'}
}
$evidence=[ordered]@{
    Status='PASSED';Version=$Version;Round=$Round;Executable=$exe;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash
    Input=$sample;InputSHA256=$originalHash;Settings=('全部公式；跟随原文字号/行距；'+$(if($NoNumbers){'不编号'}else{'添加右编号 (1)'})+'；普通生产路径')
    Timing=$timing;Signature=$signature;Output=$output;Report=$report;IndependentPackageAudit=$true
}
$evidence|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $run ($name+'.timing.json')) -Encoding utf8
Write-Output ('PRESSURE_PERFORMANCE_PASSED: '+$name+' '+[Math]::Round($timing.TotalMilliseconds/1000,3)+' s；'+$result.Succeeded+' 个成功，'+$result.NumbersAdded+' 个右编号')
