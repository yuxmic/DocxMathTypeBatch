#requires -Version 7.0
# 真实多文件拖入、顺序转换、失败清理后继续、取消整批及高级设置生效回归。
param([switch]$SkipBuild,[switch]$PreviewOnly)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
$run=Join-Path $ProjectRoot ('tests\output\批量转换-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run|Out-Null
$exe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
$baseline = New-TestFixtures -Native:(!$PreviewOnly)
$evidence=[ordered]@{Status='FAILED';Scope=$(if($PreviewOnly){'AdvancedSettingsOnly'}else{'FullBatch'});PreviewOnly=[bool]$PreviewOnly;Directory=$run;ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;Cases=@();AdvancedSettingsApplied=$false;OmmlToggleVerified=$false;QueueEditingVerified=$false;Screenshots=@()}
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class BatchUi {
    public delegate bool EnumProc(IntPtr window,IntPtr data);
    [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr window,int id);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr window,uint message,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window,uint message,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll",EntryPoint="SendMessageW",CharSet=CharSet.Unicode)] static extern IntPtr SendText(IntPtr window,uint message,IntPtr wp,string text);
    [DllImport("user32.dll",EntryPoint="SendMessageW",CharSet=CharSet.Unicode)] static extern IntPtr ReadText(IntPtr window,uint message,IntPtr wp,StringBuilder text);
    public static bool SetWindowText(IntPtr window,string text){return SendText(window,0xC,IntPtr.Zero,text)!=IntPtr.Zero;}
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window,StringBuilder text,int max);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window,StringBuilder text,int max);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr data);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalAlloc(uint flags,UIntPtr bytes);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr handle);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr handle);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalFree(IntPtr handle);
    [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr OpenProcess(uint access,bool inherit,uint pid);
    [DllImport("kernel32.dll")] static extern bool GetProcessTimes(IntPtr handle,out long created,out long exited,out long kernel,out long user);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode)] static extern bool QueryFullProcessImageName(IntPtr handle,uint flags,StringBuilder image,ref uint count);
    [DllImport("kernel32.dll")] static extern bool TerminateProcess(IntPtr handle,uint code);
    [DllImport("kernel32.dll")] static extern uint WaitForSingleObject(IntPtr handle,uint milliseconds);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    public static string Text(IntPtr window){var text=new StringBuilder(4096);ReadText(window,0xD,(IntPtr)text.Capacity,text);return text.ToString();}
    public static IntPtr Dialog(uint pid){IntPtr found=IntPtr.Zero;EnumWindows((w,p)=>{uint owner;GetWindowThreadProcessId(w,out owner);var cls=new StringBuilder(64);GetClassName(w,cls,64);if(owner==pid&&cls.ToString()=="#32770"&&Text(w)=="高级设置"){found=w;return false;}return true;},IntPtr.Zero);return found;}
    public static void Drop(IntPtr window,string[] paths){var text=Encoding.Unicode.GetBytes(string.Join("\0",paths)+"\0\0");var memory=GlobalAlloc(0x42,(UIntPtr)(20+text.Length));if(memory==IntPtr.Zero)throw new Exception("拖入内存分配失败");var pointer=GlobalLock(memory);Marshal.Copy(new int[]{20,0,0,0,1},0,pointer,5);Marshal.Copy(text,0,IntPtr.Add(pointer,20),text.Length);GlobalUnlock(memory);if(!PostMessage(window,0x233,memory,IntPtr.Zero)){GlobalFree(memory);throw new Exception("多文件拖入失败");}}
    public static void KillOwnedWord(uint pid,long expected,long earliest){var handle=OpenProcess(0x101001,false,pid);if(handle==IntPtr.Zero)throw new Exception("打开任务 Word 失败");try{long created,exited,kernel,user;uint count=32768;var image=new StringBuilder((int)count);if(!GetProcessTimes(handle,out created,out exited,out kernel,out user)||created!=expected||created<earliest||!QueryFullProcessImageName(handle,0,image,ref count)||!image.ToString().EndsWith("\\WINWORD.EXE",StringComparison.OrdinalIgnoreCase))throw new Exception("拒绝操作身份不符的 Word");if(!TerminateProcess(handle,1)||WaitForSingleObject(handle,5000)!=0)throw new Exception("注入任务 Word 故障失败");}finally{CloseHandle(handle);}}
}
'@

function Wait-BatchWindow($Process){
    $deadline=[DateTime]::UtcNow.AddSeconds(15)
    do{$window=[GuiBatchCheck]::Find($Process.Id);if($Process.HasExited -or [DateTime]::UtcNow -gt $deadline){throw '批量界面未启动。'};if($window -eq [IntPtr]::Zero){Start-Sleep -Milliseconds 25}}while($window -eq [IntPtr]::Zero)
    $window
}
function Stop-BatchTestWindow($Process){
    if(-not $Process.HasExited){[BatchUi]::PostMessage([GuiBatchCheck]::Find($Process.Id),0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null;if(-not $Process.WaitForExit(15000)){throw '测试界面未结束。'}}
    $Process.Dispose()
}

try{
    # 高级设置打开和保存不启动 Office，只有点击转换后才应用。
    $preview=Join-Path $run 'preview';New-Item -ItemType Directory -Path $preview|Out-Null
    $previewInput=Join-Path $preview 'preview.docx';Copy-Item -LiteralPath (Join-Path $baseline.Directory 'mixed.docx') -Destination $previewInput
    $process=Start-Process -FilePath $exe -ArgumentList @('--ui-test',('"'+$previewInput+'"'),('"'+$preview+'"'),'preview') -WindowStyle Hidden -PassThru
    try{
        $window=Wait-BatchWindow $process
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        while([BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,138),0x1004,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){if([DateTime]::UtcNow -gt $deadline){throw '预览队列未就绪。'};Start-Sleep -Milliseconds 25}
        [BatchUi]::PostMessage([BatchUi]::GetDlgItem($window,143),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        do{$dialog=[BatchUi]::Dialog($process.Id);if([DateTime]::UtcNow -gt $deadline){throw '高级设置未打开。'};Start-Sleep -Milliseconds 25}while($dialog -eq [IntPtr]::Zero)
        $ommlControl=[BatchUi]::GetDlgItem($dialog,246)
        if($ommlControl -eq [IntPtr]::Zero -or -not [BatchUi]::IsWindowEnabled($ommlControl) -or [BatchUi]::SendMessage($ommlControl,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '原生公式转换开关默认未开启。'}
        [BatchUi]::SendMessage($ommlControl,0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        foreach($id in @(201,202)){
            [BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,$id),0x14E,[IntPtr]4,[IntPtr]::Zero)|Out-Null
            [BatchUi]::SendMessage($dialog,0x111,[IntPtr]($id -bor (1 -shl 16)),[BatchUi]::GetDlgItem($dialog,$id))|Out-Null
        }
        [BatchUi]::SetWindowText([BatchUi]::GetDlgItem($dialog,203),'18')|Out-Null
        [BatchUi]::SetWindowText([BatchUi]::GetDlgItem($dialog,204),'24')|Out-Null
        foreach($id in @(222,223)){[BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,$id),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null}
        foreach($entry in @(@(224,4),@(225,1),@(226,2),@(227,1))){[BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,$entry[0]),0x14E,[IntPtr]$entry[1],[IntPtr]::Zero)|Out-Null}
        foreach($entry in @(@(230,'4'),@(231,'2'),@(232,'27'),@(228,'-'))){[BatchUi]::SetWindowText([BatchUi]::GetDlgItem($dialog,$entry[0]),$entry[1])|Out-Null}
        # 截取本任务自己的原生对话框，截图不依赖 Word 或第三方运行库。
        $dialogDpi=[GuiBatchCheck]::SetThreadDpiAwarenessContext([IntPtr](-4))
        $rect=[GuiBatchCheck+Rect]::new();[GuiBatchCheck]::GetWindowRect($dialog,[ref]$rect)|Out-Null
        $bitmap=[Drawing.Bitmap]::new($rect.Right-$rect.Left,$rect.Bottom-$rect.Top);$graphics=[Drawing.Graphics]::FromImage($bitmap);$dc=$graphics.GetHdc()
        try{[GuiBatchCheck]::PrintWindow($dialog,$dc,2)|Out-Null}finally{$graphics.ReleaseHdc($dc)}
        $advancedShot=Join-Path $preview '高级设置.png';$bitmap.Save($advancedShot);$graphics.Dispose();$bitmap.Dispose();$evidence.Screenshots+=$advancedShot
        [GuiBatchCheck]::SetThreadDpiAwarenessContext($dialogDpi)|Out-Null
        [BatchUi]::PostMessage($dialog,0x111,[IntPtr]1,[IntPtr]::Zero)|Out-Null
        Start-Sleep -Milliseconds 100
        if(Test-Path -LiteralPath (Join-Path $preview 'gui-worker.log')){throw '高级设置意外启动转换。'}
        # 重新打开核对提交值；取消的更改不得覆盖之前提交值。
        [BatchUi]::PostMessage([BatchUi]::GetDlgItem($window,143),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        do{$dialog=[BatchUi]::Dialog($process.Id);if([DateTime]::UtcNow -gt $deadline){throw '高级设置未重新打开。'};Start-Sleep -Milliseconds 25}while($dialog -eq [IntPtr]::Zero)
        if([BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,246),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '关闭原生公式转换后没有保存。'}
        if([BatchUi]::Text([BatchUi]::GetDlgItem($dialog,203)) -notmatch '^18' -or [BatchUi]::Text([BatchUi]::GetDlgItem($dialog,204)) -notmatch '^24' -or [BatchUi]::Text([BatchUi]::GetDlgItem($dialog,233)) -ne '预览：[IV-b-AA]'){throw '高级设置或编号预览没有保存。'}
        [BatchUi]::SetWindowText([BatchUi]::GetDlgItem($dialog,203),'99')|Out-Null
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,221),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        [BatchUi]::SetWindowText([BatchUi]::GetDlgItem($dialog,229),'(#C1.#E1)')|Out-Null
        if([BatchUi]::Text([BatchUi]::GetDlgItem($dialog,233)) -ne '预览：(4.27)'){throw '自定义编号预览未生效。'}
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,235),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        if([BatchUi]::Text([BatchUi]::GetDlgItem($dialog,233)) -ne '预览：(1)' -or [BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,201),0x147,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0 -or [BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,246),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '恢复默认未生效。'}
        [BatchUi]::PostMessage($dialog,0x111,[IntPtr]2,[IntPtr]::Zero)|Out-Null
        Start-Sleep -Milliseconds 100
        # 取消恢复默认不提交 OMML 开关；反向禁用该控件且不能改写正向选择。
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,102),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        [BatchUi]::PostMessage([BatchUi]::GetDlgItem($window,143),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        do{$dialog=[BatchUi]::Dialog($process.Id);if([DateTime]::UtcNow -gt $deadline){throw '反向高级设置未打开。'};Start-Sleep -Milliseconds 25}while($dialog -eq [IntPtr]::Zero)
        $ommlControl=[BatchUi]::GetDlgItem($dialog,246)
        if([BatchUi]::IsWindowEnabled($ommlControl) -or [BatchUi]::SendMessage($ommlControl,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '反向原生公式控件未禁用，或取消修改被意外提交。'}
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($dialog,235),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        [BatchUi]::PostMessage($dialog,0x111,[IntPtr]1,[IntPtr]::Zero)|Out-Null
        Start-Sleep -Milliseconds 100
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,101),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        [BatchUi]::PostMessage([BatchUi]::GetDlgItem($window,143),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        do{$dialog=[BatchUi]::Dialog($process.Id);if([DateTime]::UtcNow -gt $deadline){throw '正向高级设置未恢复。'};Start-Sleep -Milliseconds 25}while($dialog -eq [IntPtr]::Zero)
        $ommlControl=[BatchUi]::GetDlgItem($dialog,246)
        if(-not [BatchUi]::IsWindowEnabled($ommlControl) -or [BatchUi]::SendMessage($ommlControl,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '切回正向后原生公式选择未保留。'}
        [BatchUi]::PostMessage($dialog,0x111,[IntPtr]2,[IntPtr]::Zero)|Out-Null
        Start-Sleep -Milliseconds 100
        $evidence.OmmlToggleVerified=$true
        # 实际列表消息覆盖去重、Delete 和清空；这些操作只能改队列。
        $second=Join-Path $preview 'second.docx';$third=Join-Path $preview 'third.docx'
        Copy-Item -LiteralPath $previewInput -Destination $second;Copy-Item -LiteralPath $previewInput -Destination $third
        $sourceHash=(Get-FileHash -LiteralPath $previewInput).Hash
        [BatchUi]::Drop([BatchUi]::GetDlgItem($window,138),@($second,$third,$previewInput));Start-Sleep -Milliseconds 100
        if([BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,138),0x1004,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 3){throw '预览多文档去重失败。'}
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,138),0x100,[IntPtr]0x24,[IntPtr]::Zero)|Out-Null
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,138),0x100,[IntPtr]0x2E,[IntPtr]::Zero)|Out-Null
        if([BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,138),0x1004,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 2){throw 'Delete 未从队列移除文档。'}
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,139),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        if([BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,138),0x1004,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '队列未清空。'}
        foreach($file in @($previewInput,$second,$third)){if((Get-FileHash -LiteralPath $file).Hash -ne $sourceHash){throw '清空队列修改了文档。'}}
        $evidence.QueueEditingVerified=$true
        # 从实际主界面开始转换，独立核对传参及 Word 中真正写入的段落行距和编号。
        [BatchUi]::Drop([BatchUi]::GetDlgItem($window,138),@($previewInput));Start-Sleep -Milliseconds 100
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,137),0x14E,[IntPtr]1,[IntPtr]::Zero)|Out-Null
        [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,125),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(120)
        while([BatchUi]::Text([BatchUi]::GetDlgItem($window,123)) -ne '批次完成'){if($process.HasExited -or [DateTime]::UtcNow -gt $deadline){throw '高级设置界面转换失败。'};Start-Sleep -Milliseconds 100}
        $report=Get-Content -LiteralPath (Join-Path $preview 'report.json') -Raw|ConvertFrom-Json
        if($report.Status -ne 'COMPLETE' -or $report.InlineSpacing -ne 'exact:18.000000' -or $report.DisplaySpacing -ne 'exact:24.000000' -or $report.NumberFormat -ne '[#CI-#Sa-#EA]' -or $report.NumbersAdded -ne 4 -or $report.ConvertOmml -ne $false){throw '界面设置未传给转换进程。'}
        & python -B -X utf8 -c 'import sys;sys.path.insert(0,sys.argv[2]);from docx import Document;from native_number_text import native_visible_numbers;d=Document(sys.argv[1]);p=[p for p in d.paragraphs if p._p.xpath(".//w:object")];assert len(p)==5;assert [p.paragraph_format.line_spacing.pt for p in p]==[18,24,24,24,24];assert native_visible_numbers(d.element.body)==["[IV-b-AA]","[IV-b-BB]","[IV-b-CC]","[IV-b-DD]"]' $report.Output (Join-Path $PSScriptRoot 'helpers')
        if($LASTEXITCODE -ne 0){throw '高级行距或编号未写入 Word。'}
        $evidence.AdvancedSettingsApplied=$true
    }finally{Stop-BatchTestWindow $process}

    if($PreviewOnly){$evidence.Status='PASSED';return}

    foreach($kind in @('failed-middle','native-fault','cancel','reverse')){
        $folder=Join-Path $run $kind;New-Item -ItemType Directory -Path $folder|Out-Null
        $sources=Join-Path $folder 'sources';$other=Join-Path $folder 'more';$out=Join-Path $folder 'output';New-Item -ItemType Directory -Path $sources,$other,$out|Out-Null
        $paths=@((Join-Path $sources '文档.docx'),(Join-Path $sources '中间文档.docx'),(Join-Path $other '文档.docx'))
        $source=Join-Path $baseline.Directory $(if($kind -eq 'reverse'){'all.docx'}else{'mixed.docx'})
        foreach($path in $paths){Copy-Item -LiteralPath $source -Destination $path}
        if($kind -in @('failed-middle','reverse')){[IO.File]::WriteAllText($paths[1],'invalid DOCX fixture')}
        $hashes=@($paths|ForEach-Object{(Get-FileHash -LiteralPath $_).Hash})
        $suffix=if($kind -eq 'reverse'){'_LaTeX'}else{'_MathType'}
        $occupied=Join-Path $out ('文档'+$suffix+'.docx');Copy-Item -LiteralPath $source -Destination $occupied;$occupiedHash=(Get-FileHash -LiteralPath $occupied).Hash
        $mode=if($kind -eq 'cancel'){'stop'}elseif($kind -eq 'reverse'){'reverse'}else{'normal'}
        $clock=[Diagnostics.Stopwatch]::StartNew();$earliest=[DateTime]::UtcNow.ToFileTimeUtc()
        # 首份由参数加入，后两份通过实际 HDROP 消息加入列表，并验证去重。
        $process=Start-Process -FilePath $exe -ArgumentList @('--ui-batch-test',('"'+$folder+'"'),$mode,('"'+$paths[0]+'"')) -WindowStyle Hidden -PassThru
        $injected=$false;$saved=$false;$elapsedBefore='';$liveElapsedUpdates=0
        try{
            $window=Wait-BatchWindow $process;$ready=Join-Path $folder 'gui-ready.json';$deadline=[DateTime]::UtcNow.AddSeconds(15)
            while(-not (Test-Path -LiteralPath $ready)){if([DateTime]::UtcNow -gt $deadline){throw '批量准备超时。'};Start-Sleep -Milliseconds 20}
            [BatchUi]::Drop([BatchUi]::GetDlgItem($window,138),@($paths[1],$paths[2],$paths[0]))
            [BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,142),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
            [BatchUi]::SetWindowText([BatchUi]::GetDlgItem($window,118),$out)|Out-Null
            $deadline=[DateTime]::UtcNow.AddSeconds(10)
            while([BatchUi]::SendMessage([BatchUi]::GetDlgItem($window,138),0x1004,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 3){if([DateTime]::UtcNow -gt $deadline){throw '多文件拖入或去重失败。'};Start-Sleep -Milliseconds 20}
            $deadline=[DateTime]::UtcNow.AddSeconds(180)
            while(-not $process.WaitForExit(25)){
                if([DateTime]::UtcNow -gt $deadline){throw ('批量转换超时：'+$kind)}
                $elapsed=[BatchUi]::Text([BatchUi]::GetDlgItem($window,144));if($elapsedBefore -and $elapsed -ne $elapsedBefore -and -not (Test-Path -LiteralPath (Join-Path $folder 'batch-check.json'))){++$liveElapsedUpdates};$elapsedBefore=$elapsed
                if($kind -eq 'native-fault' -and -not $injected){
                    $logPath=Join-Path $folder 'gui-worker.log';$log=''
                    if(Test-Path -LiteralPath $logPath){$stream=[IO.File]::Open($logPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite);$reader=[IO.StreamReader]::new($stream);try{$log=$reader.ReadToEnd()}finally{$reader.Dispose()}}
                    $segment=($log -split 'UI_BATCH_BEGIN 1\r?\n')
                    if($segment.Count -eq 2 -and $segment[1] -match 'CONVERT_PROGRESS 1/'){
                        $registered=[regex]::Match($segment[1],'OWNED_WORD_PID=(\d+)\r?\nOWNED_WORD_STARTED=(\d+)')
                        if(-not $registered.Success){throw '未登记第二份文档所属 Word。'}
                        [BatchUi]::KillOwnedWord([uint32]$registered.Groups[1].Value,[long]$registered.Groups[2].Value,$earliest);$injected=$true
                    }
                }
                $check=Join-Path $folder 'batch-check.json'
                if(-not $saved -and (Test-Path -LiteralPath $check)){
                    $frozen=[BatchUi]::Text([BatchUi]::GetDlgItem($window,144));Start-Sleep -Milliseconds 200
                    if([BatchUi]::Text([BatchUi]::GetDlgItem($window,144)) -ne $frozen){throw '完成后总耗时仍在变化。'}
                    $shot=Join-Path $folder '批次结果.png';$saved=Save-GuiWindow -PidValue $process.Id -Path $shot;if($saved){$evidence.Screenshots+=$shot}
                }
            }
            if($process.ExitCode -ne 0 -or -not $saved){throw ('批量界面未正常结束：'+$kind)}
            $result=Get-Content -LiteralPath (Join-Path $folder 'batch-check.json') -Raw|ConvertFrom-Json
            foreach($item in $result.Items){
                if($item.State -ne 0){
                    if(-not $item.Details.Contains($item.Input) -or -not $item.Details.Contains('TASK_SUPERVISOR_PID=')){throw '文档结束后没有保留对应日志。'}
                    foreach($other in $result.Items){if($other.Input -ne $item.Input -and $item.Details.Contains($other.Input)){throw '不同文档的转换日志发生串用。'}}
                }
            }
            $states=@($result.Items.State);$expected=if($kind -eq 'cancel'){@(2,5,0)}else{@(2,4,2)}
            if(($states -join ',') -ne ($expected -join ',') -or $liveElapsedUpdates -lt 2 -or $result.ElapsedMilliseconds -le 0 -or $result.ElapsedMilliseconds -gt $clock.ElapsedMilliseconds){throw ('顺序、状态或实时耗时不符：'+$kind)}
            if($kind -eq 'native-fault' -and -not $injected){throw '未完成实际 Word 故障注入。'}
            if((Get-FileHash -LiteralPath $occupied).Hash -ne $occupiedHash){throw '覆盖了已有输出。'}
            for($i=0;$i -lt $paths.Count;$i++){if((Get-FileHash -LiteralPath $paths[$i]).Hash -ne $hashes[$i]){throw '修改了原件。'}}
            $records=@(Get-ChildItem -LiteralPath $folder -Filter 'item-*.json')
            if($records.Count -ne $(if($kind -eq 'cancel'){2}else{3})){throw '任务数量不符。'}
            foreach($record in $records){$task=Get-Content -LiteralPath $record.FullName -Raw|ConvertFrom-Json;if(-not $task.Cleaned -or (Test-Path -LiteralPath $task.TaskDirectory)){throw 'GUI 任务残留未清理。'}}
            if(@(Get-ChildItem -LiteralPath $out -Filter '*.work-*').Count -or (Test-Path -LiteralPath (Join-Path $out ('中间文档'+$suffix+'.docx')))){throw '失败或取消文档留下了结果或工作副本。'}
            $allowed=@($occupied)+@($result.Items|Where-Object State -eq 2|ForEach-Object Output)
            if(@(Get-ChildItem -LiteralPath $out -Force|Where-Object FullName -notin $allowed).Count){throw '输出目录出现任务残留或额外文件。'}
            foreach($item in $result.Items|Where-Object State -eq 2){$audit=(& python -B -X utf8 (Join-Path $PSScriptRoot 'helpers/audit_docx.py') $item.Output)|ConvertFrom-Json;if($LASTEXITCODE -ne 0 -or $audit.MathType -ne $(if($kind -eq 'reverse'){0}else{5})){throw '成功结果未通过包审计。'}}
            $evidence.Cases+=@{Name=$kind;States=$states;ElapsedMilliseconds=$result.ElapsedMilliseconds;NativeFaultInjected=$injected;CleanupVerified=$true;PerDocumentLogVerified=$true;OriginalsUnchanged=$true;OrderVerified=$true;TotalTimeVerified=$true;RealtimeElapsedVerified=$true;RunningTimeUpdates=$liveElapsedUpdates;Check=(Join-Path $folder 'batch-check.json')}
            Write-Output ('BATCH_CASE_PASSED: '+$kind)
        }finally{Stop-BatchTestWindow $process}
    }
    $evidence.Status='PASSED'
}finally{
    $evidence.SourceSHA256=Get-TestSourceHashes
    $evidence|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $run 'batch-results.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'batch-results.json') -Encoding utf8
}
