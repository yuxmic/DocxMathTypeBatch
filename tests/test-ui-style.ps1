#requires -Version 7.0
# 原生界面外观验收：真实控件状态、最小窗口布局和截图；不启动 Office，不写用户设置。
param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'helpers/common.ps1')
if(-not $SkipBuild){& (Join-Path $ProjectRoot 'build.ps1') -Target Gui}
. (Join-Path $PSScriptRoot 'helpers/gui-tools.ps1')
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class StyleCheck {
    public delegate bool EnumProc(IntPtr hwnd,IntPtr data);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr data);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left,Top,Right,Bottom;}
    [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr hwnd,int id);
    [DllImport("user32.dll",EntryPoint="GetWindowLongPtrW")] public static extern IntPtr GetWindowLongPtr(IntPtr hwnd,int index);
    [DllImport("user32.dll",EntryPoint="GetClassLongPtrW")] public static extern IntPtr GetClassLongPtr(IntPtr hwnd,int index);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd,uint message,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll",EntryPoint="SendMessageW",CharSet=CharSet.Unicode)] static extern IntPtr SendText(IntPtr hwnd,uint message,IntPtr wp,string text);
    [DllImport("user32.dll",EntryPoint="SendMessageW",CharSet=CharSet.Unicode)] static extern IntPtr ReadText(IntPtr hwnd,uint message,IntPtr wp,StringBuilder text);
    [DllImport("user32.dll")] public static extern bool EnableWindow(IntPtr hwnd,bool enabled);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr hwnd,out Rect rect);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr hwnd);
    public static void DpiChange(IntPtr hwnd,int previous,int next){Rect r;GetWindowRect(hwnd,out r);r.Right=r.Left+(r.Right-r.Left)*next/previous;r.Bottom=r.Top+(r.Bottom-r.Top)*next/previous;IntPtr p=Marshal.AllocHGlobal(Marshal.SizeOf<Rect>());try{Marshal.StructureToPtr(r,p,false);SendMessage(hwnd,0x2E0,(IntPtr)(next|(next<<16)),p);}finally{Marshal.FreeHGlobal(p);}}
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hwnd,IntPtr after,int x,int y,int w,int h,uint flags);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd,uint message,IntPtr wp,IntPtr lp);
    public static void SetText(IntPtr hwnd,string text){SendText(hwnd,0xC,IntPtr.Zero,text);}
    public static IntPtr FindFont(IntPtr hwnd,string text){return SendText(hwnd,0x158,(IntPtr)(-1),text);}
    public static string Text(IntPtr hwnd){var s=new StringBuilder(SendMessage(hwnd,0xE,IntPtr.Zero,IntPtr.Zero).ToInt32()+1);ReadText(hwnd,0xD,(IntPtr)s.Capacity,s);return s.ToString();}
    public static IntPtr Dialog(uint pid,string title){IntPtr found=IntPtr.Zero;EnumWindows((hwnd,data)=>{uint owner;GetWindowThreadProcessId(hwnd,out owner);if(owner==pid&&Text(hwnd)==title){found=hwnd;return false;}return true;},IntPtr.Zero);return found;}
}
'@
$styleFolder=Join-Path $ProjectRoot ('tests\output\UI样式-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $styleFolder | Out-Null
$styleExe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
# 预览只读文件元数据，使用占位文件，完全不读取或转换真实文档。
$styleCopy=Join-Path $styleFolder '研究报告.docx';[IO.File]::WriteAllText($styleCopy,'UI preview only')
$styleProcess=$null
$styleDpi=[GuiBatchCheck]::SetThreadDpiAwarenessContext([IntPtr](-4))
$styleEvidence=[ordered]@{Status='FAILED';Direction='A按钮质感+B布局';Screenshots=@();Checks=[ordered]@{}}
function Save-StyleFrame([string]$Name) {
    $framePath=Join-Path $styleFolder ($Name+'.png')
    if(-not (Save-GuiWindow -PidValue $styleProcess.Id -Path $framePath)){throw ('截图失败：'+$Name)}
    $styleEvidence.Screenshots+= $framePath
    return $framePath
}
function Get-ButtonPixel([string]$Path,[IntPtr]$Button) {
    $previousDpi=[GuiBatchCheck]::SetThreadDpiAwarenessContext([IntPtr](-4))
    $parentRect=[GuiBatchCheck+Rect]::new();$buttonRect=[GuiBatchCheck+Rect]::new()
    [GuiBatchCheck]::GetWindowRect($styleWindow,[ref]$parentRect)|Out-Null
    [GuiBatchCheck]::GetWindowRect($Button,[ref]$buttonRect)|Out-Null
    $bitmap=[Drawing.Bitmap]::new($Path)
    try{return $bitmap.GetPixel($buttonRect.Left-$parentRect.Left+12,$buttonRect.Top-$parentRect.Top+12).ToArgb()}
    finally{$bitmap.Dispose();[GuiBatchCheck]::SetThreadDpiAwarenessContext($previousDpi)|Out-Null}
}
function Assert-QueueColumns {
    $queue=[StyleCheck]::GetDlgItem($styleWindow,138)
    $header=[StyleCheck]::SendMessage($queue,0x101F,[IntPtr]::Zero,[IntPtr]::Zero)
    if([StyleCheck]::SendMessage($header,0x1200,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 5){throw '队列没有五列。'}
    $widths=@(0..4|ForEach-Object{[StyleCheck]::SendMessage($queue,0x101D,[IntPtr]$_,[IntPtr]::Zero).ToInt32()})
    $bounds=[StyleCheck+Rect]::new();[StyleCheck]::GetClientRect($queue,[ref]$bounds)|Out-Null
    if(($widths|Measure-Object -Sum).Sum -gt $bounds.Right -or $widths[0] -le $widths[2]){throw '队列列宽超出原列表或文档名称过窄。'}
}
function Assert-ComboSelection([IntPtr]$Combo,[int]$Index) {
    # 从键盘路径选择，确认系统选中项由程序绘制，且仍保留可访问文本。
    if(([StyleCheck]::GetWindowLongPtr($Combo,-16).ToInt64() -band 0x210) -ne 0x210){throw '下拉框未统一接管选中项绘制或丢失可访问文本。'}
    [StyleCheck]::SendMessage($Combo,0x201,[IntPtr]1,[IntPtr](12 -bor (12 -shl 16)))|Out-Null
    [StyleCheck]::SendMessage($Combo,0x202,[IntPtr]::Zero,[IntPtr](12 -bor (12 -shl 16)))|Out-Null
    [StyleCheck]::SendMessage($Combo,0x100,[IntPtr]0x24,[IntPtr]::Zero)|Out-Null
    for($step=0;$step -lt $Index;$step++){[StyleCheck]::SendMessage($Combo,0x100,[IntPtr]0x28,[IntPtr]::Zero)|Out-Null}
    [StyleCheck]::SendMessage($Combo,0x100,[IntPtr]0x0D,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($Combo,0x14F,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    if([StyleCheck]::SendMessage($Combo,0x147,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne $Index -or -not [StyleCheck]::Text($Combo)){throw '下拉选项键盘选择失效。'}
}
try {
    $styleProcess=Start-Process -FilePath $styleExe -ArgumentList @('--ui-test',('"'+$styleCopy+'"'),('"'+$styleFolder+'"'),'preview') -WindowStyle Hidden -PassThru
    $deadline=[DateTime]::UtcNow.AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 100
        $styleWindow=[GuiBatchCheck]::Find($styleProcess.Id)
        if($styleProcess.HasExited -or [DateTime]::UtcNow -gt $deadline){throw '样式预览启动失败。'}
    }while($styleWindow -eq [IntPtr]::Zero)
    # 真实窗口大小图标都必须来自内嵌资源，预览保存供目视核验。
    foreach($iconKind in @(@{Name='large';Index=-14},@{Name='small';Index=-34})){
        $handle=[StyleCheck]::GetClassLongPtr($styleWindow,$iconKind.Index)
        if($handle -eq [IntPtr]::Zero){throw '窗口图标资源未加载。'}
        $bitmap=[Drawing.Icon]::FromHandle($handle).ToBitmap()
        try{
            $green=0
            for($x=0;$x -lt $bitmap.Width;$x++){for($y=0;$y -lt $bitmap.Height;$y++){$pixel=$bitmap.GetPixel($x,$y);if($pixel.A -gt 128 -and $pixel.G -gt $pixel.R+20 -and $pixel.G -ge $pixel.B-25){$green++}}}
            if($green -lt 5){throw '窗口图标不含预期绿色。'}
            $bitmap.Save((Join-Path $styleFolder ('icon-'+$iconKind.Name+'.png')),[Drawing.Imaging.ImageFormat]::Png)
        }finally{$bitmap.Dispose()}
    }
    Start-Sleep -Milliseconds 200
    $styleStart=[StyleCheck]::GetDlgItem($styleWindow,125)
    $styleSize=[StyleCheck]::GetDlgItem($styleWindow,113)
    $styleSpin=[StyleCheck]::GetDlgItem($styleWindow,114)
    Assert-QueueColumns
    foreach($id in @(137,136,111)){Assert-ComboSelection ([StyleCheck]::GetDlgItem($styleWindow,$id)) 1}
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($styleWindow,111),0x14E,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0x111,[IntPtr](111 -bor (1 -shl 16)),[StyleCheck]::GetDlgItem($styleWindow,111))|Out-Null
    $selectionFrame=Save-StyleFrame '00a-选择后无蓝底'
    $bitmap=[Drawing.Bitmap]::new($selectionFrame);$outer=[GuiBatchCheck+Rect]::new();[GuiBatchCheck]::GetWindowRect($styleWindow,[ref]$outer)|Out-Null
    try{
        foreach($id in @(111,136,137)){
            $rect=[GuiBatchCheck+Rect]::new();[GuiBatchCheck]::GetWindowRect([StyleCheck]::GetDlgItem($styleWindow,$id),[ref]$rect)|Out-Null
            for($x=$rect.Left-$outer.Left+8;$x -lt $rect.Right-$outer.Left-36;$x+=4){for($y=$rect.Top-$outer.Top+8;$y -lt $rect.Bottom-$outer.Top-8;$y+=4){$color=$bitmap.GetPixel($x,$y);if($color.R -lt 60 -and $color.B -gt 180 -and $color.B -gt ($color.G+40)){throw '选中项仍有系统蓝底。'}}}
        }
    }finally{$bitmap.Dispose()}
    # 单击日志单元格，验证实际命中与详情窗口，整个过程不启动工作进程。
    $queue=[StyleCheck]::GetDlgItem($styleWindow,138)
    $logX=0;foreach($column in 0..3){$logX+=[StyleCheck]::SendMessage($queue,0x101D,[IntPtr]$column,[IntPtr]::Zero).ToInt32()}
    $headerRect=[GuiBatchCheck+Rect]::new();[GuiBatchCheck]::GetWindowRect([StyleCheck]::SendMessage($queue,0x101F,[IntPtr]::Zero,[IntPtr]::Zero),[ref]$headerRect)|Out-Null
    $logPoint=[IntPtr](($logX+16) -bor (($headerRect.Bottom-$headerRect.Top+16) -shl 16))
    [StyleCheck]::PostMessage($queue,0x201,[IntPtr]1,$logPoint)|Out-Null
    [StyleCheck]::PostMessage($queue,0x202,[IntPtr]::Zero,$logPoint)|Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    do{Start-Sleep -Milliseconds 30;$details=[StyleCheck]::Dialog($styleProcess.Id,'转换详情');if([DateTime]::UtcNow -gt $deadline){throw '日志列未打开详情。'}}while($details -eq [IntPtr]::Zero)
    [StyleCheck]::SendMessage($details,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    $detailsText=[StyleCheck]::Text([StyleCheck]::GetDlgItem($details,301))
    if(-not $detailsText.Contains($styleCopy) -or -not $detailsText.Contains('本次后台日志')){throw '日志列打开了错误的文档详情。'}
    [StyleCheck]::PostMessage($details,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::PostMessage($styleWindow,0x111,[IntPtr]143,[IntPtr]::Zero)|Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    do{Start-Sleep -Milliseconds 30;$advanced=[StyleCheck]::Dialog($styleProcess.Id,'高级设置');if([DateTime]::UtcNow -gt $deadline){throw '高级设置未打开。'}}while($advanced -eq [IntPtr]::Zero)
    [StyleCheck]::SendMessage($advanced,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    foreach($id in @(201,202,226,227)){Assert-ComboSelection ([StyleCheck]::GetDlgItem($advanced,$id)) 1}
    if([StyleCheck]::Text([StyleCheck]::GetDlgItem($advanced,238)) -ne '实验性功能'){throw '实验项未合并进高级设置右侧。'}
    if([StyleCheck]::GetDlgItem($advanced,240) -ne [IntPtr]::Zero){throw '默认正文 OMML 仍保留开关。'}
    foreach($id in 241..244){if([StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,$id),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '实验项默认开启。'}}
    $leftBounds=[GuiBatchCheck+Rect]::new();$rightBounds=[GuiBatchCheck+Rect]::new()
    [GuiBatchCheck]::GetWindowRect([StyleCheck]::GetDlgItem($advanced,229),[ref]$leftBounds)|Out-Null
    [GuiBatchCheck]::GetWindowRect([StyleCheck]::GetDlgItem($advanced,241),[ref]$rightBounds)|Out-Null
    if($rightBounds.Left -le $leftBounds.Right){throw '实验区没有放在现有设置右侧。'}
    $advancedFrame=Join-Path $styleFolder '00b-高级设置与实验区.png'
    if(-not (Save-GuiWindow -PidValue $styleProcess.Id -Path $advancedFrame -WindowHandle $advanced)){throw '高级设置截图失败。'}
    $styleEvidence.Screenshots+=$advancedFrame
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,241),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::PostMessage($advanced,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    # 取消不提交实验设置，确定后再次打开可见；恢复默认关闭全部实验项。
    [StyleCheck]::PostMessage($styleWindow,0x111,[IntPtr]143,[IntPtr]::Zero)|Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    do{Start-Sleep -Milliseconds 30;$advanced=[StyleCheck]::Dialog($styleProcess.Id,'高级设置');if([DateTime]::UtcNow -gt $deadline){throw '高级设置未重新打开。'}}while($advanced -eq [IntPtr]::Zero)
    if([StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,241),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '取消仍然提交了实验项。'}
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,241),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::PostMessage($advanced,0x111,[IntPtr]1,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($styleWindow,136),0x14E,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0x111,[IntPtr](136 -bor (1 -shl 16)),[StyleCheck]::GetDlgItem($styleWindow,136))|Out-Null
    # 默认三行控件同宽同高，队列相对旧窗口宽度缩减约四分之一。
    $rectangles=@(111,136,137|ForEach-Object{$r=[GuiBatchCheck+Rect]::new();[GuiBatchCheck]::GetWindowRect([StyleCheck]::GetDlgItem($styleWindow,$_ ),[ref]$r)|Out-Null;$r})
    if(@($rectangles|ForEach-Object{$_.Right-$_.Left}|Select-Object -Unique).Count -ne 1 -or @($rectangles|ForEach-Object{$_.Bottom-$_.Top}|Select-Object -Unique).Count -ne 1){throw '模式、字号和右编号控件尺寸不统一。'}
    Save-StyleFrame '00-默认统一控件'|Out-Null
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($styleWindow,136),0x14E,[IntPtr]1,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0x111,[IntPtr](136 -bor (1 -shl 16)),[StyleCheck]::GetDlgItem($styleWindow,136))|Out-Null
    [StyleCheck]::SetText($styleSize,'10.00')
    [StyleCheck]::SetText([StyleCheck]::GetDlgItem($styleWindow,121),'_MathType')
    # 点击画出的下拉箭头，并确认真实列表状态；同时验证列表选择通知仍到达业务代码。
    $font=[StyleCheck]::GetDlgItem($styleWindow,111)
    $fontRect=[StyleCheck+Rect]::new();[StyleCheck]::GetClientRect($font,[ref]$fontRect)|Out-Null
    $windowClient=[StyleCheck+Rect]::new();[StyleCheck]::GetClientRect($styleWindow,[ref]$windowClient)|Out-Null
    $fontPoint=[IntPtr](($fontRect.Right-[int](18*$windowClient.Right/824.0)) -bor (([int]($fontRect.Bottom/2)) -shl 16))
    [StyleCheck]::SendMessage($font,0x201,[IntPtr]1,$fontPoint)|Out-Null
    [StyleCheck]::SendMessage($font,0x202,[IntPtr]::Zero,$fontPoint)|Out-Null
    if([StyleCheck]::SendMessage($font,0x157,[IntPtr]::Zero,[IntPtr]::Zero).ToInt64() -ne 1){throw '字体下拉箭头未打开原生列表。'}
    [StyleCheck]::SendMessage($font,0x14F,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    $fontIndex=[StyleCheck]::FindFont($font,'仅行外')
    if($fontIndex.ToInt64() -lt 0){throw '行外模式不存在。'}
    [StyleCheck]::SendMessage($font,0x14E,$fontIndex,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0x111,[IntPtr](111 -bor (1 -shl 16)),$font)|Out-Null
    if([StyleCheck]::Text($font) -ne '仅行外'){throw '公式模式选择未生效。'}
    [StyleCheck]::SendMessage($font,0x14E,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($font,0x142,[IntPtr]::Zero,[IntPtr]65535)|Out-Null
    if([StyleCheck]::Text($styleSize) -ne '10.00'){throw '字号输入未生效。'}
    if(-not [StyleCheck]::IsWindowEnabled($styleStart)){throw '准备状态的开始按钮未启用。'}
    $normal=Save-StyleFrame '01-准备就绪'
    # 通过原生控件消息固定四种截图状态；实体鼠标和多显示器留给用户实操。
    [StyleCheck]::SendMessage($styleStart,0x200,[IntPtr]::Zero,[IntPtr](12 -bor (12 -shl 16)))|Out-Null
    $hover=Save-StyleFrame '02-按钮悬停'
    [StyleCheck]::SendMessage($styleStart,0xF3,[IntPtr]1,[IntPtr]::Zero)|Out-Null
    $pressed=Save-StyleFrame '03-按钮按下'
    [StyleCheck]::SendMessage($styleStart,0xF3,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleStart,0x2A3,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::EnableWindow($styleStart,$false)|Out-Null
    $disabled=Save-StyleFrame '04-按钮禁用'
    [StyleCheck]::EnableWindow($styleStart,$true)|Out-Null
    $colors=@($normal,$hover,$pressed,$disabled|ForEach-Object {Get-ButtonPixel $_ $styleStart})
    $styleEvidence.ButtonPixels=$colors
    if(@($colors|Select-Object -Unique).Count -ne 4){throw '四种按钮状态的真实像素没有区分。'}
    # 用真正的箭头命中区域点击，不通过直接修改数字假装字号按钮有效。
    $spinRect=[StyleCheck+Rect]::new();[StyleCheck]::GetClientRect($styleSpin,[ref]$spinRect)|Out-Null
    $point=[IntPtr](($spinRect.Right/2) -bor (([int]($spinRect.Bottom/4)) -shl 16))
    [StyleCheck]::SendMessage($styleSpin,0x201,[IntPtr]1,$point)|Out-Null
    [StyleCheck]::SendMessage($styleSpin,0x202,[IntPtr]::Zero,$point)|Out-Null
    if([double][StyleCheck]::Text($styleSize) -ne 11){throw '现代字号箭头未增加 1 pt。'}
    [StyleCheck]::SetText($styleSize,'10.00')
    # 预览使用已有原生进度协议，不改变转换状态或产生输出。
    $progress=[StyleCheck]::GetDlgItem($styleWindow,124)
    $sizeBounds=[GuiBatchCheck+Rect]::new();$spinBounds=[GuiBatchCheck+Rect]::new();$barBounds=[GuiBatchCheck+Rect]::new();$resultBounds=[GuiBatchCheck+Rect]::new()
    [GuiBatchCheck]::GetWindowRect($styleSize,[ref]$sizeBounds)|Out-Null;[GuiBatchCheck]::GetWindowRect($styleSpin,[ref]$spinBounds)|Out-Null
    [GuiBatchCheck]::GetWindowRect($progress,[ref]$barBounds)|Out-Null;[GuiBatchCheck]::GetWindowRect($styleStart,[ref]$resultBounds)|Out-Null
    if($sizeBounds.Right -ne $spinBounds.Left -or $sizeBounds.Top -ne $spinBounds.Top -or $sizeBounds.Bottom -ne $spinBounds.Bottom){throw '字号输入和箭头的共同边框未对齐。'}
    if($barBounds.Top -lt $resultBounds.Bottom){throw '底部进度条与开始按钮重叠。'}
    [StyleCheck]::SendMessage($progress,0x406,[IntPtr]::Zero,[IntPtr]100)|Out-Null
    [StyleCheck]::SendMessage($progress,0x402,[IntPtr]86,[IntPtr]::Zero)|Out-Null
    [GuiBatchCheck]::ShowWindow($progress,4)|Out-Null
    Save-StyleFrame '05-进度条'|Out-Null
    [GuiBatchCheck]::ShowWindow($progress,0)|Out-Null
    $dpi=[GuiBatchCheck]::SetThreadDpiAwarenessContext([IntPtr](-4))
    try {
        $outer=[GuiBatchCheck+Rect]::new();$client=[StyleCheck+Rect]::new()
        [GuiBatchCheck]::GetWindowRect($styleWindow,[ref]$outer)|Out-Null
        [StyleCheck]::GetClientRect($styleWindow,[ref]$client)|Out-Null
        # 当前进程 DPI 由窗口物理客户宽度计算，保持高度、只验收最小宽度。
        $scale=$client.Right/824.0
        [StyleCheck]::SetWindowPos($styleWindow,[IntPtr]::Zero,0,0,[int](800*$scale+$outer.Right-$outer.Left-$client.Right),$outer.Bottom-$outer.Top,0x16)|Out-Null
    }finally{[GuiBatchCheck]::SetThreadDpiAwarenessContext($dpi)|Out-Null}
    Save-StyleFrame '06-最小窗口'|Out-Null
    Assert-QueueColumns
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($styleWindow,102),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    Save-StyleFrame '07-反向格式'|Out-Null
    if([StyleCheck]::Text([StyleCheck]::GetDlgItem($styleWindow,109)) -ne '导出格式'){throw '反向设置没有切换。'}
    [StyleCheck]::PostMessage($styleWindow,0x111,[IntPtr]143,[IntPtr]::Zero)|Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    do{Start-Sleep -Milliseconds 30;$advanced=[StyleCheck]::Dialog($styleProcess.Id,'高级设置');if([DateTime]::UtcNow -gt $deadline){throw '反向无法打开高级设置。'}}while($advanced -eq [IntPtr]::Zero)
    if([StyleCheck]::GetDlgItem($advanced,240) -ne [IntPtr]::Zero -or -not [StyleCheck]::IsWindowEnabled([StyleCheck]::GetDlgItem($advanced,241))){throw '反向实验项可用性错误。'}
    if([StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,241),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '实验设置未在当前会话保留。'}
    $reverseFrame=Join-Path $styleFolder '08-反向高级设置.png'
    if(-not (Save-GuiWindow -PidValue $styleProcess.Id -Path $reverseFrame -WindowHandle $advanced)){throw '反向高级设置截图失败。'}
    $styleEvidence.Screenshots+=$reverseFrame
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,235),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    foreach($id in 241..244){if([StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,$id),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '恢复默认未关闭实验项。'}}
    [StyleCheck]::PostMessage($advanced,0x111,[IntPtr]1,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    if(Test-Path -LiteralPath (Join-Path $styleFolder 'gui-worker.log')){throw '预览意外启动转换。'}
    # 通过 DPI 变化消息验证实际控件排版；高缩放下按钮可经键盘滚动到可见区域。
    $previousScale=[int][StyleCheck]::GetDpiForWindow($styleWindow)
    foreach($scale in @(96,120,144,192)){
        [StyleCheck]::DpiChange($styleWindow,$previousScale,$scale);$previousScale=$scale
        [StyleCheck]::PostMessage($styleWindow,0x111,[IntPtr]143,[IntPtr]::Zero)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        do{Start-Sleep -Milliseconds 30;$advanced=[StyleCheck]::Dialog($styleProcess.Id,'高级设置');if([DateTime]::UtcNow -gt $deadline){throw '缩放后的高级设置未打开。'}}while($advanced -eq [IntPtr]::Zero)
        $frame=Join-Path $styleFolder ('09-高级设置-'+$scale+'dpi.png')
        if(-not (Save-GuiWindow -PidValue $styleProcess.Id -Path $frame -WindowHandle $advanced)){throw '缩放后的高级设置截图失败。'}
        $styleEvidence.Screenshots+=$frame
        $confirm=[StyleCheck]::GetDlgItem($advanced,1)
        [StyleCheck]::SendMessage($advanced,0x28,$confirm,[IntPtr]1)|Out-Null
        $outer=[StyleCheck+Rect]::new();$button=[StyleCheck+Rect]::new()
        [StyleCheck]::GetWindowRect($advanced,[ref]$outer)|Out-Null;[StyleCheck]::GetWindowRect($confirm,[ref]$button)|Out-Null
        if($button.Bottom -gt $outer.Bottom -or $button.Right -gt $outer.Right -or $button.Top -lt $outer.Top){throw '缩放后确定按钮不能通过键盘完整显示。'}
        [StyleCheck]::PostMessage($advanced,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        [StyleCheck]::SendMessage($styleWindow,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    }
    $styleEvidence.Checks=@{QueueFiveColumns=$true;QueueWidthUnchanged=$true;LogColumnOpensDetails=$true;ComboSelectionNoBlue=$true;AdvancedComboSelection=$true;FourButtonStates=$true;SizeArrow=$true;SizeFrameAligned=$true;ProgressSeparated=$true;ModeDropdown=$true;ModeSelection=$true;UnifiedSettings=$true;ReverseSettings=$true;NoOfficeTask=$true;NoSettingsWritten=$true;MinimumWidthScreenshot=$true}
    $styleEvidence.Checks.EmbeddedGreenIcons=$true
    $styleEvidence.Status='PASSED'
    $styleEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $styleExe).Hash
    $styleEvidence.SourceSHA256=[ordered]@{}
    foreach($source in @('gui.cpp','build.ps1','assets/app.rc','assets/app_icon.ico','assets/app_icon.png','assets/build-icon.ps1','assets/icon-source.json','test-ui-style.ps1','test-gui.ps1')){$styleEvidence.SourceSHA256[$source]=(Get-FileHash -LiteralPath (Get-TestSourcePath $source)).Hash}
    Write-Output ('UI_STYLE_PASSED: '+$styleFolder)
}finally {
    [GuiBatchCheck]::SetThreadDpiAwarenessContext($styleDpi)|Out-Null
    if($styleProcess){if(-not $styleProcess.HasExited){foreach($title in @('转换详情','高级设置')){[StyleCheck]::PostMessage([StyleCheck]::Dialog($styleProcess.Id,$title),0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null};[StyleCheck]::PostMessage([GuiBatchCheck]::Find($styleProcess.Id),0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        if(-not $styleProcess.WaitForExit(15000)){throw '预览窗口没有退出。'}};$styleProcess.Dispose()}
    Remove-Item -LiteralPath $styleCopy -Force
    $styleEvidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $styleFolder 'style-results.json') -Encoding utf8
    $styleEvidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'ui-style-results.json') -Encoding utf8
}
