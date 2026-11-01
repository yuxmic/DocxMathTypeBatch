#requires -Version 7.0
# 原生界面外观验收：真实控件状态、最小窗口布局和截图；不启动 Office，不写用户设置。
param([switch]$SkipBuild,[ValidateSet('All','Fast','Help')][string]$Case='All')
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
    [DllImport("user32.dll",EntryPoint="GetClassNameW",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd,StringBuilder name,int capacity);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left,Top,Right,Bottom;}
    [StructLayout(LayoutKind.Sequential)] public struct Point {public int X,Y;}
    [DllImport("user32.dll")] public static extern bool GetCursorPos(out Point point);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern IntPtr ChildWindowFromPointEx(IntPtr parent,Point point,uint flags);
    [DllImport("user32.dll")] public static extern int MapWindowPoints(IntPtr from,IntPtr to,ref Point point,uint count);
    [DllImport("user32.dll")] static extern bool SystemParametersInfo(uint action,uint param,out bool value,uint flags);
    [DllImport("dwmapi.dll")] public static extern int DwmFlush();
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
    // 按屏幕坐标命中实际子控件，不用直接发送按钮命令替代鼠标命中检查。
    public static IntPtr ClickAtScreen(IntPtr parent,int x,int y){var p=new Point{X=x,Y=y};MapWindowPoints(IntPtr.Zero,parent,ref p,1);var hit=ChildWindowFromPointEx(parent,p,0);MapWindowPoints(parent,hit,ref p,1);var pos=(IntPtr)((p.X&65535)|((p.Y&65535)<<16));SendMessage(hit,0x201,(IntPtr)1,pos);SendMessage(hit,0x202,IntPtr.Zero,pos);return hit;}
    public static bool ClientAnimations(){bool value;return SystemParametersInfo(0x1042,0,out value,0)&&value;}
    public static IntPtr Dialog(uint pid,string title){IntPtr found=IntPtr.Zero;EnumWindows((hwnd,data)=>{uint owner;GetWindowThreadProcessId(hwnd,out owner);if(owner==pid&&Text(hwnd)==title){found=hwnd;return false;}return true;},IntPtr.Zero);return found;}
    public static IntPtr Tooltip(uint pid){IntPtr found=IntPtr.Zero;EnumWindows((hwnd,data)=>{uint owner;GetWindowThreadProcessId(hwnd,out owner);var name=new StringBuilder(64);GetClassName(hwnd,name,name.Capacity);if(owner==pid&&name.ToString()=="tooltips_class32"&&IsWindowVisible(hwnd)){found=hwnd;return false;}return true;},IntPtr.Zero);return found;}
}
'@
$styleFolder=Join-Path $ProjectRoot ('tests\output\UI样式-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $styleFolder | Out-Null
$styleExe=Join-Path $ProjectRoot 'build\bin\DocxMathTypeBatch.exe'
# 预览只读文件元数据，使用占位文件，完全不读取或转换真实文档。
$styleCopy=Join-Path $styleFolder '研究报告.docx';[IO.File]::WriteAllText($styleCopy,'UI preview only')
$styleProcess=$null
$styleDpi=[GuiBatchCheck]::SetThreadDpiAwarenessContext([IntPtr](-4))
$styleEvidence=[ordered]@{Status='FAILED';Case=$Case;Direction='A按钮质感+B布局';Screenshots=@();Checks=[ordered]@{}}
$fastHelpText='舍弃版式检查，提高转换速度，但存在错位风险'
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
function Save-FastFrame([string]$Name,[switch]$NoRedraw) {
    # 直接读取本次预览的滑块像素，避免整窗重绘耗时超过动画时长。
    if(-not $NoRedraw){[GuiBatchCheck]::RedrawWindow($styleFast,[IntPtr]::Zero,[IntPtr]::Zero,0x101)|Out-Null}
    [StyleCheck]::DwmFlush()|Out-Null
    $bounds=[StyleCheck+Rect]::new();[StyleCheck]::GetWindowRect($styleFast,[ref]$bounds)|Out-Null
    $crop=[Drawing.Bitmap]::new($bounds.Right-$bounds.Left,$bounds.Bottom-$bounds.Top);$graphics=[Drawing.Graphics]::FromImage($crop)
    try{
        $graphics.CopyFromScreen($bounds.Left,$bounds.Top,0,0,$crop.Size)
        $framePath=Join-Path $styleFolder ($Name+'.png');$crop.Save($framePath,[Drawing.Imaging.ImageFormat]::Png);$styleEvidence.Screenshots+=$framePath
        return $crop.GetPixel([int]($crop.Width/4),[int]($crop.Height/2)).ToArgb()
    }finally{$graphics.Dispose();$crop.Dispose()}
}
function Assert-FastHelp {
    # 问号保留完整可访问名称和 16 DIP 点击区，圈形约 12 DIP，允许轻微光学对齐调整。
    $help=[StyleCheck]::GetDlgItem($styleWindow,148)
    if($help -eq [IntPtr]::Zero -or [StyleCheck]::Text($help) -ne $fastHelpText -or
       ([StyleCheck]::GetWindowLongPtr($help,-16).ToInt64() -band 0x1000000F) -ne 0x1000000B){throw '快速转换帮助缺失、可访问名称不符或未采用可见自绘按钮。'}
    $helpBounds=[StyleCheck+Rect]::new();$toggleBounds=[StyleCheck+Rect]::new();$labelBounds=[StyleCheck+Rect]::new();$clientBounds=[StyleCheck+Rect]::new()
    [StyleCheck]::GetWindowRect($help,[ref]$helpBounds)|Out-Null
    [StyleCheck]::GetWindowRect($styleFast,[ref]$toggleBounds)|Out-Null
    [StyleCheck]::GetWindowRect([StyleCheck]::GetDlgItem($styleWindow,147),[ref]$labelBounds)|Out-Null
    [StyleCheck]::GetClientRect($styleWindow,[ref]$clientBounds)|Out-Null
    $origin=[StyleCheck+Point]::new();[StyleCheck]::MapWindowPoints($styleWindow,[IntPtr]::Zero,[ref]$origin,1)|Out-Null
    $scale=[StyleCheck]::GetDpiForWindow($styleWindow)/96.0
    $centerOffset=[Math]::Abs(($helpBounds.Top+$helpBounds.Bottom-$labelBounds.Top-$labelBounds.Bottom)/2.0)
    if([Math]::Abs($labelBounds.Right-$labelBounds.Left-52*$scale) -gt 1 -or [Math]::Abs($helpBounds.Left-$labelBounds.Left-56*$scale) -gt 1 -or
       $centerOffset -gt 2*$scale+1 -or [Math]::Abs($helpBounds.Right-$helpBounds.Left-16*$scale) -gt 1 -or
       [Math]::Abs($helpBounds.Bottom-$helpBounds.Top-16*$scale) -gt 1){throw '快速转换问号帮助尺寸或位置不符。'}
    if($helpBounds.Left -lt $labelBounds.Right -or $helpBounds.Right -gt $toggleBounds.Left -or
       $helpBounds.Right -gt $origin.X+$clientBounds.Right -or $helpBounds.Bottom -gt $origin.Y+$clientBounds.Bottom){throw '快速转换问号帮助未独立位于标签与滑块之间。'}
    return $helpBounds
}
function Assert-FastHelpTooltip {
    # 只移动鼠标到本轮窗口验证真实悬停；无论成败均恢复原鼠标位置。
    $savedCursor=[StyleCheck+Point]::new()
    if(-not [StyleCheck]::GetCursorPos([ref]$savedCursor)){throw '无法保存鼠标位置。'}
    try{
        $helpBounds=Assert-FastHelp
        $labelBounds=[StyleCheck+Rect]::new();[StyleCheck]::GetWindowRect([StyleCheck]::GetDlgItem($styleWindow,147),[ref]$labelBounds)|Out-Null
        $awayX=$labelBounds.Left+2;$awayY=[int](($labelBounds.Top+$labelBounds.Bottom)/2)
        [StyleCheck]::SetCursorPos($awayX,$awayY)|Out-Null
        # 留出一次真实移出消息处理时间，避免紧接的移入被系统合并。
        Start-Sleep -Milliseconds 100
        $deadline=[DateTime]::UtcNow.AddSeconds(5)
        do{Start-Sleep -Milliseconds 50;$tooltip=[StyleCheck]::Tooltip($styleProcess.Id);if([DateTime]::UtcNow -gt $deadline){throw '帮助提示离开问号后仍常驻显示。'}}while($tooltip -ne [IntPtr]::Zero)
        [StyleCheck]::SetCursorPos([int](($helpBounds.Left+$helpBounds.Right)/2),[int](($helpBounds.Top+$helpBounds.Bottom)/2))|Out-Null
        # 指针已真实移入，再通知帮助控件中心，仍由系统悬停计时显示提示。
        $helpPoint=[IntPtr](([int](($helpBounds.Right-$helpBounds.Left)/2)) -bor (([int](($helpBounds.Bottom-$helpBounds.Top)/2)) -shl 16))
        [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($styleWindow,148),0x200,[IntPtr]::Zero,$helpPoint)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(5)
        do{Start-Sleep -Milliseconds 50;$tooltip=[StyleCheck]::Tooltip($styleProcess.Id);if([DateTime]::UtcNow -gt $deadline){throw '真实悬停问号后未显示原生帮助提示。'}}while($tooltip -eq [IntPtr]::Zero)
        $tooltipText=[StyleCheck]::Text($tooltip)
        if($tooltipText -ne $fastHelpText){throw '原生帮助提示未显示完整风险说明。'}
        # 等待系统提示淡入结束，避免截图只取得接近透明的第一帧。
        Start-Sleep -Milliseconds 200
        [StyleCheck]::DwmFlush()|Out-Null
        $bounds=[StyleCheck+Rect]::new();[StyleCheck]::GetWindowRect($tooltip,[ref]$bounds)|Out-Null
        $framePath=Join-Path $styleFolder '00k-快速帮助悬停提示.png'
        $bitmap=[Drawing.Bitmap]::new($bounds.Right-$bounds.Left,$bounds.Bottom-$bounds.Top);$graphics=[Drawing.Graphics]::FromImage($bitmap)
        try{$graphics.CopyFromScreen($bounds.Left,$bounds.Top,0,0,$bitmap.Size);$bitmap.Save($framePath,[Drawing.Imaging.ImageFormat]::Png)}
        finally{$graphics.Dispose();$bitmap.Dispose()}
        $styleEvidence.Screenshots+=$framePath
        [StyleCheck]::SetCursorPos($awayX,$awayY)|Out-Null
        $deadline=[DateTime]::UtcNow.AddSeconds(5)
        do{Start-Sleep -Milliseconds 50;if([DateTime]::UtcNow -gt $deadline){throw '移开鼠标后帮助提示未自动隐藏。'}}while([StyleCheck]::IsWindowVisible($tooltip))
        $styleEvidence.FastHelpTooltip=@{Text=$tooltipText;WindowClass='tooltips_class32';VisibleOnHover=$true;HiddenOnLeave=$true;Screenshot=$framePath}
    }finally{[StyleCheck]::SetCursorPos($savedCursor.X,$savedCursor.Y)|Out-Null}
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
    # Fast 定向入口只检查首页开关，不执行图标和高级设置等完整样式组。
    if($Case -eq 'All'){
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
    }
    Start-Sleep -Milliseconds 200
    $styleStart=[StyleCheck]::GetDlgItem($styleWindow,125)
    $styleSize=[StyleCheck]::GetDlgItem($styleWindow,113)
    $styleSpin=[StyleCheck]::GetDlgItem($styleWindow,114)
    # 首页开关保留原生勾选状态及空格键操作，默认关闭。
    $styleFast=[StyleCheck]::GetDlgItem($styleWindow,146)
    if($styleFast -eq [IntPtr]::Zero -or ([StyleCheck]::GetWindowLongPtr($styleFast,-16).ToInt64() -band 0x1001000F) -ne 0x10010003){throw '首页快速转换未使用原生可访问开关。'}
    if([StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '快速转换默认开启。'}
    $fastBounds=[StyleCheck+Rect]::new();$numberBounds=[StyleCheck+Rect]::new();$outputBounds=[StyleCheck+Rect]::new();$fastLabelBounds=[StyleCheck+Rect]::new()
    [StyleCheck]::GetWindowRect($styleFast,[ref]$fastBounds)|Out-Null
    [StyleCheck]::GetWindowRect([StyleCheck]::GetDlgItem($styleWindow,137),[ref]$numberBounds)|Out-Null
    [StyleCheck]::GetWindowRect([StyleCheck]::GetDlgItem($styleWindow,116),[ref]$outputBounds)|Out-Null
    [StyleCheck]::GetWindowRect([StyleCheck]::GetDlgItem($styleWindow,147),[ref]$fastLabelBounds)|Out-Null
    if($fastBounds.Top -le $numberBounds.Bottom -or $fastBounds.Bottom -ge $outputBounds.Top -or $fastLabelBounds.Left -ne $outputBounds.Left -or $fastBounds.Left -ne $numberBounds.Left){throw '快速转换未在设置末行按标签及输入控件对齐。'}
    $fastHelpBounds=Assert-FastHelp
    if($Case -eq 'Help'){
        # 问号专项只验证本次改动，不依赖动画中间帧或其他控件的焦点截图。
        [GuiBatchCheck]::ShowWindow($styleWindow,4)|Out-Null
        [StyleCheck]::SetWindowPos($styleWindow,[IntPtr](-1),0,0,0,0,0x13)|Out-Null
        Assert-FastHelpTooltip
        $outer=[StyleCheck+Rect]::new();$client=[StyleCheck+Rect]::new()
        [StyleCheck]::GetWindowRect($styleWindow,[ref]$outer)|Out-Null;[StyleCheck]::GetClientRect($styleWindow,[ref]$client)|Out-Null
        $scale=[StyleCheck]::GetDpiForWindow($styleWindow)/96.0
        [StyleCheck]::SetWindowPos($styleWindow,[IntPtr]::Zero,0,0,[int](800*$scale+$outer.Right-$outer.Left-$client.Right),$outer.Bottom-$outer.Top,0x16)|Out-Null
        Assert-FastHelp|Out-Null
        Save-StyleFrame '00j-最小800宽快速帮助'|Out-Null
        $styleEvidence.Checks=@{FastHelpAccessibleName=$true;FastHelpPosition=$true;FastHelpTooltipOnHover=$true;FastHelpTooltipHidesOnLeave=$true;MinimumFastHelpVisible=$true;NoOfficeTask=$true}
        $styleEvidence.Status='PASSED';$styleEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $styleExe).Hash
        $styleEvidence.SourceSHA256=Get-TestSourceHashes
        Write-Output ('UI_STYLE_PASSED: '+$styleFolder)
        return
    }
    # 点击文字、右侧空白都不应切换；只有滑块本身命中按钮。
    $fastY=[int](($fastBounds.Top+$fastBounds.Bottom)/2)
    foreach($fastX in @(($fastLabelBounds.Left+4),([int](($fastHelpBounds.Left+$fastHelpBounds.Right)/2)),($outputBounds.Right-4))){
        $hit=[StyleCheck]::ClickAtScreen($styleWindow,$fastX,$fastY)
        if($hit -eq $styleFast -or [StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '点击快速转换文字或同行空白仍会触发开关。'}
    }
    $hit=[StyleCheck]::ClickAtScreen($styleWindow,[int](($fastBounds.Left+$fastBounds.Right)/2),$fastY)
    if($hit -ne $styleFast -or [StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '点击滑块未切换。'}
    [StyleCheck]::SendMessage($styleFast,0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    Start-Sleep -Milliseconds 220
    # 动画期间保持预览窗口可见，避免反复显示窗口的系统动画耗尽中间帧时间。
    [GuiBatchCheck]::ShowWindow($styleWindow,4)|Out-Null
    [StyleCheck]::SetWindowPos($styleWindow,[IntPtr](-1),0,0,0,0,0x13)|Out-Null
    Save-StyleFrame '00c-首页快速转换关闭'|Out-Null
    $fastOffPixel=Save-FastFrame '00e-开关关闭细节'
    [StyleCheck]::SendMessage($styleWindow,0x28,$styleFast,[IntPtr]1)|Out-Null
    [StyleCheck]::SendMessage($styleFast,0x100,[IntPtr]0x20,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleFast,0x101,[IntPtr]0x20,[IntPtr]::Zero)|Out-Null
    if([StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '快速开关空格键未生效。'}
    Start-Sleep -Milliseconds 60
    $fastMidPixel=Save-FastFrame '00f-开关切换中间帧'
    # 动画期间再次关闭再打开，确认最终状态正确且视觉能继续过渡。
    [StyleCheck]::SendMessage($styleFast,0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleFast,0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    Start-Sleep -Milliseconds 220
    $fastOnPixel=Save-FastFrame '00g-开关开启细节'
    if($Case -eq 'All' -and [StyleCheck]::ClientAnimations() -and @($fastOffPixel,$fastMidPixel,$fastOnPixel|Select-Object -Unique).Count -ne 3){throw '快速开关未呈现独立的动画中间帧。'}
    if([StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '连续切换后状态错误。'}
    Save-StyleFrame '00d-首页快速转换开启'|Out-Null
    # 显式显示键盘焦点，再隐藏焦点提示；同一开关状态的两帧用于核验实线轮廓。
    foreach($target in @($styleWindow,$styleFast)){[StyleCheck]::SendMessage($target,0x127,[IntPtr](2 -bor (1 -shl 16)),[IntPtr]::Zero)|Out-Null}
    [StyleCheck]::SendMessage($styleWindow,0x28,$styleFast,[IntPtr]1)|Out-Null
    if(([StyleCheck]::SendMessage($styleFast,0x129,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -band 1) -ne 0 -or
       ([StyleCheck]::SendMessage($styleFast,0xF2,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -band 8) -eq 0){throw '快速开关未取得可见的键盘焦点。'}
    # 此处不由截图辅助函数补画，保留焦点消息处理后的真实即时外观。
    Save-FastFrame '00h-开关实线焦点' -NoRedraw|Out-Null
    foreach($target in @($styleWindow,$styleFast)){[StyleCheck]::SendMessage($target,0x127,[IntPtr](1 -bor (1 -shl 16)),[IntPtr]::Zero)|Out-Null}
    if(([StyleCheck]::SendMessage($styleFast,0x129,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -band 1) -eq 0){throw '快速开关未隐藏焦点提示。'}
    Save-FastFrame '00i-开关隐藏焦点' -NoRedraw|Out-Null
    $focusHash=(Get-FileHash -LiteralPath (Join-Path $styleFolder '00h-开关实线焦点.png')).Hash
    $hiddenFocusHash=(Get-FileHash -LiteralPath (Join-Path $styleFolder '00i-开关隐藏焦点.png')).Hash
    if($focusHash -eq $hiddenFocusHash){throw '快速开关显示及隐藏键盘焦点的两帧相同。'}
    $styleEvidence.FastFocusFrames=@{VisibleSHA256=$focusHash;HiddenSHA256=$hiddenFocusHash}
    Assert-FastHelpTooltip
    if($Case -eq 'Fast'){
        # 定向入口仅增加一次 800 DIP 客户宽度检查，随后由原 finally 关闭预览并保存证据。
        $outer=[StyleCheck+Rect]::new();$client=[StyleCheck+Rect]::new()
        [StyleCheck]::GetWindowRect($styleWindow,[ref]$outer)|Out-Null;[StyleCheck]::GetClientRect($styleWindow,[ref]$client)|Out-Null
        $scale=[StyleCheck]::GetDpiForWindow($styleWindow)/96.0
        [StyleCheck]::SetWindowPos($styleWindow,[IntPtr]::Zero,0,0,[int](800*$scale+$outer.Right-$outer.Left-$client.Right),$outer.Bottom-$outer.Top,0x16)|Out-Null
        Assert-FastHelp|Out-Null
        Save-StyleFrame '00j-最小800宽快速帮助'|Out-Null
        if(Test-Path -LiteralPath (Join-Path $styleFolder 'gui-worker.log')){throw '预览意外启动转换。'}
        $styleEvidence.Checks=@{HomeFastToggle=$true;FastHelpAccessibleName=$true;FastHelpPosition=$true;FastToggleHitArea=$true;FastHelpClickDoesNotToggle=$true;FastHelpTooltipOnHover=$true;FastHelpTooltipHidesOnLeave=$true;FastToggleSpaceKey=$true;FastToggleFocusCue=$true;FastFocusFramesDiffer=$true;FastToggleAnimationFramesCaptured=$true;MinimumFastHelpVisible=$true;NoOfficeTask=$true}
        $styleEvidence.Status='PASSED'
        $styleEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $styleExe).Hash
        $styleEvidence.SourceSHA256=Get-TestSourceHashes
        Write-Output ('UI_STYLE_PASSED: '+$styleFolder)
        return
    }
    [GuiBatchCheck]::ShowWindow($styleWindow,0)|Out-Null
    [StyleCheck]::SetWindowPos($styleWindow,[IntPtr](-2),0,0,0,0,0x13)|Out-Null
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
    if(-not $detailsText.Contains([IO.Path]::GetFileName($styleCopy)) -or -not $detailsText.Contains('条问题') -or $detailsText.Contains($styleCopy) -or $detailsText.Contains('本次后台日志')){throw '日志列未默认显示对应文档的简短问题摘要。'}
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($details,303),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    $detailsText=[StyleCheck]::Text([StyleCheck]::GetDlgItem($details,301))
    if(-not $detailsText.Contains($styleCopy) -or -not $detailsText.Contains('本次后台日志')){throw '切换原始日志后未显示对应文档的完整技术详情。'}
    [StyleCheck]::PostMessage($details,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::PostMessage($styleWindow,0x111,[IntPtr]143,[IntPtr]::Zero)|Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    do{Start-Sleep -Milliseconds 30;$advanced=[StyleCheck]::Dialog($styleProcess.Id,'高级设置');if([DateTime]::UtcNow -gt $deadline){throw '高级设置未打开。'}}while($advanced -eq [IntPtr]::Zero)
    [StyleCheck]::SendMessage($advanced,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    foreach($id in @(201,202,226,227)){Assert-ComboSelection ([StyleCheck]::GetDlgItem($advanced,$id)) 1}
    if([StyleCheck]::Text([StyleCheck]::GetDlgItem($advanced,238)) -ne '实验性功能'){throw '实验项未合并进高级设置右侧。'}
    foreach($id in @(239,247,248,249,250)){if([StyleCheck]::GetDlgItem($advanced,$id) -ne [IntPtr]::Zero){throw '实验区仍有副标题、说明文字或快速转换控件。'}}
    if([StyleCheck]::GetDlgItem($advanced,246) -eq [IntPtr]::Zero -or [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,246),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw 'Word 原生公式（OMML）未默认勾选。'}
    if([StyleCheck]::GetDlgItem($advanced,240) -ne [IntPtr]::Zero){throw '旧版 OMML 控件仍存在。'}
    foreach($id in 241..244){if([StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,$id),0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '实验项默认开启。'}}
    $leftBounds=[GuiBatchCheck+Rect]::new();$rightBounds=[GuiBatchCheck+Rect]::new()
    [GuiBatchCheck]::GetWindowRect([StyleCheck]::GetDlgItem($advanced,229),[ref]$leftBounds)|Out-Null
    [GuiBatchCheck]::GetWindowRect([StyleCheck]::GetDlgItem($advanced,241),[ref]$rightBounds)|Out-Null
    if($rightBounds.Left -le $leftBounds.Right){throw '实验区没有放在现有设置右侧。'}
    $advancedFrame=Join-Path $styleFolder '00b-高级设置与实验区.png'
    if(-not (Save-GuiWindow -PidValue $styleProcess.Id -Path $advancedFrame -WindowHandle $advanced)){throw '高级设置截图失败。'}
    $styleEvidence.Screenshots+=$advancedFrame
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,241),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage([StyleCheck]::GetDlgItem($advanced,235),0xF5,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::PostMessage($advanced,0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    [StyleCheck]::SendMessage($styleWindow,0,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
    if([StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '取消高级设置错误地关闭了首页快速转换。'}
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
    if(([StyleCheck]::GetWindowLongPtr($styleFast,-16).ToInt64() -band 0x10000000) -eq 0 -or -not [StyleCheck]::IsWindowEnabled($styleFast) -or [StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 1){throw '切换方向未保留快速开关。'}
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
    if([StyleCheck]::SendMessage($styleFast,0xF0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32() -ne 0){throw '确认恢复默认后首页快速开关未关闭。'}
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
    $styleEvidence.Checks.HomeFastToggle=$true
    $styleEvidence.Checks.FastToggleHitArea=$true
    $styleEvidence.Checks.FastToggleAnimation=[StyleCheck]::ClientAnimations()
    $styleEvidence.Checks.FastHelpAccessibleName=$true
    $styleEvidence.Checks.FastHelpPosition=$true
    $styleEvidence.Checks.FastHelpClickDoesNotToggle=$true
    $styleEvidence.Checks.FastHelpTooltipOnHover=$true
    $styleEvidence.Checks.FastHelpTooltipHidesOnLeave=$true
    $styleEvidence.Checks.FastToggleSpaceKey=$true
    $styleEvidence.Checks.FastToggleFocusCue=$true
    $styleEvidence.Checks.FastFocusFramesDiffer=$true
    $styleEvidence.Checks.ExperimentalCheckboxesOnly=$true
    $styleEvidence.Status='PASSED'
    $styleEvidence.ExecutableSHA256=(Get-FileHash -LiteralPath $styleExe).Hash
    $styleEvidence.SourceSHA256=Get-TestSourceHashes
    Write-Output ('UI_STYLE_PASSED: '+$styleFolder)
}finally {
    [GuiBatchCheck]::SetThreadDpiAwarenessContext($styleDpi)|Out-Null
    if($styleProcess){if(-not $styleProcess.HasExited){foreach($title in @('转换详情','高级设置')){[StyleCheck]::PostMessage([StyleCheck]::Dialog($styleProcess.Id,$title),0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null};[StyleCheck]::PostMessage([GuiBatchCheck]::Find($styleProcess.Id),0x10,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
        if(-not $styleProcess.WaitForExit(15000)){throw '预览窗口没有退出。'}};$styleProcess.Dispose()}
    Remove-Item -LiteralPath $styleCopy -Force
    $styleEvidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $styleFolder 'style-results.json') -Encoding utf8
    $styleEvidence|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $TestOutputRoot 'ui-style-results.json') -Encoding utf8
}
