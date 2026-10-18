# 开发截图仅使用 .NET 系统绘图和本任务窗口，不作为产品运行依赖，不发送全局键盘命令。
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class GuiBatchCheck {
    // 只按本次启动的进程 PID 定位窗口，包括隐藏测试窗口。
    public delegate bool EnumProc(IntPtr hwnd, IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder name, int capacity);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")] public static extern bool RedrawWindow(IntPtr hwnd, IntPtr rect, IntPtr region, uint flags);
    [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left, Top, Right, Bottom;}
    public static IntPtr Find(uint pid) {
        IntPtr found=IntPtr.Zero;
        EnumWindows((hwnd,p)=>{uint owner;GetWindowThreadProcessId(hwnd,out owner);var name=new StringBuilder(256);GetClassName(hwnd,name,256);if(owner==pid && name.ToString()=="DocxMathTypeBatch.B"){found=hwnd;return false;}return true;},IntPtr.Zero);
        return found;
    }
}
'@
function Save-GuiWindow {
    param([int]$PidValue,[string]$Path)
    $guiWindow = [GuiBatchCheck]::Find($PidValue)
    if ($guiWindow -eq [IntPtr]::Zero) {return $false}
    # PowerShell 默认可能被 DPI 虚拟化；取窗口物理像素并截图后恢复原线程上下文。
    $guiPreviousDpi=[GuiBatchCheck]::SetThreadDpiAwarenessContext([IntPtr](-4))
    $guiRect = [GuiBatchCheck+Rect]::new()
    if (-not [GuiBatchCheck]::GetWindowRect($guiWindow,[ref]$guiRect) -or $guiRect.Right -le $guiRect.Left -or $guiRect.Bottom -le $guiRect.Top) {
        if ($guiPreviousDpi -ne [IntPtr]::Zero) {[GuiBatchCheck]::SetThreadDpiAwarenessContext($guiPreviousDpi) | Out-Null}
        return $false
    }
    $guiBitmap = [Drawing.Bitmap]::new($guiRect.Right-$guiRect.Left,$guiRect.Bottom-$guiRect.Top)
    $guiGraphics = [Drawing.Graphics]::FromImage($guiBitmap)
    $guiWasVisible = [GuiBatchCheck]::IsWindowVisible($guiWindow)
    try {
        # 隐藏窗口不会递归绘制子控件；只短暂展示自己的窗口、不激活、不抢用户焦点。
        if (-not $guiWasVisible) {[GuiBatchCheck]::ShowWindow($guiWindow,4) | Out-Null}
        [GuiBatchCheck]::RedrawWindow($guiWindow,[IntPtr]::Zero,[IntPtr]::Zero,0x581) | Out-Null
        $guiDc=$guiGraphics.GetHdc()
        try {$guiPrinted=[GuiBatchCheck]::PrintWindow($guiWindow,$guiDc,2)} finally {$guiGraphics.ReleaseHdc($guiDc)}
        if (-not $guiPrinted) {return $false}
        # PrintWindow 返回真也可能得到空图；基本像素检查只用于拒绝无效截图，交付前仍目视检查。
        $guiVisiblePixels=0
        for ($guiX=30;$guiX -lt $guiBitmap.Width;$guiX+=80) {for ($guiY=30;$guiY -lt $guiBitmap.Height;$guiY+=80) {if ($guiBitmap.GetPixel($guiX,$guiY).GetBrightness() -gt 0.1) {$guiVisiblePixels++}}}
        if ($guiVisiblePixels -lt 5) {return $false}
        $guiBitmap.Save($Path,[Drawing.Imaging.ImageFormat]::Png)
        return $true
    } finally {
        if (-not $guiWasVisible) {[GuiBatchCheck]::ShowWindow($guiWindow,0) | Out-Null}
        $guiGraphics.Dispose();$guiBitmap.Dispose()
        if ($guiPreviousDpi -ne [IntPtr]::Zero) {[GuiBatchCheck]::SetThreadDpiAwarenessContext($guiPreviousDpi) | Out-Null}
    }
}
