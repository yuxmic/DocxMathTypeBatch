#requires -Version 7.0
# 原生 C++ 构建：工具链来自系统安装，源码与生成物全部位于当前项目。
param([ValidateSet('All','Native','Scan','Convert','Gui')][string]$Target='All')
$ErrorActionPreference='Stop'
$projectRoot=$PSScriptRoot
$bin=Join-Path $projectRoot 'build/bin'
$generated=Join-Path $projectRoot 'build/generated'
$objectRoot=Join-Path $projectRoot 'build/obj'
$objects=Join-Path $objectRoot ([Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $bin,$generated,$objects -Force | Out-Null
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vsRoot=(& $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath)
if(-not $vsRoot){throw '未找到 MSVC 编译工具。'}
$toolset=[IO.File]::ReadAllText((Join-Path $vsRoot.Trim() 'VC/Auxiliary/Build/Microsoft.VCToolsVersion.default.txt')).Trim()
$tools=Join-Path $vsRoot.Trim() ('VC/Tools/MSVC/'+$toolset)
$sdk=(Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots').KitsRoot10
$sdkVersion=Get-ChildItem -LiteralPath (Join-Path $sdk 'Include') -Directory | Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'um/msopc.h')} | Sort-Object {[version]$_.Name} -Descending | Select-Object -First 1 -ExpandProperty Name
if(-not $sdkVersion){throw '未找到 Windows SDK。'}
$previousInclude,$previousLib,$previousPath=$env:INCLUDE,$env:LIB,$env:PATH
try{
    $env:INCLUDE=@((Join-Path $tools 'include'))+@('ucrt','shared','um','winrt'|ForEach-Object {Join-Path $sdk ('Include/'+$sdkVersion+'/'+$_)}) -join ';'
    $env:LIB=@((Join-Path $tools 'lib/x64'))+@('ucrt','um'|ForEach-Object {Join-Path $sdk ('Lib/'+$sdkVersion+'/'+$_+'/x64')}) -join ';'
    $env:PATH=(Join-Path $tools 'bin/Hostx64/x64')+';'+$previousPath
    $compiler=Join-Path $tools 'bin/Hostx64/x64/cl.exe'
    # 守卫 DLL 仅作为构建中间物；生成头文件后嵌入程序。
    $guardName='DmtGuard_'+[Guid]::NewGuid().ToString('N')+'.dll'
    $guardDll=Join-Path $objects $guardName
    & $compiler /nologo /TC /LD /utf-8 /O1 /MT /W4 /WX /DUNICODE /D_UNICODE ('/Fo'+(Join-Path $objects 'window_guard.obj')) ('/Fe'+$guardDll) (Join-Path $projectRoot 'src/window_guard.c') /link ('/IMPLIB:'+(Join-Path $objects 'window_guard.lib')) user32.lib kernel32.lib comctl32.lib
    if($LASTEXITCODE -ne 0){throw '窗口守卫编译失败。'}
    $bytes=[IO.File]::ReadAllBytes($guardDll)
    $header=@('// 自动生成：内嵌 Word 窗口守卫。',('inline constexpr wchar_t WordWindowGuardName[]=L"'+$guardName+'";'),'inline constexpr unsigned char WordWindowGuardBytes[]={')
    for($index=0;$index -lt $bytes.Length;$index+=32){$header+=((($bytes[$index..([Math]::Min($index+31,$bytes.Length-1))]|ForEach-Object {'0x'+$_.ToString('X2')}) -join ',')+',')}
    $header+='};'
    $header|Set-Content -LiteralPath (Join-Path $generated 'window_guard_bytes.h') -Encoding utf8
    $resource=Join-Path $objects 'app.res'
    if($Target -in @('All','Gui')){
        & (Join-Path $sdk ('bin/'+$sdkVersion+'/x64/rc.exe')) /nologo /c65001 /i (Join-Path $projectRoot 'assets') /fo $resource (Join-Path $projectRoot 'assets/app.rc')
        if($LASTEXITCODE -ne 0){throw '图标资源编译失败。'}
    }
    $programs=@(
        @{Target='Native';Source=@('tests/native/native_api_probe.cpp');Name='native-api-probe';Info='build-info.json'},
        @{Target='Scan';Source=@('src/scan_docx.cpp');Name='docx-math-scan';Info='scan-build-info.json'},
        @{Target='Convert';Source=@('src/convert_docx.cpp');Name='docx-math-convert';Info='convert-build-info.json'},
        @{Target='Gui';Source=@('src/gui.cpp','src/convert_docx.cpp');Name='DocxMathTypeBatch';Info='gui-build-info.json'}
    )
    foreach($program in ($programs|Where-Object {$Target -eq 'All' -or $_.Target -eq $Target})){
        $exe=Join-Path $bin ($program.Name+'.exe')
        $extra=@();$links=@();$sources=@($program.Source|ForEach-Object {Join-Path $projectRoot $_})
        if($program.Target -eq 'Gui'){
            $extra=@('/DDOCX_MATH_GUI');$sources+=$resource
            $links=@('/SUBSYSTEM:WINDOWS','comctl32.lib','comdlg32.lib','shell32.lib','gdi32.lib','gdiplus.lib','d2d1.lib','dwrite.lib','advapi32.lib','uxtheme.lib','uuid.lib')
        }
        & $compiler /nologo /EHsc /std:c++17 /utf-8 /MT /W4 /WX /O1 /DUNICODE /D_UNICODE /I (Join-Path $projectRoot 'src') /I $generated /I (Join-Path $projectRoot 'tests/native') @extra ('/Fe'+$exe) ('/Fo'+$objects+'\') @sources /link ole32.lib oleaut32.lib user32.lib advapi32.lib xmllite.lib crypt32.lib @links
        if($LASTEXITCODE -ne 0){throw ('编译失败：'+$program.Name)}
        $imports=& (Join-Path $tools 'bin/Hostx64/x64/dumpbin.exe') /nologo /dependents $exe
        if($LASTEXITCODE -ne 0){throw 'PE 依赖读取失败。'}
        $dlls=@([regex]::Matches(($imports -join "`n"),'(?im)^\s+([a-z0-9_.-]+\.dll)\s*$')|ForEach-Object {$_.Groups[1].Value})
        if(-not $dlls -or ($dlls|Where-Object {$_ -match '^(mscoree|coreclr|clr|vcruntime.*|msvcp.*|concrt.*|ucrtbase|api-ms-win-crt.*)\.dll$'})){throw '成品包含额外运行库依赖。'}
        $peBytes=[IO.File]::ReadAllBytes($exe);$pe=[BitConverter]::ToInt32($peBytes,0x3c);$optional=$pe+24;$clr=$optional+112+14*8
        if([BitConverter]::ToUInt16($peBytes,$pe+4) -ne 0x8664 -or [BitConverter]::ToUInt16($peBytes,$optional) -ne 0x20b -or [BitConverter]::ToUInt32($peBytes,$clr) -ne 0 -or [BitConverter]::ToUInt32($peBytes,$clr+4) -ne 0){throw '成品必须为原生 x64 PE。'}
        [ordered]@{Executable=('build/bin/'+$program.Name+'.exe');ExecutableSHA256=(Get-FileHash -LiteralPath $exe).Hash;ExeBytes=$peBytes.Length;Machine='x64';ImportedDlls=$dlls;ClrDirectory=0}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $bin $program.Info) -Encoding utf8
        Write-Output ('BUILD_PASSED: build/bin/'+$program.Name+'.exe')
    }
}finally{
    $env:INCLUDE,$env:LIB,$env:PATH=$previousInclude,$previousLib,$previousPath
    # 仅删除本轮创建的编译临时目录，先核对其绝对路径边界。
    $resolved=[IO.Path]::GetFullPath($objects);$allowedRoot=[IO.Path]::GetFullPath($objectRoot)
    if(-not $resolved.StartsWith($allowedRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw '编译临时目录越界。'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
