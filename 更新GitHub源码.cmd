@echo off
chcp 65001 >nul
rem 使用 PowerShell 7 执行更新，保留窗口显示执行结果。
pwsh -NoProfile -File "%~dp0update-github.ps1" %*
set "update_exit_code=%errorlevel%"
pause
exit /b %update_exit_code%
