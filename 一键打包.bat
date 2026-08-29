@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo 正在生成免安装版 EXE，请稍候...
rem ExecutionPolicy Bypass 只对本次 PowerShell 进程有效，不会修改系统的永久执行策略。
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_exe.ps1"

if errorlevel 1 (
    echo.
    echo 打包失败，请查看上面的错误信息。
) else (
    echo.
    echo 打包成功！
    echo 程序位置：%~dp0dist\消费统计转换工具.exe
)

echo.
pause
