@echo off
chcp 65001 >nul

:: Ensure we are running in the directory where the .bat file is located
cd /d "%~dp0"

echo [1/2] Cleaning old obfuscated files...
del /q *.js 2>nul

echo [2/2] Obfuscating Javascript files...
call javascript-obfuscator ..\js\ --output .\

echo.
echo ===========================================
echo   BUILD SUCCESSFUL! (Torque Profit Calc)
echo ===========================================
echo You can now deploy or commit your changes.
pause
