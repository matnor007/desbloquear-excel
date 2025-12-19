@echo off
title Desbloqueador de Planilha Excel

if "%~1"=="" (
    echo Arraste um arquivo .xlsx ou .xlsm sobre este arquivo.
    pause
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass ^
-File "%~dp0desbloquear.ps1" ^
-Arquivo "%~1"

echo.
echo Processo finalizado.
pause
