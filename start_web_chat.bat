@echo off
chcp 65001 > nul
title Qwen Web Chat

rem Проверяем запущена ли Ollama, если нет - запускаем в фоне
tasklist /FI "IMAGENAME eq ollama.exe" 2>nul | find /I /N "ollama.exe" >nul
if "%ERRORLEVEL%"=="1" (
    echo [INFO] Фоновая служба Ollama не была активна. Запускаем...
    start "" /B "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" serve
    timeout /t 2 /nobreak >nul
)

node "%~dp0server.js"
pause
