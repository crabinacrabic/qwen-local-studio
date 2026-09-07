@echo off
chcp 65001 > nul
title Qwen 1.7B - Local Chat
"%LOCALAPPDATA%\Programs\Ollama\ollama.exe" run qwen3:1.7b
pause
