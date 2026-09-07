@echo off
chcp 65001 > nul
title Qwen Web Chat
node "%~dp0server.js"
pause
