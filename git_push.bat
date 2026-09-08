@echo off
chcp 65001 > nul
title Git Push to GitHub
color 0B

echo ========================================================
echo          Быстрый пуш проекта на GitHub
echo          Репозиторий: crabinacrabic/qwen-local-studio
echo ========================================================
echo.

rem Показываем текущий статус файлов
echo Измененные файлы:
git status -s
echo.

rem Спрашиваем описание коммита
set /p commit_msg="Введи описание изменений (или нажми Enter для авто-описания): "

if "%commit_msg%"=="" (
    set commit_msg=Update: local studio improvements
)

echo.
echo --------------------------------------------------------
echo 1. Добавляем файлы в индекс (git add .)...
git add .

echo 2. Создаем коммит...
git commit -m "%commit_msg%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [INFO] Новых изменений для коммита нет.
    goto end
)

echo 3. Отправляем на GitHub (git push)...
git push origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================================
    echo ✅ Успешно запушено на GitHub!
    echo Ссылка: https://github.com/crabinacrabic/qwen-local-studio
    echo ========================================================
) else (
    echo.
    echo ❌ Ошибка при отправке. Проверь подключение к интернету.
)

:end
echo.
pause
