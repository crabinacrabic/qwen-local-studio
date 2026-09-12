# Скрипт фонового запуска Qwen AI Studio в режиме отдельного приложения

$qwenDir = "C:\Users\MSI\Desktop\Qwen3"
$ollamaExe = "C:\Users\MSI\AppData\Local\Programs\Ollama\ollama.exe"
$edgeExe = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# 1. Проверяем службу Ollama
$ollamaRunning = $false
try {
    $res = Invoke-RestMethod -Uri "http://127.0.0.1:11434/api/tags" -TimeoutSec 2 -ErrorAction Stop
    $ollamaRunning = $true
} catch {
    $ollamaRunning = $false
}

if (-not $ollamaRunning) {
    if (Test-Path $ollamaExe) {
        Start-Process $ollamaExe -ArgumentList "serve" -WindowStyle Hidden
        Start-Sleep -Seconds 2
    }
}

# 2. Проверяем локальный сервер Qwen Studio (порт 3000)
$serverRunning = $false
try {
    $res = Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/tags" -TimeoutSec 2 -ErrorAction Stop
    $serverRunning = $true
} catch {
    $serverRunning = $false
}

if (-not $serverRunning) {
    Start-Process "node" -ArgumentList "server.js", "--no-open" -WorkingDirectory $qwenDir -WindowStyle Hidden
    # Ожидаем старта сервера
    for ($i = 0; $i -lt 15; $i++) {
        Start-Sleep -Milliseconds 400
        try {
            $check = Invoke-RestMethod -Uri "http://127.0.0.1:3000/api/tags" -TimeoutSec 1 -ErrorAction Stop
            if ($check) { break }
        } catch {}
    }
}

# 3. Открываем интерфейс в режиме отдельного нативного окна (Edge App Mode, без Google Chrome и вкладок)
if (Test-Path $edgeExe) {
    Start-Process $edgeExe -ArgumentList "--app=http://localhost:3000", "--window-size=1300,880"
} else {
    Start-Process "http://localhost:3000"
}
