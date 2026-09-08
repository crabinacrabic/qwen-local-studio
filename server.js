const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec, spawn } = require('child_process');

const PORT = 3000;
const OLLAMA_PORT = 11434;
const LOG_FILE = path.join(__dirname, 'app.log');

// Функция записи в лог-файл и консоль с локальным временем (не UTC!)
function getLocalTimestamp() {
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function writeLog(level, message) {
    const timestamp = getLocalTimestamp();
    const logLine = `[${timestamp}] [${level}] ${message}\n`;
    process.stdout.write(logLine);
    try {
        fs.appendFileSync(LOG_FILE, logLine, 'utf8');
    } catch (e) {
        console.error('Ошибка записи лога:', e);
    }
}

// Автоматический запуск Ollama, если она не запущена
function ensureOllamaRunning() {
    const ollamaPath = path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Ollama', 'ollama.exe');
    
    const checkReq = http.get({ hostname: '127.0.0.1', port: OLLAMA_PORT, path: '/api/tags' }, (res) => {
        writeLog('INFO', 'Служба Ollama активна и отвечает.');
    });

    checkReq.on('error', () => {
        writeLog('WARN', 'Служба Ollama не была активна после включения ПК. Запускаем автоматически...');
        try {
            const child = spawn(ollamaPath, ['serve'], {
                detached: true,
                stdio: 'ignore'
            });
            child.unref();
            writeLog('INFO', 'Ollama успешно запущена в фоне.');
        } catch (err) {
            writeLog('ERROR', 'Не удалось автоматически запустить Ollama: ' + err.message);
        }
    });
}

writeLog('INFO', '=== Запуск Qwen Local Server ===');
ensureOllamaRunning();

const server = http.createServer((req, res) => {
    // Логирование событий от фронтенда
    if (req.url === '/api/client-log' && req.method === 'POST') {
        let body = '';
        req.on('data', chunk => body += chunk);
        req.on('end', () => {
            try {
                const data = JSON.parse(body);
                writeLog(data.level || 'INFO', `[UI] ${data.message}`);
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end('{"status":"ok"}');
            } catch (err) {
                res.writeHead(400);
                res.end('bad json');
            }
        });
        return;
    }

    // Проксирование запросов к Ollama
    if (req.url.startsWith('/api/')) {
        writeLog('DEBUG', `Proxy ${req.method} ${req.url}`);

        const proxyReq = http.request({
            hostname: '127.0.0.1',
            port: OLLAMA_PORT,
            path: req.url,
            method: req.method,
            headers: req.headers
        }, (proxyRes) => {
            res.writeHead(proxyRes.statusCode, proxyRes.headers);
            proxyRes.pipe(res);
        });

        proxyReq.on('error', (err) => {
            writeLog('ERROR', `Ошибка связи с Ollama (${req.url}): ${err.message}`);
            // Если соединение отклонено, пробуем еще раз пнуть запуск Ollama
            if (err.code === 'ECONNREFUSED') {
                ensureOllamaRunning();
            }
            res.writeHead(502, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Служба Ollama еще запускается. Пожалуйста, обновите страницу через 3 секунды.' }));
        });

        req.pipe(proxyReq);
        return;
    }

    // Отдача веб-страницы чата с запретом кэширования
    const filePath = path.join(__dirname, 'chat.html');
    fs.readFile(filePath, (err, data) => {
        if (err) {
            writeLog('ERROR', `Не удалось прочитать chat.html: ${err.message}`);
            res.writeHead(500);
            res.end('Error loading chat file');
            return;
        }
        res.writeHead(200, {
            'Content-Type': 'text/html; charset=utf-8',
            'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
        });
        res.end(data);
    });
});

server.listen(PORT, '127.0.0.1', () => {
    writeLog('INFO', `Сервер чата запущен и слушает http://localhost:${PORT}`);
    writeLog('INFO', `Лог-файл пишется в ${LOG_FILE}`);
    exec(`start http://localhost:${PORT}`);
});
