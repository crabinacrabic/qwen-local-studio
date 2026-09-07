const http = require('http');
const fs = require('fs');
const path = require('path');

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

writeLog('INFO', '=== Запуск Qwen Local Server ===');

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
            res.writeHead(502, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Не удалось связаться с Ollama: ' + err.message }));
        });

        req.pipe(proxyReq);
        return;
    }

    // Отдача веб-страницы чата с запретом кэширования (чтобы F5 всегда отдавал свежий файл!)
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
    const { exec } = require('child_process');
    exec(`start http://localhost:${PORT}`);
});
