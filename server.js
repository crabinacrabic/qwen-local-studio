const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec, spawn, execSync } = require('child_process');

const PORT = 3000;
const OLLAMA_PORT = 11434;
const LOG_FILE = path.join(__dirname, 'app.log');
const KNOWLEDGE_DIR = path.join(__dirname, 'knowledge');
const INDEX_FILE = path.join(__dirname, 'knowledge_index.json');

// Обеспечиваем наличие папки для документов базы знаний
if (!fs.existsSync(KNOWLEDGE_DIR)) {
    try {
        fs.mkdirSync(KNOWLEDGE_DIR, { recursive: true });
    } catch (e) {
        console.error('Ошибка создания папки knowledge:', e);
    }
}

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

// ==========================================
// RAG: Векторные эмбеддинги и семантический поиск
// ==========================================

// Загрузка индекса из JSON
function loadKnowledgeIndex() {
    try {
        if (fs.existsSync(INDEX_FILE)) {
            const raw = fs.readFileSync(INDEX_FILE, 'utf8');
            return JSON.parse(raw);
        }
    } catch (e) {
        writeLog('ERROR', 'Ошибка чтения knowledge_index.json: ' + e.message);
    }
    return [];
}

// Сохранение индекса в JSON
function saveKnowledgeIndex(chunks) {
    try {
        fs.writeFileSync(INDEX_FILE, JSON.stringify(chunks, null, 2), 'utf8');
        return true;
    } catch (e) {
        writeLog('ERROR', 'Ошибка записи knowledge_index.json: ' + e.message);
        return false;
    }
}

// Получение векторного эмбеддинга от локальной Ollama (nomic-embed-text)
function getEmbedding(text) {
    return new Promise((resolve, reject) => {
        const payload = JSON.stringify({
            model: 'nomic-embed-text',
            prompt: text
        });

        const req = http.request({
            hostname: '127.0.0.1',
            port: OLLAMA_PORT,
            path: '/api/embeddings',
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(payload)
            }
        }, (res) => {
            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                try {
                    const json = JSON.parse(data);
                    if (json.embedding && Array.isArray(json.embedding)) {
                        resolve(json.embedding);
                    } else {
                        reject(new Error(json.error || 'Ollama не вернула вектор эмбеддинга'));
                    }
                } catch (e) {
                    reject(new Error('Сбой парсинга ответа эмбеддинга: ' + e.message));
                }
            });
        });

        req.on('error', (err) => {
            reject(new Error('Ошибка соединения с Ollama: ' + err.message));
        });

        req.setTimeout(30000, () => {
            req.destroy();
            reject(new Error('Таймаут генерации эмбеддинга'));
        });

        req.write(payload);
        req.end();
    });
}

// Евклидова длина (норма) вектора
function vectorNorm(vec) {
    let sum = 0;
    for (let i = 0; i < vec.length; i++) {
        sum += vec[i] * vec[i];
    }
    return Math.sqrt(sum);
}

// Косинусное сходство двух векторов
function cosineSimilarity(vecA, normA, vecB, normB) {
    if (!vecA || !vecB || vecA.length !== vecB.length) return 0;
    let dot = 0;
    for (let i = 0; i < vecA.length; i++) {
        dot += vecA[i] * vecB[i];
    }
    const denom = normA * normB;
    return denom === 0 ? 0 : dot / denom;
}

// Извлечение текста из Word-файлов (.docx) через системную утилиту tar.exe
function extractTextFromDocx(filePath) {
    try {
        const xml = execSync(`tar.exe -xf "${filePath}" -O word/document.xml`, {
            maxBuffer: 50 * 1024 * 1024,
            encoding: 'utf8',
            stdio: ['ignore', 'pipe', 'ignore']
        });

        const text = xml
            .replace(/<\/w:p>/g, '\n')
            .replace(/<w:tab\/>/g, '\t')
            .replace(/<w:br\/>/g, '\n')
            .replace(/<[^>]+>/g, '')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/[ \t]+/g, ' ')
            .replace(/\n\s*\n+/g, '\n\n')
            .trim();

        return text;
    } catch (err) {
        writeLog('ERROR', `Ошибка извлечения текста из docx (${filePath}): ${err.message}`);
        throw new Error('Не удалось распаковать docx файл: ' + err.message);
    }
}

// Интеллектуальное разбиение текста на фрагменты (чанки) с жестким ограничением длины
function chunkText(text, maxChars = 500, overlap = 70) {
    if (!text || text.trim().length === 0) return [];

    const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
    const paragraphs = normalized.split(/\n+/).map(p => p.trim()).filter(Boolean);
    const chunks = [];
    let currentChunk = '';

    for (const para of paragraphs) {
        if ((currentChunk + ' ' + para).length > maxChars) {
            if (currentChunk.trim().length > 0) {
                chunks.push(currentChunk.trim());
                const words = currentChunk.split(/\s+/);
                let overlapText = '';
                for (let i = words.length - 1; i >= 0; i--) {
                    if ((words[i] + ' ' + overlapText).length <= overlap) {
                        overlapText = (words[i] + ' ' + overlapText).trim();
                    } else break;
                }
                currentChunk = overlapText;
            }

            if (para.length > maxChars) {
                const words = para.split(/\s+/).filter(Boolean);
                for (const word of words) {
                    if ((currentChunk + ' ' + word).length > maxChars) {
                        if (currentChunk.trim().length > 0) {
                            chunks.push(currentChunk.trim());
                            currentChunk = '';
                        }
                        if (word.length > maxChars) {
                            let w = word;
                            while (w.length > maxChars) {
                                chunks.push(w.slice(0, maxChars));
                                w = w.slice(maxChars);
                            }
                            currentChunk = w;
                        } else {
                            currentChunk = word;
                        }
                    } else {
                        currentChunk = currentChunk ? (currentChunk + ' ' + word) : word;
                    }
                }
            } else {
                currentChunk = currentChunk ? (currentChunk + ' ' + para) : para;
            }
        } else {
            currentChunk = currentChunk ? (currentChunk + ' ' + para) : para;
        }
    }

    if (currentChunk && currentChunk.trim().length > 0) {
        chunks.push(currentChunk.trim());
    }

    return chunks.filter(c => c.length >= 10);
}

// Парсер JSON тела запроса
function readJsonBody(req, cb) {
    let body = '';
    req.on('data', chunk => {
        body += chunk;
        if (body.length > 50 * 1024 * 1024) {
            req.destroy();
            cb(new Error('Размер данных превышает лимит 50 МБ'));
        }
    });
    req.on('end', () => {
        try {
            const data = JSON.parse(body);
            cb(null, data);
        } catch (e) {
            cb(e, null);
        }
    });
    req.on('error', cb);
}

writeLog('INFO', '=== Запуск Qwen Local Server ===');
ensureOllamaRunning();

const server = http.createServer((req, res) => {
    // 1. Логирование событий от фронтенда
    if (req.url === '/api/client-log' && req.method === 'POST') {
        readJsonBody(req, (err, data) => {
            if (err || !data) {
                res.writeHead(400);
                res.end('bad json');
                return;
            }
            writeLog(data.level || 'INFO', `[UI] ${data.message}`);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end('{"status":"ok"}');
        });
        return;
    }

    // 2. RAG: Получение списка документов в базе
    if (req.url === '/api/rag/documents' && req.method === 'GET') {
        const index = loadKnowledgeIndex();
        const docMap = new Map();
        for (const item of index) {
            if (!docMap.has(item.filename)) {
                docMap.set(item.filename, {
                    filename: item.filename,
                    chunksCount: 0,
                    totalChars: 0,
                    createdAt: item.createdAt || 'Неизвестно'
                });
            }
            const doc = docMap.get(item.filename);
            doc.chunksCount++;
            doc.totalChars += (item.text ? item.text.length : 0);
        }
        const documents = Array.from(docMap.values());
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ status: 'ok', documents, totalChunks: index.length }));
        return;
    }

    // 3. RAG: Индексация нового документа (разбивка + эмбеддинги)
    if (req.url === '/api/rag/index' && req.method === 'POST') {
        readJsonBody(req, async (err, body) => {
            if (err || !body || !body.filename) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется имя файла' }));
                return;
            }

            const filename = path.basename(body.filename).trim();
            let text = body.text ? String(body.text).trim() : '';

            // Поддержка Base64 (например, загрузка файлов .docx, бинарных файлов)
            if (body.base64) {
                try {
                    const commaIdx = body.base64.indexOf(',');
                    const pureB64 = commaIdx >= 0 ? body.base64.slice(commaIdx + 1) : body.base64;
                    const buffer = Buffer.from(pureB64, 'base64');

                    const tempDocxPath = path.join(KNOWLEDGE_DIR, `temp_${Date.now()}_${filename}`);
                    fs.writeFileSync(tempDocxPath, buffer);

                    if (filename.toLowerCase().endsWith('.docx')) {
                        writeLog('INFO', `[RAG] Распаковка текста из Word-документа (.docx): "${filename}"...`);
                        text = extractTextFromDocx(tempDocxPath);
                        writeLog('INFO', `[RAG] Извлечено ${text.length} символов текста из "${filename}".`);
                    } else {
                        text = buffer.toString('utf8');
                    }

                    try { fs.unlinkSync(tempDocxPath); } catch (e) {}
                } catch (b64Err) {
                    writeLog('ERROR', `[RAG] Ошибка обработки файла "${filename}": ${b64Err.message}`);
                    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                    res.end(JSON.stringify({ error: 'Ошибка обработки файла: ' + b64Err.message }));
                    return;
                }
            } else if (filename.toLowerCase().endsWith('.docx') && text.startsWith('PK')) {
                // Если файл .docx был передан как бинарная строка
                try {
                    const tempDocxPath = path.join(KNOWLEDGE_DIR, `temp_${Date.now()}_${filename}`);
                    fs.writeFileSync(tempDocxPath, text, 'binary');
                    text = extractTextFromDocx(tempDocxPath);
                    try { fs.unlinkSync(tempDocxPath); } catch (e) {}
                } catch (pkErr) {
                    writeLog('ERROR', `[RAG] Ошибка распаковки docx "${filename}": ${pkErr.message}`);
                }
            }

            if (!filename || !text) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Пустой документ или не удалось извлечь текст' }));
                return;
            }

            writeLog('INFO', `[RAG] Начало индексации документа: "${filename}" (${text.length} символов)...`);

            // Сохраняем текстовую копию в папку knowledge
            try {
                fs.writeFileSync(path.join(KNOWLEDGE_DIR, filename + '.txt'), text, 'utf8');
            } catch (e) {
                writeLog('WARN', `[RAG] Не удалось сохранить копию в knowledge: ${e.message}`);
            }

            const chunks = chunkText(text, 500, 70);
            if (chunks.length === 0) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Текст слишком короткий для индексации' }));
                return;
            }

            writeLog('INFO', `[RAG] Документ "${filename}" разбит на ${chunks.length} фрагментов. Вычисление векторов через nomic-embed-text...`);

            try {
                const newChunks = [];
                for (let i = 0; i < chunks.length; i++) {
                    const chunkStr = chunks[i];
                    const embedding = await getEmbedding(chunkStr);
                    const norm = vectorNorm(embedding);
                    newChunks.push({
                        id: `${filename}_chunk_${i}_${Date.now()}`,
                        filename: filename,
                        chunkIndex: i,
                        text: chunkStr,
                        embedding: embedding,
                        norm: norm,
                        createdAt: getLocalTimestamp()
                    });
                }

                // Заменяем старые чанки с таким же именем файла
                const currentIndex = loadKnowledgeIndex();
                const filteredIndex = currentIndex.filter(c => c.filename !== filename);
                const updatedIndex = [...filteredIndex, ...newChunks];
                saveKnowledgeIndex(updatedIndex);

                writeLog('INFO', `[RAG] Документ "${filename}" успешно добавлен в базу (${chunks.length} чанков). Всего чанков в базе: ${updatedIndex.length}`);

                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({
                    status: 'ok',
                    filename: filename,
                    chunksIndexed: chunks.length,
                    totalChunksInDb: updatedIndex.length
                }));
            } catch (embErr) {
                writeLog('ERROR', `[RAG] Сбой векторизации для "${filename}": ${embErr.message}`);
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Ошибка векторизации: ' + embErr.message }));
            }
        });
        return;
    }

    // 4. RAG: Векторный семантический поиск по базе знаний
    if (req.url === '/api/rag/search' && req.method === 'POST') {
        readJsonBody(req, async (err, body) => {
            if (err || !body || !body.query) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется поле query' }));
                return;
            }

            const query = String(body.query).trim();
            const topK = parseInt(body.topK) || 3;
            const minScore = typeof body.minScore === 'number' ? body.minScore : 0.35;

            const index = loadKnowledgeIndex();
            if (index.length === 0) {
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', results: [], totalIndexed: 0 }));
                return;
            }

            try {
                const queryVec = await getEmbedding(query);
                const queryNorm = vectorNorm(queryVec);

                const scored = [];
                for (const item of index) {
                    const itemNorm = item.norm || vectorNorm(item.embedding);
                    const score = cosineSimilarity(queryVec, queryNorm, item.embedding, itemNorm);
                    if (score >= minScore) {
                        scored.push({
                            filename: item.filename,
                            chunkIndex: item.chunkIndex,
                            text: item.text,
                            score: score
                        });
                    }
                }

                scored.sort((a, b) => b.score - a.score);
                const topResults = scored.slice(0, topK).map(r => ({
                    filename: r.filename,
                    chunkIndex: r.chunkIndex,
                    text: r.text,
                    score: Math.round(r.score * 100) / 100
                }));

                writeLog('INFO', `[RAG] Поиск "${query.substring(0, 40)}...": найдено ${topResults.length} совпадений (макс: ${topResults[0] ? topResults[0].score : 'нет'})`);

                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({
                    status: 'ok',
                    results: topResults,
                    totalIndexed: index.length
                }));
            } catch (searchErr) {
                writeLog('ERROR', `[RAG] Ошибка поиска: ${searchErr.message}`);
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: searchErr.message }));
            }
        });
        return;
    }

    // 5. RAG: Удаление документа из базы знаний
    if (req.url === '/api/rag/delete' && req.method === 'POST') {
        readJsonBody(req, (err, body) => {
            if (err || !body || !body.filename) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется поле filename' }));
                return;
            }

            const filename = body.filename;
            const index = loadKnowledgeIndex();
            const initialCount = index.length;
            const newIndex = index.filter(c => c.filename !== filename);
            saveKnowledgeIndex(newIndex);

            const filePath = path.join(KNOWLEDGE_DIR, filename);
            if (fs.existsSync(filePath)) {
                try { fs.unlinkSync(filePath); } catch (e) {}
            }

            writeLog('INFO', `[RAG] Документ "${filename}" удален (снято ${initialCount - newIndex.length} чанков).`);

            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({
                status: 'ok',
                deleted: filename,
                removedChunks: initialCount - newIndex.length,
                remainingChunks: newIndex.length
            }));
        });
        return;
    }

    // 6. RAG: Полная очистка базы знаний
    if (req.url === '/api/rag/clear' && req.method === 'POST') {
        saveKnowledgeIndex([]);
        try {
            const files = fs.readdirSync(KNOWLEDGE_DIR);
            for (const file of files) {
                try { fs.unlinkSync(path.join(KNOWLEDGE_DIR, file)); } catch (e) {}
            }
        } catch (e) {}
        writeLog('WARN', '[RAG] Вся база знаний очищена пользователем.');
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ status: 'ok', message: 'База знаний полностью очищена' }));
        return;
    }

    // 7. Проксирование остальных запросов к Ollama (/api/chat, /api/tags, /api/pull, /api/delete и т.д.)
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
            if (err.code === 'ECONNREFUSED') {
                ensureOllamaRunning();
            }
            res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' });
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
