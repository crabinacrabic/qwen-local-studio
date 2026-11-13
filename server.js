const http = require('http');
const https = require('https');
const { URL } = require('url');
const fs = require('fs');
const path = require('path');
const { exec, spawn, execSync, execFile } = require('child_process');
const db = require('./db.js');

const PORT = 3000;
const OLLAMA_PORT = 11434;
const LOG_FILE = path.join(__dirname, 'app.log');
const KNOWLEDGE_DIR = path.join(__dirname, 'knowledge');
const MODELS_DIR = path.join(__dirname, 'models');
const INDEX_FILE = path.join(__dirname, 'knowledge_index.json');

function ensureDirExists(dirPath, label) {
    if (!fs.existsSync(dirPath)) {
        try {
            fs.mkdirSync(dirPath, { recursive: true });
        } catch (e) {
            console.error(`Ошибка создания папки ${label}:`, e);
        }
    }
}

// Обеспечиваем наличие папок для документов базы знаний и локальных GGUF-моделей
ensureDirExists(KNOWLEDGE_DIR, 'knowledge');
ensureDirExists(MODELS_DIR, 'models');

function getOllamaExecutablePath() {
    const winDefault = path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Ollama', 'ollama.exe');
    return (process.platform === 'win32' && fs.existsSync(winDefault)) ? winDefault : 'ollama';
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
    const ollamaPath = getOllamaExecutablePath();

    const checkReq = http.get({ hostname: '127.0.0.1', port: OLLAMA_PORT, path: '/api/tags' }, () => {
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

// ==========================================
// Web Search & Fetch (Выход в интернет)
// ==========================================

function unescapeHtml(html) {
    if (!html) return '';
    return html
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&#x27;/g, "'")
        .replace(/&nbsp;/g, ' ')
        .replace(/&#(\d+);/g, (match, dec) => {
            try { return String.fromCharCode(dec); } catch(e) { return ''; }
        });
}

// Загрузка и очистка содержимого веб-страницы по URL
function fetchPageContent(targetUrl, maxChars = 2000, maxRedirects = 3) {
    return new Promise((resolve) => {
        if (maxRedirects <= 0) return resolve('');
        try {
            const parsedUrl = new URL(targetUrl);
            if (!['http:', 'https:'].includes(parsedUrl.protocol)) {
                return resolve('');
            }
            const pathname = parsedUrl.pathname.toLowerCase();
            if (pathname.match(/\.(pdf|zip|rar|7z|exe|dmg|iso|mp3|mp4|avi|mkv|jpg|jpeg|png|gif|webp)$/)) {
                return resolve('');
            }

            const client = parsedUrl.protocol === 'https:' ? https : http;
            const options = {
                hostname: parsedUrl.hostname,
                port: parsedUrl.port || (parsedUrl.protocol === 'https:' ? 443 : 80),
                path: parsedUrl.pathname + parsedUrl.search,
                method: 'GET',
                headers: {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
                    'Accept': 'text/html,application/xhtml+xml,text/plain;q=0.9,*/*;q=0.8',
                    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7'
                },
                timeout: 7000
            };

            const req = client.get(options, (res) => {
                if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
                    try {
                        const redirectUrl = new URL(res.headers.location, targetUrl).href;
                        return fetchPageContent(redirectUrl, maxChars, maxRedirects - 1).then(resolve);
                    } catch (e) {
                        return resolve('');
                    }
                }

                if (res.statusCode !== 200) {
                    return resolve('');
                }

                let raw = '';
                res.on('data', chunk => {
                    raw += chunk;
                    if (raw.length > 500 * 1024) {
                        req.destroy();
                    }
                });

                res.on('end', () => {
                    try {
                        let text = raw
                            .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, ' ')
                            .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, ' ')
                            .replace(/<header\b[^<]*(?:(?!<\/header>)<[^<]*)*<\/header>/gi, ' ')
                            .replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, ' ')
                            .replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, ' ')
                            .replace(/<aside\b[^<]*(?:(?!<\/aside>)<[^<]*)*<\/aside>/gi, ' ')
                            .replace(/<svg\b[^<]*(?:(?!<\/svg>)<[^<]*)*<\/svg>/gi, ' ')
                            .replace(/<noscript\b[^<]*(?:(?!<\/noscript>)<[^<]*)*<\/noscript>/gi, ' ')
                            .replace(/<[^>]+>/g, ' ')
                            .replace(/\s+/g, ' ');
                        text = unescapeHtml(text).trim();
                        if (text.length > maxChars) {
                            text = text.substring(0, maxChars) + '...';
                        }
                        resolve(text);
                    } catch (e) {
                        resolve('');
                    }
                });
            });

            req.on('error', () => resolve(''));
            req.on('timeout', () => {
                req.destroy();
                resolve('');
            });
        } catch (e) {
            resolve('');
        }
    });
}

// Поиск через DuckDuckGo HTML
function searchDuckDuckGo(query, maxResults = 5) {
    return new Promise((resolve, reject) => {
        const postData = 'q=' + encodeURIComponent(query);
        const options = {
            hostname: 'html.duckduckgo.com',
            port: 443,
            path: '/html/',
            method: 'POST',
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
                'Content-Type': 'application/x-www-form-urlencoded',
                'Content-Length': Buffer.byteLength(postData)
            },
            timeout: 10000
        };

        const req = https.request(options, (res) => {
            let html = '';
            res.on('data', chunk => html += chunk);
            res.on('end', () => {
                try {
                    const results = [];
                    const resultBlocks = html.split(/class="[^"]*result\s+results_links[^"]*"/);
                    for (let i = 1; i < resultBlocks.length && results.length < maxResults; i++) {
                        const block = resultBlocks[i];
                        
                        const urlMatch = block.match(/<a[^>]+class="result__url"[^>]+href="([^"]+)"[^>]*>/) ||
                                         block.match(/<a[^>]+class="result__snippet"[^>]+href="([^"]+)"/) ||
                                         block.match(/<h2[^>]*>[\s\S]*?<a[^>]+href="([^"]+)"[^>]*>/);
                        
                        let url = '';
                        if (urlMatch) {
                            url = urlMatch[1];
                            if (url.includes('uddg=')) {
                                const uddgMatch = url.match(/uddg=([^&]+)/);
                                if (uddgMatch) {
                                    try { url = decodeURIComponent(uddgMatch[1]); } catch(e){}
                                }
                            }
                        }

                        let title = '';
                        const h2Match = block.match(/<h2[^>]*class="result__title"[^>]*>([\s\S]*?)<\/h2>/);
                        if (h2Match) {
                            title = unescapeHtml(h2Match[1].replace(/<[^>]+>/g, '')).trim();
                        }

                        let snippet = '';
                        const snippetMatch = block.match(/class="result__snippet"[^>]*>([\s\S]*?)<\/a>/);
                        if (snippetMatch) {
                            snippet = unescapeHtml(snippetMatch[1].replace(/<[^>]+>/g, '')).trim();
                        }

                        if (title && url) {
                            results.push({ title, url, snippet });
                        }
                    }
                    resolve(results);
                } catch (parseErr) {
                    reject(new Error('Ошибка парсинга выдачи DDG: ' + parseErr.message));
                }
            });
        });

        req.on('error', (err) => reject(new Error('Сетевая ошибка DDG: ' + err.message)));
        req.on('timeout', () => {
            req.destroy();
            reject(new Error('Таймаут поиска DuckDuckGo'));
        });

        req.write(postData);
        req.end();
    });
}

// Резервный поиск через DuckDuckGo Lite
function searchDuckDuckGoLite(query, maxResults = 5) {
    return new Promise((resolve) => {
        const postData = 'q=' + encodeURIComponent(query);
        const req = https.request({
            hostname: 'lite.duckduckgo.com',
            port: 443,
            path: '/lite/',
            method: 'POST',
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
                'Content-Type': 'application/x-www-form-urlencoded',
                'Content-Length': Buffer.byteLength(postData)
            },
            timeout: 8000
        }, (res) => {
            let html = '';
            res.on('data', chunk => html += chunk);
            res.on('end', () => {
                try {
                    const results = [];
                    const linkMatches = [...html.matchAll(/<a[^>]+class=['"]result-link['"][^>]+href=['"]([^'"]+)['"][^>]*>([\s\S]*?)<\/a>/g)];
                    const snippetMatches = [...html.matchAll(/<td[^>]+class=['"]result-snippet['"][^>]*>([\s\S]*?)<\/td>/g)];

                    for (let i = 0; i < linkMatches.length && results.length < maxResults; i++) {
                        let url = linkMatches[i][1];
                        if (url.includes('uddg=')) {
                            const uddgMatch = url.match(/uddg=([^&]+)/);
                            if (uddgMatch) {
                                try { url = decodeURIComponent(uddgMatch[1]); } catch(e){}
                            }
                        }
                        const title = unescapeHtml(linkMatches[i][2].replace(/<[^>]+>/g, '')).trim();
                        const snippet = snippetMatches[i] ? unescapeHtml(snippetMatches[i][1].replace(/<[^>]+>/g, '')).trim() : '';
                        if (title && url) {
                            results.push({ title, url, snippet });
                        }
                    }
                    resolve(results);
                } catch (e) {
                    resolve([]);
                }
            });
        });
        req.on('error', () => resolve([]));
        req.on('timeout', () => {
            req.destroy();
            resolve([]);
        });
        req.write(postData);
        req.end();
    });
}

// Комплексный поиск с автопереключением и глубоким чтением страниц
async function performWebSearch(query, maxResults = 5, deepFetch = true) {
    let results = [];
    try {
        results = await searchDuckDuckGo(query, maxResults);
    } catch (e) {
        writeLog('WARN', `[WebSearch] Ошибка DDG HTML: ${e.message}`);
    }

    if (!results || results.length === 0) {
        writeLog('INFO', `[WebSearch] Попытка через DDG Lite fallback для "${query}"...`);
        try {
            results = await searchDuckDuckGoLite(query, maxResults);
        } catch (e) {
            writeLog('WARN', `[WebSearch] Ошибка DDG Lite: ${e.message}`);
        }
    }

    if (deepFetch && results && results.length > 0) {
        const topToFetch = results.slice(0, 2);
        for (const item of topToFetch) {
            try {
                const pageText = await fetchPageContent(item.url, 1600);
                if (pageText && pageText.length > 80) {
                    item.pageContent = pageText;
                }
            } catch (fetchErr) {
                // Игнорируем ошибку чтения одной страницы
            }
        }
    }

    return results || [];
}

// ==========================================
// Модуль сканирования железа ПК (Hardware Advisor)
// ==========================================

let hardwareSpecsCache = {
    timestamp: 0,
    data: null
};

function getSystemHardwareSpecs() {
    return new Promise((resolve) => {
        const now = Date.now();
        if (hardwareSpecsCache.data && (now - hardwareSpecsCache.timestamp < 30 * 1000)) {
            return resolve(hardwareSpecsCache.data);
        }

        const psCmd = `
$vramBytes = 0
$reg = Get-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}\\0*' -ErrorAction SilentlyContinue | Where-Object { $_.'HardwareInformation.qwMemorySize' } | Select-Object -First 1
if ($reg) { $vramBytes = $reg.'HardwareInformation.qwMemorySize' }

$gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1 Name, AdapterRAM
if ($vramBytes -eq 0 -and $gpu.AdapterRAM) { $vramBytes = $gpu.AdapterRAM }

$os = Get-CimInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize, FreePhysicalMemory
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Size, FreeSpace
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1 Name

[PSCustomObject]@{
    gpuName = if ($gpu.Name) { $gpu.Name.Trim() } else { 'GPU' }
    vramBytes = [int64]$vramBytes
    vramGB = [math]::Round($vramBytes / 1GB, 1)
    ramTotalGB = [math]::Round($os.TotalVisibleMemorySize * 1KB / 1GB, 1)
    ramFreeGB = [math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 1)
    diskTotalGB = [math]::Round($disk.Size / 1GB, 1)
    diskFreeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
    cpuName = if ($cpu.Name) { $cpu.Name.Trim() } else { 'CPU' }
} | ConvertTo-Json -Compress
        `.trim();

        execFile('powershell', ['-NoProfile', '-Command', psCmd], { timeout: 6000 }, (err, stdout) => {
            if (err || !stdout) {
                const fallback = {
                    gpuName: 'AMD Radeon RX 6600',
                    vramGB: 8.0,
                    ramTotalGB: 16.0,
                    ramFreeGB: 9.0,
                    diskTotalGB: 475.0,
                    diskFreeGB: 290.0,
                    cpuName: 'AMD Ryzen 7 5700X',
                    fallback: true
                };
                return resolve(fallback);
            }
            try {
                const specs = JSON.parse(stdout.trim());
                hardwareSpecsCache = {
                    timestamp: now,
                    data: specs
                };
                resolve(specs);
            } catch (parseErr) {
                resolve({
                    gpuName: 'AMD Radeon RX 6600',
                    vramGB: 8.0,
                    ramTotalGB: 16.0,
                    ramFreeGB: 9.0,
                    diskTotalGB: 475.0,
                    diskFreeGB: 290.0,
                    cpuName: 'AMD Ryzen 7 5700X',
                    fallback: true
                });
            }
        });
    });
}

// ==========================================
// Модуль классификации моделей, дат и преемственности
// ==========================================

const MONTH_NAMES_RU = {
    'jan': 'Январь', 'feb': 'Февраль', 'mar': 'Март', 'apr': 'Апрель',
    'may': 'Май', 'jun': 'Июнь', 'jul': 'Июль', 'aug': 'Август',
    'sep': 'Сентябрь', 'oct': 'Октябрь', 'nov': 'Ноябрь', 'dec': 'Декабрь'
};

const OLLAMA_DATES_CACHE = {
    timestamp: 0,
    families: [],
    dates: {
        'qwen3.8-flash-next': 'Сентябрь 2026',
        'qwen3.6': 'Сентябрь 2026',
        'qwen3.5': 'Сентябрь 2026',
        'qwen3.8': 'Август 2026',
        'qwen3-vl': 'Октябрь 2025',
        'qwen3-embedding': 'Сентябрь 2025',
        'qwen3-next': 'Декабрь 2025',
        'qwen3': 'Октябрь 2025',
        'qwen2.5': 'Сентябрь 2024',
        'qwen3-coder': 'Сентябрь 2025'
    }
};

function parseOllamaDate(rawDateStr) {
    if (!rawDateStr) return null;
    const m = rawDateStr.match(/([A-Za-z]{3})\s+([0-9]{1,2}),?\s+([0-9]{4})/);
    if (m) {
        const mon = MONTH_NAMES_RU[m[1].toLowerCase()] || m[1];
        return `${mon} ${m[3]}`;
    }
    return rawDateStr;
}

let updateCheckCache = {
    timestamp: 0,
    data: null
};

function getLocalOllamaModels() {
    return new Promise((resolve) => {
        const req = http.get({
            hostname: '127.0.0.1',
            port: OLLAMA_PORT,
            path: '/api/tags',
            timeout: 5000
        }, (res) => {
            let data = '';
            res.on('data', c => data += c);
            res.on('end', () => {
                try {
                    const json = JSON.parse(data);
                    resolve(json.models || []);
                } catch (e) {
                    resolve([]);
                }
            });
        });
        req.on('error', () => resolve([]));
        req.on('timeout', () => { req.destroy(); resolve([]); });
    });
}

function fetchOllamaQwenLibrary() {
    return new Promise((resolve) => {
        const now = Date.now();
        if (OLLAMA_DATES_CACHE.families.length > 0 && (now - OLLAMA_DATES_CACHE.timestamp < 15 * 60 * 1000)) {
            return resolve(OLLAMA_DATES_CACHE);
        }

        const req = https.get('https://ollama.com/search?q=qwen', {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
            },
            timeout: 6000
        }, res => {
            let data = '';
            res.on('data', c => data += c);
            res.on('end', () => {
                const found = [];
                const items = data.match(/<li[\s\S]*?<\/li>/gi) || [];
                for (const it of items) {
                    const nameMatch = it.match(/href="\/library\/([^"/]+)"/);
                    if (!nameMatch) continue;
                    const family = nameMatch[1].trim().toLowerCase();
                    if (family.includes('qwen') && !found.includes(family)) {
                        found.push(family);
                    }
                    const titleMatch = it.match(/title="([A-Za-z]{3}\s+[0-9]{1,2},\s+[0-9]{4}[^"]*)"/);
                    if (titleMatch) {
                        const parsed = parseOllamaDate(titleMatch[1]);
                        if (parsed) {
                            OLLAMA_DATES_CACHE.dates[family] = parsed;
                        }
                    }
                }
                if (found.length > 0) {
                    OLLAMA_DATES_CACHE.families = found;
                }
                OLLAMA_DATES_CACHE.timestamp = now;
                resolve(OLLAMA_DATES_CACHE);
            });
        });
        req.on('error', () => resolve(OLLAMA_DATES_CACHE));
        req.on('timeout', () => { req.destroy(); resolve(OLLAMA_DATES_CACHE); });
    });
}

function computeHardwareCompatibility(sizeGB, hwSpecs = null) {
    const vramLimit = hwSpecs && hwSpecs.vramGB ? hwSpecs.vramGB : 8.0;
    const ramLimit = hwSpecs && hwSpecs.ramTotalGB ? hwSpecs.ramTotalGB : 16.0;
    const val = Number(sizeGB) || 0;

    if (val <= (vramLimit - 0.5)) {
        return {
            status: 'gpu_ready',
            color: '#2ea043',
            badge: '🟢 100% GPU',
            text: `Влезает в видеопамять (${vramLimit} GB VRAM) — максимальная скорость`
        };
    }
    if (val <= (ramLimit - 2.0)) {
        return {
            status: 'ram_offload',
            color: '#d29922',
            badge: '🟡 RAM Offload',
            text: `Требует оперативную память (${ramLimit} GB RAM) — умеренная скорость`
        };
    }
    if (val <= (vramLimit + ramLimit - 4.0)) {
        return {
            status: 'ram_offload',
            color: '#d29922',
            badge: '🟡 RAM Offload',
            text: `Частичная выгрузка в VRAM + ${ramLimit} GB RAM (~5–8 ток/сек)`
        };
    }
    return {
        status: 'oom_risk',
        color: '#da3633',
        badge: '🔴 Недостаточно памяти',
        text: `Слишком тяжелая модель (${val.toFixed(1)} GB) для ${vramLimit} GB VRAM + ${ramLimit} GB RAM`
    };
}

function parseSizeGBFromLabel(sizeStr) {
    if (!sizeStr) return 5.0;
    const m = String(sizeStr).match(/(\d+(?:\.\d+)?)\s*(GB|MB)/i);
    if (!m) return 5.0;
    const num = parseFloat(m[1]);
    return m[2].toUpperCase() === 'MB' ? +(num / 1024).toFixed(2) : num;
}

function inferTierAndVram(name, isUncensored = false) {
    if (name.includes('8x7b') || name.includes('8x22b')) {
        return { tier: 'expert', tierName: '🧠 MoE (Экспертный)', approxVramGB: name.includes('8x22b') ? 80.0 : 26.0 };
    }
    if (name.includes('70b') || name.includes('72b')) {
        return { tier: 'expert', tierName: '🧠 Экспертный (70B)', approxVramGB: 42.0 };
    }
    if (name.includes('27b') || name.includes('30b') || name.includes('32b')) {
        return { tier: 'expert', tierName: '🧠 Экспертный', approxVramGB: 18.0 };
    }
    if (name.includes('14b')) {
        return { tier: 'expert', tierName: '🧠 Продвинутый (14B)', approxVramGB: 9.0 };
    }
    if (name.includes('0.5b') || name.includes('0.6b') || name.includes('0.8b') || name.includes(':1b') || name.includes('1.5b') || name.includes('1.7b') || name.includes(':2b')) {
        return { tier: 'light', tierName: '🌱 Ультралегкий', approxVramGB: isUncensored ? 2.0 : 1.8 };
    }
    if (name.includes(':3b') || name.includes(':4b')) {
        return { tier: 'balanced', tierName: '🚀 Сбалансированный', approxVramGB: 3.4 };
    }
    if (name.includes('7b') || name.includes('8b') || name.includes('9b')) {
        return { tier: 'flagship', tierName: '⚡ Флагман', approxVramGB: isUncensored ? 5.5 : 5.4 };
    }
    return { tier: 'expert', tierName: '🧠 Экспертный', approxVramGB: 18.0 };
}

function classifyModel(modelName, sizeBytes = 0, hwSpecs = null) {
    const name = (modelName || '').toLowerCase();
    let modelClass = 'chat';
    let className = '💬 Чат и логика';
    let tier = 'flagship';
    let tierName = '⚡ Флагман (8 GB GPU)';
    let generation = '3.0';
    let releaseDate = '2025';
    let lineage = '';
    let approxVramGB = 5.0;

    if (name.includes('embed')) {
        modelClass = 'embedding';
        className = '⚡ Эмбеддинги (RAG)';
        if (name.includes('0.6b')) {
            tier = 'light';
            tierName = '🌱 Ультралегкий';
            approxVramGB = 0.6;
        } else {
            tier = 'flagship';
            tierName = '⚡ Продвинутый';
            approxVramGB = 4.0;
        }
        generation = '3.0';
        releaseDate = OLLAMA_DATES_CACHE.dates['qwen3-embedding'] || 'Сентябрь 2025';
        lineage = 'Официальные эмбеддинги Qwen для базы знаний (контекст 32K)';
    } else if (name.includes('abliterated') || name.includes('uncensored') || name.includes('dolphin')) {
        modelClass = 'uncensored';
        className = '🔓 Без цензуры (Abliterated)';
        const inferred = inferTierAndVram(name, true);
        tier = inferred.tier;
        tierName = inferred.tierName;
        approxVramGB = inferred.approxVramGB;

        if (name.includes('dolphin3')) {
            generation = 'Dolphin 3.0';
            releaseDate = '2025';
            lineage = 'Dolphin 3.0 (база Llama 3.1) — без фильтров отказа';
        } else if (name.includes('dolphin-mistral')) {
            generation = 'Dolphin Mistral';
            releaseDate = '2024';
            lineage = 'Dolphin Mistral 7B — без фильтров отказа';
        } else if (name.includes('dolphin') && name.includes('qwen')) {
            generation = 'Dolphin Qwen';
            releaseDate = '2024–2025';
            lineage = 'Dolphin Qwen — открытая модель без шаблонов отказа';
        } else if (name.includes('dolphin-llama3')) {
            generation = 'Dolphin 2.9';
            releaseDate = '2024';
            lineage = 'Dolphin Llama 3 — без фильтров отказа';
        } else {
            generation = name.includes('3.5') ? '3.5' : (name.includes('qwen3') ? '3.0' : 'Abliterated');
            releaseDate = '2025–2026';
            lineage = 'Модифицированная версия со снятыми шаблонами отказа (abliterated)';
        }
    } else if (name.includes('-vl') || name.includes('vl:')) {
        modelClass = 'vision';
        className = '👁️ Зрение (Vision)';
        if (name.includes('2b')) {
            tier = 'light';
            tierName = '🌱 Компактный';
            approxVramGB = 1.8;
        } else if (name.includes('4b')) {
            tier = 'balanced';
            tierName = '🚀 Сбалансированный';
            approxVramGB = 3.3;
        } else if (name.includes('8b')) {
            tier = 'flagship';
            tierName = '⚡ Флагман';
            approxVramGB = 6.5;
        } else {
            tier = 'expert';
            tierName = '🧠 Экспертный';
            approxVramGB = 18.0;
        }
        generation = '3.0';
        releaseDate = OLLAMA_DATES_CACHE.dates['qwen3-vl'] || 'Октябрь 2025';
        lineage = 'Мультимодальное зрение: анализ фото, скриншотов и схем';
    } else if (name.includes('coder')) {
        modelClass = 'coder';
        className = '💻 Кодинг';
        if (name.includes('1.5b') || name.includes('3b')) {
            tier = 'light';
            tierName = '🌱 Компактный';
            approxVramGB = 2.0;
        } else if (name.includes('7b')) {
            tier = 'flagship';
            tierName = '⚡ Флагман';
            approxVramGB = 4.7;
        } else if (name.includes('14b')) {
            tier = 'expert';
            tierName = '🧠 Экспертный (14B)';
            approxVramGB = 9.5;
        } else {
            tier = 'expert';
            tierName = '🧠 Экспертный';
            approxVramGB = 18.0;
        }
        generation = name.includes('qwen3') ? '3.0' : '2.5';
        releaseDate = name.includes('qwen3') ? (OLLAMA_DATES_CACHE.dates['qwen3-coder'] || 'Сентябрь 2025') : (OLLAMA_DATES_CACHE.dates['qwen2.5'] || 'Сентябрь 2024');
        lineage = 'Специализированная модель для разработки и рефакторинга';
    } else if (name.startsWith('deepseek-r1') || name.startsWith('llama3') || name.startsWith('mistral') || name.startsWith('mixtral') || name.startsWith('gemma')) {
        modelClass = 'openweights';
        className = '🌐 Open-Weights';
        const inferred = inferTierAndVram(name, false);
        tier = inferred.tier;
        tierName = inferred.tierName;
        approxVramGB = inferred.approxVramGB;

        if (name.startsWith('deepseek-r1')) {
            generation = 'DeepSeek-R1';
            releaseDate = 'Январь 2025';
            lineage = 'DeepSeek-R1 Distill — пошаговое рассуждение (<think>)';
        } else if (name.startsWith('llama3.2')) {
            generation = 'Llama 3.2';
            releaseDate = 'Сентябрь 2024';
            lineage = 'Meta Llama 3.2 — компактная открытая модель';
        } else if (name.startsWith('llama3.1') || name.startsWith('llama3')) {
            generation = 'Llama 3.1';
            releaseDate = 'Июль 2024';
            lineage = 'Meta Llama 3.1 — мировой стандарт open-weights';
        } else if (name.startsWith('mixtral')) {
            generation = 'Mixtral MoE';
            releaseDate = 'Декабрь 2023';
            lineage = 'Mistral AI Mixture-of-Experts (8x7B / 8x22B)';
        } else if (name.startsWith('mistral')) {
            generation = 'Mistral v0.3';
            releaseDate = 'Май 2024';
            lineage = 'Mistral AI 7B — быстрая универсальная модель';
        } else {
            generation = 'Gemma 2';
            releaseDate = 'Июнь 2024';
            lineage = 'Google Gemma 2 — эффективная открытая модель (контекст 8K)';
        }
    } else {
        modelClass = 'chat';
        className = '💬 Чат и логика';
        const inferred = inferTierAndVram(name, false);
        tier = inferred.tier;
        tierName = inferred.tierName;
        approxVramGB = name.includes(':7b') || name.includes(':8b') || name.includes(':9b') ? 6.6 : inferred.approxVramGB;

        if (name.startsWith('qwen3.8-flash-next')) {
            generation = '3.8-Next';
            releaseDate = OLLAMA_DATES_CACHE.dates['qwen3.8-flash-next'] || 'Сентябрь 2026';
            lineage = 'Экспериментальный preview будущего поколения Qwen4';
        } else if (name.startsWith('qwen3.8')) {
            generation = '3.8';
            releaseDate = OLLAMA_DATES_CACHE.dates['qwen3.8'] || 'Август 2026';
            lineage = 'Тяжелый флагман с глубоким reasoning (2026)';
        } else if (name.startsWith('qwen3.6')) {
            generation = '3.6';
            releaseDate = OLLAMA_DATES_CACHE.dates['qwen3.6'] || 'Сентябрь 2026';
            lineage = 'Поколение 3.6 (Сентябрь 2026)';
        } else if (name.startsWith('qwen3.5')) {
            generation = '3.5';
            releaseDate = OLLAMA_DATES_CACHE.dates['qwen3.5'] || 'Сентябрь 2026';
            if (tier === 'flagship') lineage = 'Преемник Qwen3 8B (Поколение 3.5, 2026)';
            else if (tier === 'balanced') lineage = 'Преемник Qwen3 4B (Поколение 3.5, 2026)';
            else lineage = 'Преемник Qwen3 1.7B (Поколение 3.5, 2026)';
        } else if (name.startsWith('qwen3')) {
            generation = '3.0';
            releaseDate = OLLAMA_DATES_CACHE.dates['qwen3'] || 'Октябрь 2025';
            lineage = 'Классическое поколение Qwen3 (Октябрь 2025)';
        } else if (name.startsWith('qwen2.5')) {
            generation = '2.5';
            releaseDate = OLLAMA_DATES_CACHE.dates['qwen2.5'] || 'Сентябрь 2024';
            lineage = 'Предыдущее поколение Qwen2.5 (Сентябрь 2024)';
        }
    }

    if (sizeBytes > 0) {
        approxVramGB = Math.max(approxVramGB, parseFloat((sizeBytes / 1024 / 1024 / 1024).toFixed(1)));
    }

    const compatibility = computeHardwareCompatibility(approxVramGB, hwSpecs);

    return {
        class: modelClass,
        className: className,
        tier: tier,
        tierName: tierName,
        generation: generation,
        releaseDate: releaseDate,
        lineage: lineage,
        approxVramGB: approxVramGB,
        compatibility: compatibility
    };
}

// Расчет числового веса модели для хронологической сортировки от старых годов к новым (2024 -> 2025 -> 2026)
function getModelSortScore(m) {
    if (!m) return 0;
    const name = (m.name || m.tag || '').toLowerCase();
    const c = m.classification || {};
    const relDate = (c.releaseDate || m.releaseDate || '').toLowerCase();
    const gen = (c.generation || m.generation || '').toString();

    let year = 2025;
    let month = 1;

    // 1. Определение эпохи / года выпуска
    if (relDate.includes('2023') || name.includes('mixtral')) {
        year = 2023;
        month = 12;
    } else if (relDate.includes('2024') || name.includes('qwen2.5') || gen === '2.5' || name.includes('llama3') || name.includes('mistral') || name.includes('gemma2')) {
        year = 2024;
        if (relDate.includes('май')) month = 5;
        else if (relDate.includes('июн')) month = 6;
        else if (relDate.includes('июл')) month = 7;
        else if (relDate.includes('сентябр')) month = 9;
        else month = 8;
    } else if (relDate.includes('2026') || name.includes('qwen3.5') || name.includes('qwen3.6') || name.includes('qwen3.8') || gen.startsWith('3.5') || gen.startsWith('3.6') || gen.startsWith('3.8')) {
        year = 2026;
        if (relDate.includes('август') || name.includes('3.8')) month = 8;
        else if (relDate.includes('сентябр') || name.includes('3.5') || name.includes('3.6')) month = 9;
        else month = 10;
    } else if (relDate.includes('2025') || name.includes('qwen3') || name.includes('deepseek-r1') || name.includes('dolphin3') || gen === '3.0') {
        year = 2025;
        if (name.includes('deepseek-r1') || relDate.includes('январ')) month = 1;
        else if (relDate.includes('май')) month = 5;
        else if (relDate.includes('сентябр')) month = 9;
        else if (relDate.includes('октябр')) month = 10;
        else month = 6;
    }

    // 2. Параметры модели (в миллионах) для сортировки от меньших к большим внутри одного поколения
    let paramScore = 0;
    if (name.includes('8x7b')) {
        paramScore = 56000;
    } else {
        const match = name.match(/(\d+(?:\.\d+)?)\s*b/i);
        if (match) {
            paramScore = parseFloat(match[1]) * 1000;
        } else if (m.size) {
            paramScore = m.size / (1024 * 1024);
        }
    }

    return (year * 10000000) + (month * 100000) + paramScore;
}

function buildQwenCatalog(installedModels = [], hwSpecs = null) {
    const installedNames = (installedModels || []).map(m => (m.name || '').toLowerCase());
    const d = OLLAMA_DATES_CACHE.dates || {};
    const date25 = d['qwen2.5'] || 'Сентябрь 2024';
    const dateCoder25 = d['qwen2.5-coder'] || 'Май 2025';
    const date3 = d['qwen3'] || 'Октябрь 2025';
    const dateVL = d['qwen3-vl'] || 'Октябрь 2025';
    const dateEmbed = d['qwen3-embedding'] || 'Сентябрь 2025';
    const dateCoder3 = d['qwen3-coder'] || 'Сентябрь 2025';
    const date35 = d['qwen3.5'] || 'Сентябрь 2026';
    const date36 = d['qwen3.6'] || 'Сентябрь 2026';
    const date38 = d['qwen3.8'] || 'Август 2026';

    const catalogRaw = [
        // ===== ЧАТ И ЛОГИКА: 🌱 УЛЬТРАЛЕГКИЙ (0.5B – 2B) =====
        {
            tag: 'qwen2.5:0.5b',
            displayName: 'Qwen2.5 0.5B',
            family: 'qwen2.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: '2.5',
            releaseDate: date25,
            params: '0.5B',
            sizeApprox: '398 MB',
            description: 'Базовая ультракомпактная модель 2024 года. Моментальный запуск при минимальном потреблении.',
            lineage: 'Qwen2.5 0.5B ➔ Qwen3 0.6B ➔ Qwen3.5 0.8B',
            predecessor: null,
            successor: 'qwen3:0.6b',
            parameterAlert: null
        },
        {
            tag: 'qwen2.5:1.5b',
            displayName: 'Qwen2.5 1.5B',
            family: 'qwen2.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: '2.5',
            releaseDate: date25,
            params: '1.5B',
            sizeApprox: '986 MB',
            description: 'Популярная легкая модель поколения 2.5 для быстрых ответов.',
            lineage: 'Qwen2.5 1.5B ➔ Qwen3 1.7B ➔ Qwen3.5 2B',
            predecessor: null,
            successor: 'qwen3:1.7b',
            parameterAlert: null
        },
        {
            tag: 'qwen3:0.6b',
            displayName: 'Qwen3 0.6B',
            family: 'qwen3',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: '3.0',
            releaseDate: date3,
            params: '0.6B',
            sizeApprox: '522 MB',
            description: 'Поколение 3.0: ультралегкая модель с мгновенным откликом (100+ ток/сек).',
            lineage: 'Qwen2.5 0.5B ➔ Qwen3 0.6B ➔ Qwen3.5 0.8B',
            predecessor: 'qwen2.5:0.5b',
            successor: 'qwen3.5:0.8b',
            parameterAlert: null
        },
        {
            tag: 'qwen3:1.7b',
            displayName: 'Qwen3 1.7B',
            family: 'qwen3',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: '3.0',
            releaseDate: date3,
            params: '1.7B',
            sizeApprox: '1.4 GB',
            description: 'Эффективная модель поколения 3.0 с хорошим балансом рассуждений и скорости.',
            lineage: 'Qwen2.5 1.5B ➔ Qwen3 1.7B ➔ Qwen3.5 2B',
            predecessor: 'qwen2.5:1.5b',
            successor: 'qwen3.5:2b',
            parameterAlert: null
        },
        {
            tag: 'qwen3.5:0.8b',
            displayName: 'Qwen3.5 0.8B',
            family: 'qwen3.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: '3.5',
            releaseDate: date35,
            params: '0.8B',
            sizeApprox: '1.0 GB',
            description: 'Новейшая микро-модель 2026 года с блоком рассуждений и низким расходом памяти.',
            lineage: 'Qwen3 0.6B ➔ Qwen3.5 0.8B',
            predecessor: 'qwen3:0.6b',
            successor: null,
            parameterAlert: '✅ Прямой наследник Qwen3 0.6B: параметры выросли до 0.8B (+0.2B), сохраняя ультранизкий вес ~1.0 GB.'
        },
        {
            tag: 'qwen3.5:2b',
            displayName: 'Qwen3.5 2B',
            family: 'qwen3.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: '3.5',
            releaseDate: date35,
            params: '2.3B',
            sizeApprox: '2.7 GB',
            description: 'Новейшая легкая модель 2026 года: глубокое пошаговое мышление (<thought>), опережает 7B прошлых лет.',
            lineage: 'Qwen2.5 1.5B ➔ Qwen3 1.7B ➔ Qwen3.5 2B',
            predecessor: 'qwen3:1.7b',
            successor: null,
            parameterAlert: '⚠️ Внимание: модель является прямым наследником Qwen3 1.7B, но параметров стало больше: 2.3B вместо 1.7B (+35% к памяти: 2.7 GB против 1.4 GB).'
        },

        // ===== ЧАТ И ЛОГИКА: 🚀 СБАЛАНСИРОВАННЫЙ (3B – 4B) =====
        {
            tag: 'qwen2.5:3b',
            displayName: 'Qwen2.5 3B',
            family: 'qwen2.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'balanced',
            tierName: '🚀 Сбалансированный',
            generation: '2.5',
            releaseDate: date25,
            params: '3.1B',
            sizeApprox: '1.9 GB',
            description: 'Сбалансированная модель 2024 года, проверенная надежная классика.',
            lineage: 'Qwen2.5 3B ➔ Qwen3 4B ➔ Qwen3.5 4B',
            predecessor: null,
            successor: 'qwen3:4b',
            parameterAlert: null
        },
        {
            tag: 'qwen3:4b',
            displayName: 'Qwen3 4B',
            family: 'qwen3',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'balanced',
            tierName: '🚀 Сбалансированный',
            generation: '3.0',
            releaseDate: date3,
            params: '4.7B',
            sizeApprox: '2.6 GB',
            description: 'Популярная рабочая лошадка первого поколения Qwen3. Скорость до 70 ток/сек.',
            lineage: 'Qwen2.5 3B ➔ Qwen3 4B ➔ Qwen3.5 4B',
            predecessor: 'qwen2.5:3b',
            successor: 'qwen3.5:4b',
            parameterAlert: null
        },
        {
            tag: 'qwen3.5:4b',
            displayName: 'Qwen3.5 4B',
            family: 'qwen3.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'balanced',
            tierName: '🚀 Сбалансированный',
            generation: '3.5',
            releaseDate: date35,
            params: '4.7B',
            sizeApprox: '3.4 GB',
            description: 'Новейшая модель 2026 года: баланс интеллекта, скорости 60+ ток/сек и мультимодальности.',
            lineage: 'Qwen3 4B ➔ Qwen3.5 4B',
            predecessor: 'qwen3:4b',
            successor: null,
            parameterAlert: '✅ Сопоставимый размер параметров (~4.7B, 3.4 GB). Занимает менее половины VRAM вашей видеокарты RX 6600 (8 GB).'
        },

        // ===== ЧАТ И ЛОГИКА: ⚡ ФЛАГМАН (7B – 9B) =====
        {
            tag: 'qwen2.5:7b',
            displayName: 'Qwen2.5 7B',
            family: 'qwen2.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: '2.5',
            releaseDate: date25,
            params: '7.6B',
            sizeApprox: '4.7 GB',
            description: 'Классический флагман 2024 года, завоевавший признание среди open-source нейросетей.',
            lineage: 'Qwen2.5 7B ➔ Qwen3 8B ➔ Qwen3.5 9B',
            predecessor: null,
            successor: 'qwen3:8b',
            parameterAlert: null
        },
        {
            tag: 'qwen3:8b',
            displayName: 'Qwen3 8B',
            family: 'qwen3',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: '3.0',
            releaseDate: date3,
            params: '8.2B',
            sizeApprox: '5.2 GB',
            description: 'Хит 2025 года: золотой стандарт для видеокарт с 8 GB VRAM. Отличная логика и русский язык.',
            lineage: 'Qwen2.5 7B ➔ Qwen3 8B ➔ Qwen3.5 9B',
            predecessor: 'qwen2.5:7b',
            successor: 'qwen3.5:9b',
            parameterAlert: null
        },
        {
            tag: 'qwen3.5:9b',
            displayName: 'Qwen3.5 9B',
            family: 'qwen3.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: '3.5',
            releaseDate: date35,
            params: '9.7B',
            sizeApprox: '6.6 GB',
            description: 'Новейший флагман 2026: контекст 256K, нативное глубокое мышление (<thought>), мультимодальность.',
            lineage: 'Qwen3 8B ➔ Qwen3.5 9B',
            predecessor: 'qwen3:8b',
            successor: null,
            parameterAlert: 'ℹ️ Модель является прямым наследником Qwen3 8B, параметров стало 9.7B вместо 8.2B (+18% веса: 6.6 GB против 5.2 GB). 100% помещается в 8 GB VRAM RX 6600!'
        },

        // ===== ЧАТ И ЛОГИКА: 🧠 ЭКСПЕРТНЫЙ (27B – 35B+) =====
        {
            tag: 'qwen2.5:32b',
            displayName: 'Qwen2.5 32B',
            family: 'qwen2.5',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'expert',
            tierName: '🧠 Экспертный',
            generation: '2.5',
            releaseDate: date25,
            params: '32.8B',
            sizeApprox: '20 GB',
            description: 'Тяжелый эксперт поколения 2.5 для сложнейшего анализа данных.',
            lineage: 'Qwen2.5 32B ➔ Qwen3 30B ➔ Qwen3.6 27B ➔ Qwen3.8 27B',
            predecessor: null,
            successor: 'qwen3:30b',
            parameterAlert: null
        },
        {
            tag: 'qwen3:30b',
            displayName: 'Qwen3 30B',
            family: 'qwen3',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'expert',
            tierName: '🧠 Экспертный',
            generation: '3.0',
            releaseDate: date3,
            params: '30B',
            sizeApprox: '18 GB',
            description: 'Экспертная модель поколения 3.0. Требует выгрузки в 16 GB системной памяти.',
            lineage: 'Qwen2.5 32B ➔ Qwen3 30B ➔ Qwen3.6 27B ➔ Qwen3.8 27B',
            predecessor: 'qwen2.5:32b',
            successor: 'qwen3.6:27b',
            parameterAlert: null
        },
        {
            tag: 'qwen3.6:27b',
            displayName: 'Qwen3.6 27B',
            family: 'qwen3.6',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'expert',
            tierName: '🧠 Экспертный',
            generation: '3.6',
            releaseDate: date36,
            params: '27B',
            sizeApprox: '17 GB',
            description: 'Поколение 3.6: усиленный блок высшей математики и логических доказательств.',
            lineage: 'Qwen3 30B ➔ Qwen3.6 27B ➔ Qwen3.8 27B',
            predecessor: 'qwen3:30b',
            successor: 'qwen3.8:27b',
            parameterAlert: '⚠️ Требует 16 GB RAM Offload. Скорость генерации около 5–8 ток/сек.'
        },
        {
            tag: 'qwen3.8:27b',
            displayName: 'Qwen3.8 27B',
            family: 'qwen3.8',
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'expert',
            tierName: '🧠 Экспертный',
            generation: '3.8',
            releaseDate: date38,
            params: '27B',
            sizeApprox: '17 GB',
            description: 'Новейшая архитектура 3.8 с глубоким автономным мышлением высшего класса.',
            lineage: 'Qwen3.6 27B ➔ Qwen3.8 27B',
            predecessor: 'qwen3.6:27b',
            successor: null,
            parameterAlert: '⚠️ Модель высшего класса: требует 16 GB RAM Offload. Плотная архитектура 27B.'
        },

        // ===== 👁️ ЗРЕНИЕ (VISION) =====
        {
            tag: 'qwen3-vl:2b',
            displayName: 'Qwen3-VL 2B',
            family: 'qwen3-vl',
            modelClass: 'vision',
            className: '👁️ Зрение (Vision)',
            tier: 'light',
            tierName: '🌱 Компактный',
            generation: '3.0',
            releaseDate: dateVL,
            params: '2.2B',
            sizeApprox: '1.8 GB',
            description: 'Миниатюрная мультимодальная модель для распознавания простых изображений.',
            lineage: 'Qwen3-VL 2B',
            predecessor: null,
            successor: null,
            parameterAlert: null
        },
        {
            tag: 'qwen3-vl:4b',
            displayName: 'Qwen3-VL 4B',
            family: 'qwen3-vl',
            modelClass: 'vision',
            className: '👁️ Зрение (Vision)',
            tier: 'balanced',
            tierName: '🚀 Сбалансированный',
            generation: '3.0',
            releaseDate: dateVL,
            params: '4.4B',
            sizeApprox: '3.3 GB',
            description: 'Рекомендуемая мультимодальная модель: чтение текста на картинках, скриншотах и диаграммах.',
            lineage: 'Qwen3-VL 4B',
            predecessor: null,
            successor: null,
            parameterAlert: '✅ Рекомендуется: отличный баланс распознавания скриншотов и скорости (100% в 8 GB VRAM).'
        },
        {
            tag: 'qwen3-vl:8b',
            displayName: 'Qwen3-VL 8B',
            family: 'qwen3-vl',
            modelClass: 'vision',
            className: '👁️ Зрение (Vision)',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: '3.0',
            releaseDate: dateVL,
            params: '8.3B',
            sizeApprox: '5.6 GB',
            description: 'Высокодетализированный анализ сложных схем, чертежей и рукописного текста.',
            lineage: 'Qwen3-VL 8B',
            predecessor: null,
            successor: null,
            parameterAlert: 'ℹ️ Высокое разрешение анализа, занимает 5.6 GB VRAM.'
        },

        // ===== 💻 КОДИНГ (CODER) =====
        {
            tag: 'qwen2.5-coder:1.5b',
            displayName: 'Qwen2.5-Coder 1.5B',
            family: 'qwen2.5-coder',
            modelClass: 'coder',
            className: '💻 Кодинг',
            tier: 'light',
            tierName: '🌱 Компактный',
            generation: '2.5',
            releaseDate: dateCoder25,
            params: '1.5B',
            sizeApprox: '986 MB',
            description: 'Компактная модель для автодополнения кода и написания коротких скриптов.',
            lineage: 'Qwen2.5-Coder 1.5B',
            predecessor: null,
            successor: null,
            parameterAlert: null
        },
        {
            tag: 'qwen2.5-coder:7b',
            displayName: 'Qwen2.5-Coder 7B',
            family: 'qwen2.5-coder',
            modelClass: 'coder',
            className: '💻 Кодинг',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: '2.5',
            releaseDate: dateCoder25,
            params: '7.6B',
            sizeApprox: '4.7 GB',
            description: 'Золотой стандарт для разработки под 8 GB VRAM: генерация кода, поиск багов, написание тестов.',
            lineage: 'Qwen2.5-Coder 7B',
            predecessor: null,
            successor: null,
            parameterAlert: '✅ Рекомендуется для кодинга: 100% помещается в 8 GB VRAM RX 6600.'
        },
        {
            tag: 'qwen2.5-coder:14b',
            displayName: 'Qwen2.5-Coder 14B',
            family: 'qwen2.5-coder',
            modelClass: 'coder',
            className: '💻 Кодинг',
            tier: 'expert',
            tierName: '🧠 Экспертный',
            generation: '2.5',
            releaseDate: dateCoder25,
            params: '14.8B',
            sizeApprox: '9.0 GB',
            description: 'Профессиональная разработка больших проектов и сложных алгоритмов.',
            lineage: 'Qwen2.5-Coder 14B',
            predecessor: null,
            successor: null,
            parameterAlert: '⚠️ Требует частичного оффлоада в RAM.'
        },

        // ===== ⚡ ЭМБЕДДИНГИ (RAG) =====
        {
            tag: 'qwen3-embedding:0.6b',
            displayName: 'Qwen3-Embedding 0.6B',
            family: 'qwen3-embedding',
            modelClass: 'embedding',
            className: '⚡ Эмбеддинги (RAG)',
            tier: 'light',
            tierName: '🌱 Стандарт RAG',
            generation: '3.0',
            releaseDate: dateEmbed,
            params: '0.6B',
            sizeApprox: '639 MB',
            description: 'Официальная модель эмбеддингов Qwen для локальной базы знаний с контекстом 32K.',
            lineage: 'nomic-embed-text ➔ Qwen3-Embedding 0.6B',
            predecessor: 'nomic-embed-text',
            successor: null,
            parameterAlert: '✅ Официальные эмбеддинги Qwen для RAG: превосходная семантика русского языка.'
        },
        {
            tag: 'qwen3-embedding:4b',
            displayName: 'Qwen3-Embedding 4B',
            family: 'qwen3-embedding',
            modelClass: 'embedding',
            className: '⚡ Эмбеддинги (RAG)',
            tier: 'flagship',
            tierName: '⚡ Продвинутый',
            generation: '3.0',
            releaseDate: dateEmbed,
            params: '4.2B',
            sizeApprox: '2.5 GB',
            description: 'Тяжелая модель векторного поиска для огромных корпоративных архивов.',
            lineage: 'Qwen3-Embedding 4B',
            predecessor: null,
            successor: null,
            parameterAlert: null
        },

        // ===== 🔓 БЕЗ ЦЕНЗУРЫ И ОГРАНИЧЕНИЙ (ABLITERATED) =====
        {
            tag: 'huihui_ai/qwen3.5-abliterated:4B',
            displayName: 'Qwen3.5 4B Abliterated',
            family: 'qwen3.5-abliterated',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'balanced',
            tierName: '🚀 Сбалансированный',
            generation: '3.5',
            releaseDate: '2026',
            params: '4B',
            sizeApprox: '3.3 GB',
            description: 'Версия Qwen3.5 со снятыми шаблонами отказа (abliterated). 100% в VRAM, моментальный отклик, зрение (Vision) и кодинг.',
            lineage: 'Qwen3.5 4B ➔ huihui_ai/qwen3.5-abliterated:4B',
            predecessor: null,
            successor: null,
            parameterAlert: '⚡ Идеально для RX 6600 (8 GB): модель весит всего 3.3 GB, оставляя 4.7 GB VRAM под длинный контекст и картинки.'
        },
        {
            tag: 'huihui_ai/qwen3.5-abliterated:9b',
            displayName: 'Qwen3.5 9B Abliterated',
            family: 'qwen3.5-abliterated',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: '3.5',
            releaseDate: '2026',
            params: '9B',
            sizeApprox: '6.6 GB',
            description: 'Флагманская модель без ограничений. Глубокие рассуждения, поддержка Vision и текста, помещается в 8 GB VRAM.',
            lineage: 'Qwen3.5 9B ➔ huihui_ai/qwen3.5-abliterated:9b',
            predecessor: null,
            successor: null,
            parameterAlert: '🟢 Полностью помещается в 8 GB VRAM (занимает ~6.6 GB).'
        },
        {
            tag: 'huihui_ai/qwen3.5-abliterated:2B',
            displayName: 'Qwen3.5 2B Abliterated',
            family: 'qwen3.5-abliterated',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: '3.5',
            releaseDate: '2026',
            params: '2B',
            sizeApprox: '1.9 GB',
            description: 'Ультралегкая компактная модель 2026 года без фильтров отказа. Мгновенная генерация 60+ ток/сек.',
            lineage: 'Qwen3.5 2B ➔ huihui_ai/qwen3.5-abliterated:2B',
            predecessor: null,
            successor: null,
            parameterAlert: null
        },
        {
            tag: 'dolphin-llama3:8b',
            displayName: 'Dolphin Llama3 8B',
            family: 'dolphin',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'Llama3',
            releaseDate: '2024–2025',
            params: '8B',
            sizeApprox: '4.7 GB',
            description: 'Легендарная модель от Эрика Хартфорда с полностью снятыми фильтрами и шаблонами отказа.',
            lineage: 'Llama 3 8B ➔ Dolphin Llama3 8B',
            predecessor: null,
            successor: null,
            parameterAlert: '🟢 100% в VRAM (4.7 GB). Проверенная временем классика сообщества открытого ИИ.'
        },
        {
            tag: 'richardyoung/qwen3-14b-abliterated',
            displayName: 'Qwen3 14B Abliterated',
            family: 'qwen3-abliterated',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'expert',
            tierName: '🧠 Экспертный (14B)',
            generation: '3.0',
            releaseDate: '2025',
            params: '14B',
            sizeApprox: '9.0 GB',
            description: 'Тяжелая продвинутая модель для сложных логических цепочек и кода со снятыми ограничениями.',
            lineage: 'Qwen3 14B ➔ richardyoung/qwen3-14b-abliterated',
            predecessor: null,
            successor: null,
            parameterAlert: '🟡 Требует частичного оффлоада в ваши 16 GB системной RAM.'
        },
        {
            tag: 'dolphin3:8b',
            displayName: 'Dolphin 3.0 Llama 3.1 8B',
            family: 'dolphin',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'Dolphin 3.0',
            releaseDate: '2025',
            params: '8B',
            sizeApprox: '4.9 GB',
            description: 'Актуальная версия Dolphin 3.0 на базе Llama 3.1 без цензуры, нотаций и отказов. Отлично подходит для кода, агента и системных вопросов.',
            lineage: 'Llama 3.1 8B ➔ Dolphin 3.0 (8B)',
            predecessor: 'dolphin-llama3:8b',
            successor: null,
            parameterAlert: '🟢 100% помещается в 8 GB VRAM (4.9 GB), поддерживает контекст до 32K.'
        },
        {
            tag: 'dolphin-mistral:7b',
            displayName: 'Dolphin Mistral 7B',
            family: 'dolphin-mistral',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'Dolphin Mistral',
            releaseDate: '2024',
            params: '7B',
            sizeApprox: '4.1 GB',
            description: 'Классическая незацензуренная модель на базе Mistral 7B (v2.8): лёгкая, быстрая и отзывчивая без встроенной морали.',
            lineage: 'Mistral 7B ➔ Dolphin Mistral 7B',
            predecessor: null,
            successor: null,
            parameterAlert: '🟢 Всего 4.1 GB — оставляет почти половину VRAM свободными.'
        },
        {
            tag: 'CognitiveComputations/dolphin-2.9.2-qwen2-7b:Q4_K_M',
            displayName: 'Dolphin 2.9.2 Qwen2 7B',
            family: 'dolphin-qwen',
            modelClass: 'uncensored',
            className: '🔓 Без цензуры (Abliterated)',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'Dolphin Qwen',
            releaseDate: '2024',
            params: '7.6B',
            sizeApprox: '4.7 GB',
            description: 'Сборка Dolphin на архитектуре Qwen2 7B от CognitiveComputations: сочетает качество Qwen в коде и русском языке со свободой ответов Dolphin.',
            lineage: 'Qwen2 7B ➔ Dolphin 2.9.2 Qwen2 7B',
            predecessor: null,
            successor: null,
            parameterAlert: '🟢 Квантование Q4_K_M (4.7 GB) полностью помещается в 8 GB VRAM.'
        },

        // ===== 🌐 OPEN-WEIGHTS: СТОРОННИЕ СЕМЕЙСТВА =====
        {
            tag: 'llama3.2:3b',
            displayName: 'Meta Llama 3.2 3B',
            family: 'llama3.2',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'balanced',
            tierName: '🚀 Сбалансированный',
            generation: 'Llama 3.2',
            releaseDate: 'Сентябрь 2024',
            params: '3.2B',
            sizeApprox: '2.0 GB',
            description: 'Компактная мультиязычная модель от Meta для быстрых повседневных задач и инструкций.',
            lineage: 'Meta Llama 3.2 (3B)',
            predecessor: null,
            successor: 'llama3.1:8b',
            parameterAlert: '⚡ Очень быстрая: весит всего 2.0 GB, летает на 8 GB VRAM.'
        },
        {
            tag: 'llama3.1:8b',
            displayName: 'Meta Llama 3.1 8B',
            family: 'llama3.1',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'Llama 3.1',
            releaseDate: 'Июль 2024',
            params: '8.0B',
            sizeApprox: '4.9 GB',
            description: 'Мировой стандарт open-weights моделей от Meta: стабильное следование инструкциям, длинный контекст и tool-calling.',
            lineage: 'Meta Llama 3.1 (8B)',
            predecessor: 'llama3.2:3b',
            successor: null,
            parameterAlert: '🟢 100% в 8 GB VRAM (4.9 GB).'
        },
        {
            tag: 'deepseek-r1:7b',
            displayName: 'DeepSeek-R1 Distill Qwen 7B',
            family: 'deepseek-r1',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'DeepSeek-R1',
            releaseDate: 'Январь 2025',
            params: '7.6B',
            sizeApprox: '4.7 GB',
            description: 'Дистилляция DeepSeek-R1 на базе Qwen-7B с пошаговыми рассуждениями (<think>): сильна в математике, логике и алгоритмах.',
            lineage: 'DeepSeek-R1 ➔ Distill Qwen 7B',
            predecessor: null,
            successor: 'deepseek-r1:8b',
            parameterAlert: '🟢 Оптимальна под 8 GB VRAM (4.7 GB) с глубоким chain-of-thought.'
        },
        {
            tag: 'deepseek-r1:8b',
            displayName: 'DeepSeek-R1 Distill Llama 8B',
            family: 'deepseek-r1',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'DeepSeek-R1',
            releaseDate: 'Январь 2025',
            params: '8.0B',
            sizeApprox: '5.2 GB',
            description: 'Версия DeepSeek-R1 Distill 8B на базе Llama 3.1: открытый флагман логических рассуждений и анализа кода.',
            lineage: 'DeepSeek-R1 Distill 7B ➔ 8B',
            predecessor: 'deepseek-r1:7b',
            successor: 'deepseek-r1:14b',
            parameterAlert: '🟢 5.2 GB — полностью умещается в 8 GB VRAM RX 6600.'
        },
        {
            tag: 'deepseek-r1:14b',
            displayName: 'DeepSeek-R1 Distill Qwen 14B',
            family: 'deepseek-r1',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'expert',
            tierName: '🧠 Продвинутый (14B)',
            generation: 'DeepSeek-R1',
            releaseDate: 'Январь 2025',
            params: '14.8B',
            sizeApprox: '9.0 GB',
            description: 'Старшая дистиллированная reasoning-модель DeepSeek-R1 14B для сложного анализа, математики и архитектуры кода.',
            lineage: 'DeepSeek-R1 Distill 8B ➔ 14B',
            predecessor: 'deepseek-r1:8b',
            successor: null,
            parameterAlert: '🟡 Вес 9.0 GB — требует частичной выгрузки в 16 GB системной RAM.'
        },
        {
            tag: 'mistral:7b',
            displayName: 'Mistral 7B v0.3',
            family: 'mistral',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'Mistral v0.3',
            releaseDate: 'Май 2024',
            params: '7.2B',
            sizeApprox: '4.4 GB',
            description: 'Быстрая европейская open-weights модель от Mistral AI (v0.3): расширенный словарь, поддержка function calling и высокая скорость.',
            lineage: 'Mistral 7B v0.3',
            predecessor: null,
            successor: 'mixtral:8x7b',
            parameterAlert: '🟢 4.4 GB — уверенно работает полностью в VRAM.'
        },
        {
            tag: 'mixtral:8x7b',
            displayName: 'Mixtral 8x7B MoE',
            family: 'mixtral',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'expert',
            tierName: '🧠 MoE (8x7B)',
            generation: 'Mixtral MoE',
            releaseDate: 'Декабрь 2023',
            params: '46.7B MoE',
            sizeApprox: '26 GB',
            description: 'Mixture-of-Experts модель от Mistral AI (8 экспертов по 7B). В каталоге для полноты линейки семейства Mistral.',
            lineage: 'Mistral 7B ➔ Mixtral 8x7B MoE',
            predecessor: 'mistral:7b',
            successor: null,
            parameterAlert: '🔴 Весит ~26 GB: слишком велика для конфигурации 8 GB VRAM + 16 GB RAM.'
        },
        {
            tag: 'gemma2:2b',
            displayName: 'Google Gemma 2 2B',
            family: 'gemma2',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'light',
            tierName: '🌱 Ультралегкий',
            generation: 'Gemma 2',
            releaseDate: 'Июль 2024',
            params: '2.6B',
            sizeApprox: '1.6 GB',
            description: 'Компактная и очень качественная для своего размера модель от Google с контекстом 8K.',
            lineage: 'Google Gemma 2 (2B)',
            predecessor: null,
            successor: 'gemma2:9b',
            parameterAlert: '⚡ Занимает всего 1.6 GB VRAM, нативный контекст 8K.'
        },
        {
            tag: 'gemma2:9b',
            displayName: 'Google Gemma 2 9B',
            family: 'gemma2',
            modelClass: 'openweights',
            className: '🌐 Open-Weights',
            tier: 'flagship',
            tierName: '⚡ Флагман',
            generation: 'Gemma 2',
            releaseDate: 'Июнь 2024',
            params: '9.2B',
            sizeApprox: '5.4 GB',
            description: 'Открытая модель от Google: сильна в формулировке текстов, объяснениях и переводах. Контекстное окно 8K.',
            lineage: 'Google Gemma 2 (2B ➔ 9B)',
            predecessor: 'gemma2:2b',
            successor: null,
            parameterAlert: '🟢 5.4 GB — помещается в 8 GB VRAM (контекст 8K).'
        }
    ];

    return catalogRaw.map(item => {
        const isInstalled = installedNames.some(n => n === item.tag.toLowerCase() || n.startsWith(item.tag.toLowerCase() + ':'));
        const compat = computeHardwareCompatibility(parseSizeGBFromLabel(item.sizeApprox), hwSpecs);

        return {
            ...item,
            isInstalled: isInstalled,
            compatibility: compat
        };
    });
}

async function checkModelUpdates(forceRefresh = false) {
    const now = Date.now();
    if (!forceRefresh && updateCheckCache.data && (now - updateCheckCache.timestamp < 10 * 60 * 1000)) {
        return updateCheckCache.data;
    }

    const installed = await getLocalOllamaModels();
    const libData = await fetchOllamaQwenLibrary();
    const libraryFamilies = libData.families || [];
    const installedNames = installed.map(m => m.name.toLowerCase());
    const hwSpecs = await getSystemHardwareSpecs();
    const updates = [];

    // Правила поколенческих обновлений внутри классов и подклассов мощности
    const upgradeRules = [
        {
            check: (name) => name.startsWith('qwen3:8b') || name.startsWith('qwen2.5:7b') || name.startsWith('qwen2.5:8b'),
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'flagship',
            tierName: '⚡ ФЛАГМАН (8 GB GPU)',
            targetFamily: 'qwen3.5',
            targetTag: 'qwen3.5:9b',
            title: 'Qwen3.5 9B (Новый флагман)',
            sizeApprox: '~6.6 GB',
            oldGeneration: 'Qwen3 8B (Октябрь 2025)',
            newGeneration: 'Qwen3.5 9B (Сентябрь 2026)',
            targetReleaseDate: OLLAMA_DATES_CACHE.dates['qwen3.5'] || 'Сентябрь 2026',
            evolution: 'Qwen3 8B (Окт 2025) ➔ Qwen3.5 9B (Сент 2026) ✨',
            parameterAlert: 'ℹ️ Модель является прямым наследником Qwen3 8B, параметров стало 9.7B вместо 8.2B (+18% веса: 6.6 GB против 5.2 GB). 100% помещается в 8 GB VRAM RX 6600!',
            description: 'Новейшая архитектура: контекст 256K, глубокое мышление (<thought>), нативная мультимодальность. Заметно умнее первого поколения Qwen3 8B.',
            vramFit: '🟢 100% GPU (Идеально под RX 6600 8 GB)'
        },
        {
            check: (name) => name.startsWith('qwen3:4b') || name.startsWith('qwen2.5:3b'),
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'balanced',
            tierName: '🚀 СБАЛАНСИРОВАННЫЙ',
            targetFamily: 'qwen3.5',
            targetTag: 'qwen3.5:4b',
            title: 'Qwen3.5 4B (Сверхбыстрая)',
            sizeApprox: '~3.4 GB',
            oldGeneration: 'Qwen3 4B (Октябрь 2025)',
            newGeneration: 'Qwen3.5 4B (Сентябрь 2026)',
            targetReleaseDate: OLLAMA_DATES_CACHE.dates['qwen3.5'] || 'Сентябрь 2026',
            evolution: 'Qwen3 4B (Окт 2025) ➔ Qwen3.5 4B (Сент 2026) ⚡',
            parameterAlert: '✅ Сопоставимый размер параметров (~4.7B, 3.4 GB). Занимает менее половины VRAM вашей RX 6600 (8 GB).',
            description: 'Свежая компактная модель: скорость 50+ токенов/сек, мультимодальность, идеально для повседневных задач.',
            vramFit: '🟢 100% GPU (Занимает меньше половины VRAM)'
        },
        {
            check: (name) => name.startsWith('qwen3:0.6b') || name.startsWith('qwen3:1.7b'),
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🌱 УЛЬТРАЛЕГКИЙ',
            targetFamily: 'qwen3.5',
            targetTag: 'qwen3.5:2b',
            title: 'Qwen3.5 2B (Компактная)',
            sizeApprox: '~2.7 GB',
            oldGeneration: 'Qwen3 1.7B / 0.6B (Октябрь 2025)',
            newGeneration: 'Qwen3.5 2B (Сентябрь 2026)',
            targetReleaseDate: OLLAMA_DATES_CACHE.dates['qwen3.5'] || 'Сентябрь 2026',
            evolution: 'Qwen3 0.6B/1.7B (Окт 2025) ➔ Qwen3.5 2B (Сент 2026) 🚀',
            parameterAlert: '⚠️ Внимание: модель является прямым наследником Qwen3 1.7B, но параметров стало больше: 2.3B вместо 1.7B (+35% к памяти: 2.7 GB против 1.4 GB).',
            description: 'Миниатюрная модель 2026 года нового поколения с блоками размышлений.',
            vramFit: '🟢 100% GPU (Минимальная нагрузка)'
        },
        {
            check: (name) => name.startsWith('nomic-embed-text'),
            modelClass: 'embedding',
            className: '⚡ Эмбеддинги (RAG)',
            tier: 'light',
            tierName: '🌱 СТАНДАРТ RAG',
            targetFamily: 'qwen3-embedding',
            targetTag: 'qwen3-embedding:0.6b',
            title: 'Qwen3-Embedding 0.6B (RAG)',
            sizeApprox: '~600 MB',
            oldGeneration: 'nomic-embed-text (Базовая)',
            newGeneration: 'Qwen3-Embedding 0.6B (Сентябрь 2025)',
            targetReleaseDate: OLLAMA_DATES_CACHE.dates['qwen3-embedding'] || 'Сентябрь 2025',
            evolution: 'nomic-embed-text ➔ Qwen3-Embedding 0.6B ✨',
            parameterAlert: '✅ Официальная модель эмбеддингов Qwen для базы знаний: расширенный контекст 32K, идеальная семантика русского языка.',
            description: 'Официальная модель эмбеддингов Qwen для базы знаний: расширенный контекст 32K, идеальная семантика русского языка.',
            vramFit: '🟢 Минимальный вес (600 MB)'
        }
    ];

    for (const inst of installed) {
        const instName = inst.name;
        for (const rule of upgradeRules) {
            if (rule.check(instName)) {
                const familyAvailable = libraryFamilies.length === 0 || libraryFamilies.includes(rule.targetFamily);
                const isAlreadyInstalled = installedNames.some(n => n === rule.targetTag || n.startsWith(rule.targetTag + ':'));

                if (familyAvailable && !isAlreadyInstalled) {
                    updates.push({
                        oldModel: instName,
                        oldSizeFormatted: inst.size ? `${(inst.size / 1024 / 1024 / 1024).toFixed(1)} GB` : '',
                        newModel: rule.targetTag,
                        newTitle: rule.title,
                        newSize: rule.sizeApprox,
                        modelClass: rule.modelClass,
                        className: rule.className,
                        tier: rule.tier,
                        tierName: rule.tierName,
                        oldGeneration: rule.oldGeneration,
                        newGeneration: rule.newGeneration,
                        targetReleaseDate: rule.targetReleaseDate,
                        evolution: rule.evolution,
                        parameterAlert: rule.parameterAlert,
                        description: rule.description,
                        vramFit: rule.vramFit
                    });
                }
            }
        }
    }

    // Рекомендация Vision модели, если у пользователя нет ни одной мультимодальной
    const hasVisionModel = installedNames.some(n => n.includes('-vl') || n.includes('vl:'));
    if (!hasVisionModel && (libraryFamilies.length === 0 || libraryFamilies.includes('qwen3-vl'))) {
        const isAlreadyInstalled = installedNames.some(n => n.startsWith('qwen3-vl:4b'));
        if (!isAlreadyInstalled) {
            updates.push({
                oldModel: null,
                oldSizeFormatted: null,
                newModel: 'qwen3-vl:4b',
                newTitle: 'Qwen3-VL 4B (Компьютерное зрение)',
                newSize: '~3.5 GB',
                modelClass: 'vision',
                className: '👁️ Зрение (Vision)',
                tier: 'balanced',
                tierName: '🚀 СБАЛАНСИРОВАННЫЙ',
                oldGeneration: null,
                newGeneration: 'Qwen3-VL 4B (Октябрь 2025)',
                targetReleaseDate: OLLAMA_DATES_CACHE.dates['qwen3-vl'] || 'Октябрь 2025',
                evolution: '✨ Новая категория: Компьютерное зрение',
                parameterAlert: '✅ Рекомендуется: отличный баланс распознавания скриншотов и скорости (100% в 8 GB VRAM).',
                description: 'Новая модель со зрением: распознавание изображений, графиков, скриншотов и документов.',
                vramFit: '🟢 Идеально для 8 GB VRAM'
            });
        }
    }

    const catalog = buildQwenCatalog(installed, hwSpecs);

    const result = {
        status: 'ok',
        hardware: hwSpecs,
        installedCount: installed.length,
        updatesCount: updates.length,
        updates: updates,
        catalog: catalog,
        checkedAt: getLocalTimestamp()
    };

    updateCheckCache = {
        timestamp: now,
        data: result
    };

    return result;
}

writeLog('INFO', '=== Запуск Qwen Local Server ===');
ensureOllamaRunning();

// ==========================================
// ReAct AI Agent утилиты (Deep Research)
// ==========================================

function callOllamaChat(model, messages, options = {}, onReqCreated = null) {
    return new Promise((resolve, reject) => {
        const payload = JSON.stringify({
            model: model,
            messages: messages,
            stream: false,
            options: {
                temperature: options.temperature !== undefined ? options.temperature : 0.2,
                num_ctx: options.num_ctx || 8192,
                stop: options.stop || undefined
            }
        });

        let isDone = false;
        const req = http.request({
            hostname: '127.0.0.1',
            port: OLLAMA_PORT,
            path: '/api/chat',
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(payload)
            }
        }, (res) => {
            let data = '';
            res.on('data', chunk => { data += chunk; });
            res.on('end', () => {
                isDone = true;
                try {
                    const json = JSON.parse(data);
                    if (json.message) {
                        resolve(json);
                    } else if (json.error) {
                        reject(new Error(json.error));
                    } else {
                        reject(new Error('Пустой ответ от Ollama'));
                    }
                } catch (e) {
                    reject(e);
                }
            });
        });

        if (onReqCreated) onReqCreated(req);

        req.on('error', (err) => {
            if (!isDone) reject(err);
        });

        req.setTimeout(120000, () => {
            if (!isDone) {
                req.destroy();
                reject(new Error('Таймаут запроса к Ollama (120с)'));
            }
        });

        req.write(payload);
        req.end();
    });
}

function streamOllamaChat(model, messages, onChunk, onEnd, onError, options = {}, onReqCreated = null) {
    const payload = JSON.stringify({
        model: model,
        messages: messages,
        stream: true,
        options: {
            temperature: options.temperature !== undefined ? options.temperature : 0.4,
            num_ctx: options.num_ctx || 16384
        }
    });

    let isDone = false;
    const req = http.request({
        hostname: '127.0.0.1',
        port: OLLAMA_PORT,
        path: '/api/chat',
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(payload)
        }
    }, (res) => {
        let buffer = '';
        res.on('data', chunk => {
            buffer += chunk.toString();
            const lines = buffer.split('\n');
            buffer = lines.pop();
            for (const line of lines) {
                if (!line.trim()) continue;
                try {
                    const parsed = JSON.parse(line);
                    onChunk(parsed);
                } catch (e) {}
            }
        });

        res.on('end', () => {
            isDone = true;
            if (buffer.trim()) {
                try {
                    const parsed = JSON.parse(buffer);
                    onChunk(parsed);
                } catch (e) {}
            }
            onEnd();
        });
    });

    if (onReqCreated) onReqCreated(req);

    req.on('error', (err) => {
        if (!isDone) onError(err);
    });

    req.write(payload);
    req.end();
}

function parseReActOutput(text) {
    let thought = '';
    let action = '';
    let actionParam = '';

    // Очищаем от тегов рассуждений <think>...</think>, если модель их возвращает
    let cleanText = text.replace(/<think>[\s\S]*?<\/think>/gi, '').trim();
    if (!cleanText) cleanText = text;

    // Извлечение мысли Thought:
    const thoughtMatch = cleanText.match(/Thought:\s*([\s\S]*?)(?=(?:Action:|$))/i);
    if (thoughtMatch) {
        thought = thoughtMatch[1].trim();
    } else {
        const parts = cleanText.split(/Action:/i);
        thought = parts[0].trim();
    }

    // Извлечение действия Action: (имя + остаток строки как параметр)
    const actionMatch = cleanText.match(/Action:\s*([a-zA-Z_]+)[ \t]*([^\n]*)/i);
    if (actionMatch) {
        action = actionMatch[1].toLowerCase().trim();
        let rawParam = (actionMatch[2] || '').trim();
        rawParam = rawParam.replace(/^\(\s*/, '').replace(/\s*\)\s*\.?$/, '').trim();
        rawParam = rawParam.replace(/^["'`«]+|["'`»]+$/g, '').trim();
        actionParam = rawParam;
    }

    if (!action || action === 'none') {
        if (/finish/i.test(cleanText)) {
            action = 'finish';
        } else if (/search/i.test(cleanText)) {
            action = 'search';
            const m = cleanText.match(/search\(?["']?([^"'\)\n]+)/i);
            if (m) actionParam = m[1].trim();
        } else if (/fetch/i.test(cleanText)) {
            action = 'fetch';
            const m = cleanText.match(/https?:\/\/[^\s"'<>\)]+/i);
            if (m) {
                action = 'fetch';
                actionParam = m[0];
            }
        } else {
            action = 'finish';
        }
    }

    if (action === 'fetch' && !actionParam.startsWith('http')) {
        const urlInText = cleanText.match(/https?:\/\/[^\s"'<>\)]+/);
        if (urlInText) {
            actionParam = urlInText[0];
        } else {
            action = 'finish';
        }
    }

    return { thought, action, actionParam };
}

// ==========================================
// Файловые инструменты агента (только чтение, песочница корней)
// ==========================================

const os = require('os');

// Разрешённые корневые папки. Можно переопределить: AGENT_FS_ROOTS="C:\Projects;D:\Docs"
const AGENT_FS_ROOTS = (process.env.AGENT_FS_ROOTS
    ? process.env.AGENT_FS_ROOTS.split(';')
    : [path.join(os.homedir(), 'Desktop')]
).map(p => p.trim()).filter(Boolean).map(p => path.resolve(p));

const AGENT_FS_DENY_FILE = /(^|[\\/])(\.env(\..*)?|id_rsa[^\\/]*|id_ed25519[^\\/]*|[^\\/]*\.(pem|key|pfx|p12|kdbx|ppk)|\.npmrc|\.git-credentials)$/i;
const AGENT_FS_HEAVY_DIRS = new Set(['node_modules', '.git', '__pycache__', '.venv', 'venv', 'dist', 'build', '.next']);
const AGENT_FS_MAX_LIST = 120;
const AGENT_FS_MAX_READ_BYTES = 256 * 1024;

function agentIsInsideRoot(realPath) {
    const target = realPath.toLowerCase();
    return AGENT_FS_ROOTS.some(root => {
        const r = root.toLowerCase();
        return target === r || target.startsWith(r.endsWith(path.sep) ? r : r + path.sep);
    });
}

// Возвращает { ok:true, path } либо { ok:false, error }
function resolveAgentPath(inputPath) {
    if (!inputPath) return { ok: false, error: 'Путь не указан.' };
    let p = String(inputPath).trim().replace(/^["'`]+|["'`]+$/g, '');
    if (!path.isAbsolute(p)) {
        p = path.join(AGENT_FS_ROOTS[0], p);
    }
    p = path.resolve(p);
    let real;
    try {
        real = fs.realpathSync(p);
    } catch (e) {
        return { ok: false, error: `Путь не существует: ${p}` };
    }
    if (!agentIsInsideRoot(real)) {
        return { ok: false, error: `Доступ запрещён: путь вне разрешённых папок (${AGENT_FS_ROOTS.join('; ')}).` };
    }
    return { ok: true, path: real };
}

function agentListDir(inputPath) {
    const r = resolveAgentPath(inputPath);
    if (!r.ok) return { ok: false, error: r.error };
    let stat;
    try { stat = fs.statSync(r.path); } catch (e) { return { ok: false, error: e.message }; }
    if (!stat.isDirectory()) return { ok: false, error: `Это файл, а не папка: ${r.path}. Используй read_file.` };

    let entries;
    try { entries = fs.readdirSync(r.path, { withFileTypes: true }); } catch (e) { return { ok: false, error: e.message }; }

    entries.sort((a, b) => (b.isDirectory() - a.isDirectory()) || a.name.localeCompare(b.name));
    const total = entries.length;
    const lines = entries.slice(0, AGENT_FS_MAX_LIST).map(e => {
        if (e.isDirectory()) {
            return `[папка] ${e.name}${AGENT_FS_HEAVY_DIRS.has(e.name) ? '  (служебная, пропускай)' : ''}`;
        }
        let size = '';
        try { size = ` (${Math.max(1, Math.round(fs.statSync(path.join(r.path, e.name)).size / 1024))} KB)`; } catch (_) {}
        return `[файл]  ${e.name}${size}`;
    });
    if (total > AGENT_FS_MAX_LIST) lines.push(`... и ещё ${total - AGENT_FS_MAX_LIST} элементов`);
    return { ok: true, path: r.path, count: total, text: lines.join('\n') || '(папка пуста)' };
}

function agentReadFile(inputPath, maxChars = 4000) {
    const r = resolveAgentPath(inputPath);
    if (!r.ok) return { ok: false, error: r.error };
    if (AGENT_FS_DENY_FILE.test(r.path)) {
        return { ok: false, error: 'Доступ запрещён: файл содержит секреты (ключи/пароли).' };
    }
    let stat;
    try { stat = fs.statSync(r.path); } catch (e) { return { ok: false, error: e.message }; }
    if (stat.isDirectory()) return { ok: false, error: `Это папка: ${r.path}. Используй list_dir.` };

    let buf;
    try {
        const fd = fs.openSync(r.path, 'r');
        const toRead = Math.min(stat.size, AGENT_FS_MAX_READ_BYTES);
        buf = Buffer.alloc(toRead);
        fs.readSync(fd, buf, 0, toRead, 0);
        fs.closeSync(fd);
    } catch (e) {
        return { ok: false, error: e.message };
    }
    if (buf.subarray(0, 4096).includes(0)) {
        return { ok: false, error: 'Бинарный файл, прочитать как текст нельзя.' };
    }
    let text = buf.toString('utf8');
    const fullLen = text.length;
    let truncated = false;
    if (text.length > maxChars) { text = text.substring(0, maxChars); truncated = true; }
    return {
        ok: true,
        path: r.path,
        size: stat.size,
        chars: text.length,
        truncated: truncated || stat.size > AGENT_FS_MAX_READ_BYTES,
        text: text + (truncated ? `\n... [файл обрезан: показано ${text.length} из ${fullLen}+ символов]` : '')
    };
}

// Находит в запросе пользователя путь Windows (в т.ч. с пробелами) и возвращает существующий вариант
function extractLocalPathFromQuery(query) {
    const m = String(query).match(/[A-Za-z]:\\[^\r\n"<>|?*]*/);
    if (!m) return null;
    let candidate = m[0].replace(/[\s.,;:!?)\]]+$/, '');
    // Постепенно отбрасываем слова с конца, пока путь не найдётся ("revers win -7 пожалуйста" -> "revers win -7")
    while (candidate.length > 3) {
        if (fs.existsSync(candidate)) return candidate;
        const cut = candidate.lastIndexOf(' ');
        if (cut < 3) break;
        candidate = candidate.substring(0, cut).replace(/[\s.,;:!?)\]]+$/, '');
    }
    return fs.existsSync(candidate) ? candidate : m[0].trim();
}

// ==========================================
// Stage 3: GGUF & Hugging Face Model Importer
// ==========================================

const GGUF_TEMPLATES = {
    chatml: {
        id: 'chatml',
        label: 'ChatML (<|im_start|> / Qwen / Dolphin)',
        defaultNumCtx: 32768,
        template: `{{- if .Messages }}
{{- if .System }}<|im_start|>system
{{ .System }}<|im_end|>
{{ end }}
{{- range .Messages }}
{{- if ne .Role "system" }}<|im_start|>{{ .Role }}
{{ .Content }}<|im_end|>
{{ end }}
{{- end }}<|im_start|>assistant
{{ else }}
{{- if .System }}<|im_start|>system
{{ .System }}<|im_end|>
{{ end }}<|im_start|>user
{{ .Prompt }}<|im_end|>
<|im_start|>assistant
{{ end }}`,
        stopTokens: ['<|im_start|>', '<|im_end|>']
    },
    llama3: {
        id: 'llama3',
        label: 'Llama 3 / 3.1 / 3.2 (<|start_header_id|>)',
        defaultNumCtx: 32768,
        template: `{{- if .Messages }}
{{- if .System }}<|start_header_id|>system<|end_header_id|>

{{ .System }}<|eot_id|>{{ end }}
{{- range .Messages }}
{{- if ne .Role "system" }}<|start_header_id|>{{ .Role }}<|end_header_id|>

{{ .Content }}<|eot_id|>{{ end }}
{{- end }}<|start_header_id|>assistant<|end_header_id|>

{{ else }}
{{- if .System }}<|start_header_id|>system<|end_header_id|>

{{ .System }}<|eot_id|>{{ end }}<|start_header_id|>user<|end_header_id|>

{{ .Prompt }}<|eot_id|><|start_header_id|>assistant<|end_header_id|>

{{ end }}`,
        stopTokens: ['<|start_header_id|>', '<|end_header_id|>', '<|eot_id|>']
    },
    mistral: {
        id: 'mistral',
        label: 'Mistral / Mixtral ([INST])',
        defaultNumCtx: 32768,
        template: `{{- if .Messages }}
{{- if .System }}<s>[INST] {{ .System }} [/INST]{{ end }}
{{- range .Messages }}
{{- if eq .Role "user" }}<s>[INST] {{ .Content }} [/INST]
{{- else if eq .Role "assistant" }} {{ .Content }}</s>{{ end }}
{{- end }}
{{ else }}<s>[INST] {{ if .System }}{{ .System }}

{{ end }}{{ .Prompt }} [/INST]{{ end }}`,
        stopTokens: ['[INST]', '[/INST]', '</s>']
    },
    gemma2: {
        id: 'gemma2',
        label: 'Google Gemma 2 (<start_of_turn>)',
        defaultNumCtx: 8192,
        template: `{{- if .Messages }}
{{- if .System }}<start_of_turn>system
{{ .System }}<end_of_turn>
{{ end }}
{{- range .Messages }}
{{- if eq .Role "user" }}<start_of_turn>user
{{ .Content }}<end_of_turn>
{{ else if eq .Role "assistant" }}<start_of_turn>model
{{ .Content }}<end_of_turn>
{{ end }}
{{- end }}<start_of_turn>model
{{ else }}
{{- if .System }}<start_of_turn>user
{{ .System }}

{{ .Prompt }}<end_of_turn>
{{ else }}<start_of_turn>user
{{ .Prompt }}<end_of_turn>
{{ end }}<start_of_turn>model
{{ end }}`,
        stopTokens: ['<start_of_turn>', '<end_of_turn>']
    },
    deepseek: {
        id: 'deepseek',
        label: 'DeepSeek R1 / V2 / Coder',
        defaultNumCtx: 32768,
        template: `{{- if .Messages }}
{{- if .System }}{{ .System }}
{{ end }}
{{- range .Messages }}
{{- if eq .Role "user" }}<｜User｜>{{ .Content }}<｜Assistant｜>
{{- else if eq .Role "assistant" }}{{ .Content }}<｜end▁of▁sentence｜>{{ end }}
{{- end }}
{{ else }}{{ if .System }}{{ .System }}
{{ end }}<｜User｜>{{ .Prompt }}<｜Assistant｜>{{ end }}`,
        stopTokens: ['<｜User｜>', '<｜Assistant｜>', '<｜end▁of▁sentence｜>']
    },
    auto: {
        id: 'auto',
        label: 'Авто (встроенный шаблон из GGUF)',
        defaultNumCtx: 32768,
        template: null,
        stopTokens: []
    }
};

function sanitizeGgufFilename(rawName) {
    const base = path.basename(String(rawName || '').trim());
    if (!base || base === '.' || base === '..') return null;
    const cleaned = base.replace(/[^a-zA-Z0-9._-]/g, '_');
    if (!cleaned.toLowerCase().endsWith('.gguf')) return null;
    if (cleaned.startsWith('.')) return null;
    return cleaned;
}

function resolveSafeModelFilePath(rawFilename) {
    const safeName = sanitizeGgufFilename(rawFilename);
    if (!safeName) return null;
    const rootResolved = path.resolve(MODELS_DIR);
    const fullPath = path.resolve(rootResolved, safeName);
    if (!fullPath.startsWith(rootResolved + path.sep)) return null;
    return { filename: safeName, fullPath };
}

function sanitizeOllamaModelName(rawName, fallbackFilename = '') {
    let val = String(rawName || '').trim().toLowerCase();
    if (!val && fallbackFilename) {
        return suggestOllamaModelName(fallbackFilename);
    }
    val = val.replace(/\s+/g, '-').replace(/[^a-z0-9._:/-]/g, '');
    // Preserve optional single :tag suffix
    const colonIdx = val.indexOf(':');
    let base = colonIdx >= 0 ? val.slice(0, colonIdx) : val;
    let tag = colonIdx >= 0 ? val.slice(colonIdx + 1).replace(/:/g, '-') : 'latest';
    base = base.replace(/^[^a-z0-9]+|[^a-z0-9]+$/g, '').slice(0, 64) || 'custom-gguf';
    tag = tag.replace(/^[^a-z0-9]+|[^a-z0-9]+$/g, '').slice(0, 32) || 'latest';
    return `${base}:${tag}`;
}

function suggestOllamaModelName(filename) {
    let stem = path.basename(String(filename || 'model.gguf')).replace(/\.gguf$/i, '');
    let quant = 'latest';
    const qMatch = stem.match(/[-_.]((?:IQ|Q|FP|BF|F)\d+(?:_[A-Z0-9_]+)?)$/i);
    if (qMatch) {
        quant = qMatch[1].toLowerCase().replace(/[^a-z0-9._-]/g, '-');
        stem = stem.slice(0, qMatch.index);
    }
    const cleanBase = stem
        .toLowerCase()
        .replace(/[^a-z0-9._-]+/g, '-')
        .replace(/^-+|-+$/g, '')
        .slice(0, 64) || 'custom-gguf';
    return `${cleanBase}:${quant}`;
}

/**
 * Lightweight streaming GGUF v2/v3 metadata header reader (never loads tensor weights into RAM).
 */
function readGgufMetadata(filePath) {
    const result = {
        validMagic: false,
        version: null,
        architecture: '',
        modelName: '',
        chatTemplate: '',
        contextLength: null,
        fileType: null
    };
    const FIXED_SIZE_MAP = { 0: 1, 1: 1, 2: 2, 3: 2, 4: 4, 5: 4, 6: 4, 7: 1, 10: 8, 11: 8, 12: 8 };
    const MAX_SCAN_OFFSET = 16 * 1024 * 1024;
    let fd = null;

    try {
        const stat = fs.statSync(filePath);
        if (stat.size < 16) return result;
        fd = fs.openSync(filePath, 'r');

        const CHUNK_SIZE = 65536;
        const buf = Buffer.allocUnsafe(CHUNK_SIZE);
        let bufStart = 0;
        let bufLen = 0;
        let pos = 0;

        const ensure = (needBytes) => {
            if (pos + needBytes > stat.size || pos + needBytes > MAX_SCAN_OFFSET) return false;
            if (pos >= bufStart && pos + needBytes <= bufStart + bufLen) return true;
            if (needBytes > CHUNK_SIZE) return false;
            bufStart = pos;
            bufLen = fs.readSync(fd, buf, 0, Math.min(CHUNK_SIZE, stat.size - pos), pos);
            return bufLen >= needBytes;
        };

        const readU32 = () => {
            if (!ensure(4)) throw new Error('EOF');
            const v = buf.readUInt32LE(pos - bufStart);
            pos += 4;
            return v;
        };

        const readU64Num = () => {
            if (!ensure(8)) throw new Error('EOF');
            const big = buf.readBigUInt64LE(pos - bufStart);
            pos += 8;
            return Number(big);
        };

        const readMagic = () => {
            if (!ensure(4)) return '';
            const m = buf.toString('ascii', pos - bufStart, pos - bufStart + 4);
            pos += 4;
            return m;
        };

        const readStr = (maxKeepLen = 8192) => {
            const len = readU64Num();
            if (len < 0 || pos + len > stat.size) throw new Error('BAD_LEN');
            if (len === 0) return '';
            if (len <= maxKeepLen && len <= CHUNK_SIZE && ensure(len)) {
                const s = buf.toString('utf8', pos - bufStart, pos - bufStart + len);
                pos += len;
                return s;
            }
            const previewLen = Math.min(len, maxKeepLen);
            let out = '';
            if (previewLen > 0 && ensure(previewLen)) {
                out = buf.toString('utf8', pos - bufStart, pos - bufStart + previewLen);
            }
            pos += len;
            return out;
        };

        const readScalarNumber = (vtype) => {
            const sz = FIXED_SIZE_MAP[vtype];
            if (!sz || !ensure(sz)) throw new Error('EOF');
            const off = pos - bufStart;
            let val = null;
            if (vtype === 0) val = buf.readUInt8(off);
            else if (vtype === 1) val = buf.readInt8(off);
            else if (vtype === 2) val = buf.readUInt16LE(off);
            else if (vtype === 3) val = buf.readInt16LE(off);
            else if (vtype === 4) val = buf.readUInt32LE(off);
            else if (vtype === 5) val = buf.readInt32LE(off);
            else if (vtype === 7) val = buf.readUInt8(off) !== 0;
            else if (vtype === 10) val = Number(buf.readBigUInt64LE(off));
            else if (vtype === 11) val = Number(buf.readBigInt64LE(off));
            pos += sz;
            return val;
        };

        const skipValue = (vtype, depth = 0) => {
            if (depth > 4) throw new Error('DEPTH');
            if (FIXED_SIZE_MAP[vtype]) {
                pos += FIXED_SIZE_MAP[vtype];
                return;
            }
            if (vtype === 8) {
                const len = readU64Num();
                pos += len;
                return;
            }
            if (vtype === 9) {
                const elemType = readU32();
                const count = readU64Num();
                if (FIXED_SIZE_MAP[elemType]) {
                    pos += count * FIXED_SIZE_MAP[elemType];
                    return;
                }
                const maxItems = Math.min(count, 262144);
                for (let i = 0; i < maxItems; i++) {
                    skipValue(elemType, depth + 1);
                    if (pos > MAX_SCAN_OFFSET || pos >= stat.size) return;
                }
                return;
            }
            throw new Error('UNKNOWN_TYPE');
        };

        const magic = readMagic();
        if (magic !== 'GGUF') return result;
        result.validMagic = true;

        const version = readU32();
        result.version = version;
        if (version < 2 || version > 3) return result;

        readU64Num(); // tensor_count
        const kvCount = Math.min(readU64Num(), 1024);

        for (let i = 0; i < kvCount; i++) {
            if (pos >= MAX_SCAN_OFFSET || pos >= stat.size) break;
            const key = readStr(256);
            const vtype = readU32();

            if (key === 'general.architecture' && vtype === 8) {
                result.architecture = readStr(128).trim();
            } else if (key === 'general.name' && vtype === 8) {
                result.modelName = readStr(256).trim();
            } else if (key === 'tokenizer.chat_template' && vtype === 8) {
                result.chatTemplate = readStr(8192);
            } else if (key === 'general.file_type' && FIXED_SIZE_MAP[vtype]) {
                result.fileType = readScalarNumber(vtype);
            } else if (key.endsWith('.context_length') && FIXED_SIZE_MAP[vtype]) {
                result.contextLength = readScalarNumber(vtype);
            } else {
                skipValue(vtype, 0);
            }

            if (result.architecture && result.modelName && result.chatTemplate && result.contextLength) {
                break;
            }
        }
    } catch (_) {
        // Ignore truncated/non-standard tails and return collected header metadata
    } finally {
        if (fd !== null) {
            try { fs.closeSync(fd); } catch (_) {}
        }
    }
    return result;
}

function detectTemplateFamily(filename, ggufMeta = null) {
    const tmpl = (ggufMeta && ggufMeta.chatTemplate ? ggufMeta.chatTemplate : '');
    const arch = (ggufMeta && ggufMeta.architecture ? ggufMeta.architecture : '').toLowerCase();
    const fname = String(filename || '').toLowerCase();

    // 1. Authoritative check from embedded GGUF tokenizer.chat_template
    if (tmpl.includes('<|start_header_id|>') || tmpl.includes('<|eot_id|>')) return 'llama3';
    if (tmpl.includes('<|im_start|>') || tmpl.includes('<|im_end|>')) return 'chatml';
    if (tmpl.includes('[INST]')) return 'mistral';
    if (tmpl.includes('<start_of_turn>')) return 'gemma2';
    if (tmpl.includes('<｜User｜>') || tmpl.includes('<|User|>') || tmpl.includes('<｜Assistant｜>')) return 'deepseek';

    // 2. Fallback to GGUF architecture + filename hints
    if (fname.includes('dolphin-mistral') || (fname.includes('dolphin') && fname.includes('2.9'))) return 'chatml';
    if (fname.includes('dolphin-llama3') || fname.includes('dolphin3') || fname.includes('llama-3') || fname.includes('llama3')) return 'llama3';
    if (arch.includes('qwen') || fname.includes('qwen') || fname.includes('dolphin') || fname.includes('hermes') || fname.includes('smollm') || fname.includes('yi')) return 'chatml';
    if (arch.includes('gemma') || fname.includes('gemma')) return 'gemma2';
    if (arch.includes('deepseek') || fname.includes('deepseek')) return 'deepseek';
    if (fname.includes('mistral') || fname.includes('mixtral') || fname.includes('codestral')) return 'mistral';
    if (arch === 'llama') return 'llama3';
    return 'chatml';
}

function buildModelfileContent({
    ggufFilename = '',
    ggufAbsPath = '',
    templateFamily = 'chatml',
    numCtx = 32768,
    temperature = 0.7,
    systemPrompt = '',
    useRelativePath = true
}) {
    const familyKey = GGUF_TEMPLATES[templateFamily] ? templateFamily : 'chatml';
    const spec = GGUF_TEMPLATES[familyKey];
    const fromTarget = useRelativePath && ggufFilename
        ? `./${path.basename(ggufFilename)}`
        : (ggufAbsPath || ggufFilename || './model.gguf').replace(/\\/g, '/');

    const safeCtx = Math.max(2048, Math.min(131072, parseInt(numCtx, 10) || spec.defaultNumCtx || 32768));
    const parsedTemp = Number.isFinite(Number(temperature)) ? Math.max(0, Math.min(2, Number(temperature))) : 0.7;

    const lines = [
        `# Auto-generated by Qwen AI Studio (${spec.label})`,
        `FROM "${fromTarget}"`,
        ''
    ];

    if (spec.template) {
        lines.push(`TEMPLATE """${spec.template}"""`);
        lines.push('');
    }
    for (const stopTok of (spec.stopTokens || [])) {
        lines.push(`PARAMETER stop "${stopTok}"`);
    }
    lines.push(`PARAMETER num_ctx ${safeCtx}`);
    lines.push(`PARAMETER temperature ${parsedTemp}`);

    const cleanSys = String(systemPrompt || '').replace(/"""/g, "'''").trim();
    if (cleanSys) {
        lines.push('');
        lines.push(`SYSTEM """${cleanSys}"""`);
    }
    return lines.join('\n') + '\n';
}

function isPrivateOrLocalHostname(hostname) {
    const h = String(hostname || '').toLowerCase().replace(/^\[|\]$/g, '');
    if (!h || h === 'localhost' || h.endsWith('.local') || h.endsWith('.internal')) return true;
    if (h === '::1' || h.startsWith('fe80:') || h.startsWith('fc') || h.startsWith('fd')) return true;
    const ipv4 = h.match(/^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/);
    if (!ipv4) return !h.includes('.');
    const [a, b] = [Number(ipv4[1]), Number(ipv4[2])];
    if (a === 127 || a === 10 || a === 0) return true;
    if (a === 169 && b === 254) return true;
    if (a === 172 && b >= 16 && b <= 31) return true;
    if (a === 192 && b === 168) return true;
    return false;
}

function validateExternalModelUrl(rawUrl) {
    let parsed;
    try {
        parsed = new URL(String(rawUrl || '').trim());
    } catch (_) {
        throw new Error('Некорректный URL');
    }
    if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') {
        throw new Error('Разрешены только HTTP/HTTPS ссылки');
    }
    if (isPrivateOrLocalHostname(parsed.hostname)) {
        throw new Error('Локальные и внутренние адреса запрещены');
    }
    return parsed;
}

function parseHuggingFaceInput(rawInput) {
    const trimmed = String(rawInput || '').trim();
    if (!trimmed) throw new Error('Введите ссылку или имя репозитория Hugging Face');

    // Direct owner/repo shorthand (e.g. "Qwen/Qwen2.5-0.5B-Instruct-GGUF")
    const slugMatch = trimmed.match(/^([a-zA-Z0-9._-]+)\/([a-zA-Z0-9._-]+)$/);
    if (slugMatch) {
        return { kind: 'repo', owner: slugMatch[1], repo: slugMatch[2], branch: 'main' };
    }

    const parsed = validateExternalModelUrl(trimmed);
    const host = parsed.hostname.toLowerCase();
    const isHfHost = host === 'huggingface.co' || host.endsWith('.huggingface.co') || host === 'hf.co';

    // Convert HF /blob/branch/file.gguf -> /resolve/branch/file.gguf automatically
    if (isHfHost && parsed.pathname.includes('/blob/')) {
        parsed.pathname = parsed.pathname.replace('/blob/', '/resolve/');
    }

    const lowerPath = parsed.pathname.toLowerCase();
    if (lowerPath.endsWith('.gguf')) {
        const filename = sanitizeGgufFilename(decodeURIComponent(path.basename(parsed.pathname))) || 'model.gguf';
        return { kind: 'file', url: parsed.toString(), filename };
    }

    if (isHfHost) {
        const parts = parsed.pathname.split('/').filter(Boolean);
        if (parts.length >= 2) {
            const owner = parts[0];
            const repo = parts[1];
            const branch = (parts[2] === 'tree' && parts[3]) ? parts[3] : 'main';
            return { kind: 'repo', owner, repo, branch };
        }
    }

    throw new Error('Укажите прямую ссылку на .gguf или репозиторий Hugging Face (owner/repo)');
}

function requestWithRedirects(targetUrl, { method = 'GET', hfToken = '', redirectCount = 0 } = {}) {
    return new Promise((resolve, reject) => {
        if (redirectCount > 10) {
            reject(new Error('Слишком много перенаправлений HTTP'));
            return;
        }
        let parsed;
        try {
            parsed = validateExternalModelUrl(targetUrl);
        } catch (err) {
            reject(err);
            return;
        }

        const lib = parsed.protocol === 'https:' ? https : http;
        const headers = { 'User-Agent': 'QwenAIStudio/1.0' };
        const host = parsed.hostname.toLowerCase();
        // Send optional HF token strictly to Hugging Face domains, never to external CDN redirects
        if (hfToken && (host === 'huggingface.co' || host.endsWith('.huggingface.co') || host === 'hf.co')) {
            headers['Authorization'] = `Bearer ${hfToken.trim()}`;
        }

        const req = lib.request(parsed.toString(), { method, headers, timeout: 25000 }, (res) => {
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
                res.resume();
                const nextUrl = new URL(res.headers.location, parsed.toString()).toString();
                requestWithRedirects(nextUrl, { method, hfToken, redirectCount: redirectCount + 1 })
                    .then(resolve)
                    .catch(reject);
                return;
            }
            resolve({ res, finalUrl: parsed.toString(), req });
        });

        req.on('timeout', () => {
            req.destroy(new Error('Таймаут соединения при обращении к серверу'));
        });
        req.on('error', reject);
        req.end();
    });
}

async function inspectHuggingFaceInput(rawInput, hfToken = '') {
    const parsedInput = parseHuggingFaceInput(rawInput);
    const hwSpecs = await getSystemHardwareSpecs();

    if (parsedInput.kind === 'file') {
        let sizeBytes = 0;
        let statusNote = 'Готов к загрузке';
        try {
            const { res } = await requestWithRedirects(parsedInput.url, { method: 'HEAD', hfToken });
            res.resume();
            if (res.statusCode === 401 || res.statusCode === 403) {
                throw new Error('Модель требует авторизации (HF Token) на Hugging Face');
            }
            if (res.statusCode >= 400) {
                throw new Error(`Сервер вернул HTTP ${res.statusCode}`);
            }
            sizeBytes = parseInt(res.headers['content-length'] || '0', 10) || 0;
        } catch (headErr) {
            if (headErr.message.includes('HF Token') || headErr.message.includes('HTTP 4')) throw headErr;
            statusNote = 'Размер уточнится при старте загрузки';
        }
        const sizeGB = sizeBytes > 0 ? +(sizeBytes / (1024 ** 3)).toFixed(2) : 0;
        const detectedTemplate = detectTemplateFamily(parsedInput.filename, null);
        return {
            kind: 'file',
            file: {
                filename: parsedInput.filename,
                path: parsedInput.filename,
                url: parsedInput.url,
                sizeBytes,
                sizeGB,
                statusNote,
                suggestedOllamaName: suggestOllamaModelName(parsedInput.filename),
                detectedTemplate,
                compatibility: sizeGB > 0 ? computeHardwareCompatibility(sizeGB, hwSpecs) : null
            },
            hardware: hwSpecs
        };
    }

    const { owner, repo, branch } = parsedInput;
    const apiUrl = `https://huggingface.co/api/models/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/tree/${encodeURIComponent(branch)}?recursive=true`;
    const { res } = await requestWithRedirects(apiUrl, { method: 'GET', hfToken });
    const rawBody = await new Promise((resolve, reject) => {
        let data = '';
        res.on('data', chunk => {
            data += chunk;
            if (data.length > 8 * 1024 * 1024) res.destroy(new Error('Ответ Hugging Face API слишком большой'));
        });
        res.on('end', () => resolve(data));
        res.on('error', reject);
    });

    if (res.statusCode === 401 || res.statusCode === 403) {
        throw new Error('Закрытый репозиторий Hugging Face: укажите HF Token');
    }
    if (res.statusCode === 404) {
        throw new Error(`Репозиторий "${owner}/${repo}" или ветка "${branch}" не найдены`);
    }
    if (res.statusCode >= 400) {
        throw new Error(`Hugging Face API вернул HTTP ${res.statusCode}`);
    }

    const tree = JSON.parse(rawBody);
    if (!Array.isArray(tree)) throw new Error('Неожиданный формат ответа Hugging Face API');

    const files = tree
        .filter(item => item && item.type === 'file' && typeof item.path === 'string' && item.path.toLowerCase().endsWith('.gguf'))
        .map(item => {
            const filename = sanitizeGgufFilename(path.basename(item.path)) || path.basename(item.path);
            const sizeBytes = Number(item.size || (item.lfs && item.lfs.size) || 0);
            const sizeGB = +(sizeBytes / (1024 ** 3)).toFixed(2);
            const encodedPath = item.path.split('/').map(encodeURIComponent).join('/');
            return {
                filename,
                repoPath: item.path,
                url: `https://huggingface.co/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/resolve/${encodeURIComponent(branch)}/${encodedPath}?download=true`,
                sizeBytes,
                sizeGB,
                suggestedOllamaName: suggestOllamaModelName(filename),
                detectedTemplate: detectTemplateFamily(filename, null),
                compatibility: computeHardwareCompatibility(sizeGB, hwSpecs)
            };
        })
        .sort((a, b) => {
            const rank = f => {
                const n = f.filename.toLowerCase();
                if (n.includes('q4_k_m')) return 0;
                if (n.includes('q4_k_s') || n.includes('q4_0')) return 1;
                if (n.includes('q5_k_m')) return 2;
                return 3;
            };
            return rank(a) - rank(b) || a.sizeBytes - b.sizeBytes;
        });

    if (files.length === 0) {
        throw new Error(`В репозитории ${owner}/${repo} не найдено ни одного .gguf файла`);
    }

    return {
        kind: 'repo',
        repoId: `${owner}/${repo}`,
        branch,
        files,
        hardware: hwSpecs
    };
}

let hfDownloadState = {
    active: false,
    status: 'idle',
    filename: '',
    url: '',
    downloadedBytes: 0,
    totalBytes: 0,
    percent: 0,
    speedMBps: 0,
    error: '',
    startedAt: null,
    finishedAt: null
};
let activeHfRequest = null;
let activeHfResponse = null;
let activeHfWriteStream = null;
let activeHfPartPath = '';

function getPublicHfDownloadState() {
    return { ...hfDownloadState };
}

function cleanupPartFile(partPath) {
    if (partPath && fs.existsSync(partPath)) {
        try { fs.unlinkSync(partPath); } catch (_) {}
    }
}

function destroyWriteStreamAndCleanup(ws, partPath) {
    if (ws && !ws.destroyed) {
        ws.once('close', () => cleanupPartFile(partPath));
        try { ws.destroy(); } catch (_) { cleanupPartFile(partPath); }
    } else {
        cleanupPartFile(partPath);
    }
}

function cancelActiveHfDownload() {
    if (!hfDownloadState.active) return false;
    hfDownloadState.active = false;
    hfDownloadState.status = 'cancelled';
    hfDownloadState.error = 'Загрузка отменена пользователем';
    hfDownloadState.finishedAt = getLocalTimestamp();
    try { if (activeHfResponse) activeHfResponse.destroy(); } catch (_) {}
    try { if (activeHfRequest) activeHfRequest.destroy(); } catch (_) {}
    destroyWriteStreamAndCleanup(activeHfWriteStream, activeHfPartPath);
    activeHfRequest = null;
    activeHfResponse = null;
    activeHfWriteStream = null;
    activeHfPartPath = '';
    return true;
}

async function startHfDownloadJob(rawInput, customFilename = '', hfToken = '') {
    if (hfDownloadState.active) {
        throw new Error(`Уже идёт загрузка "${hfDownloadState.filename}". Дождитесь окончания или отмените её.`);
    }
    const parsedInput = parseHuggingFaceInput(rawInput);
    if (parsedInput.kind !== 'file') {
        throw new Error('Для скачивания выберите конкретный .gguf файл из репозитория');
    }

    const safeTarget = resolveSafeModelFilePath(customFilename || parsedInput.filename);
    if (!safeTarget) {
        throw new Error('Некорректное имя .gguf файла');
    }

    ensureDirExists(MODELS_DIR, 'models');
    const partPath = `${safeTarget.fullPath}.part_${Date.now()}`;
    activeHfPartPath = partPath;

    hfDownloadState = {
        active: true,
        status: 'connecting',
        filename: safeTarget.filename,
        url: parsedInput.url,
        downloadedBytes: 0,
        totalBytes: 0,
        percent: 0,
        speedMBps: 0,
        error: '',
        startedAt: getLocalTimestamp(),
        finishedAt: null
    };

    const failJob = (msg) => {
        if (hfDownloadState.status === 'cancelled') return;
        hfDownloadState.active = false;
        hfDownloadState.status = 'error';
        hfDownloadState.error = msg;
        hfDownloadState.finishedAt = getLocalTimestamp();
        destroyWriteStreamAndCleanup(activeHfWriteStream, partPath);
        activeHfRequest = null;
        activeHfResponse = null;
        activeHfWriteStream = null;
        activeHfPartPath = '';
        writeLog('ERROR', `[HF-Download] ${safeTarget.filename}: ${msg}`);
    };

    (async () => {
        try {
            const { res, req } = await requestWithRedirects(parsedInput.url, { method: 'GET', hfToken });
            activeHfRequest = req;
            activeHfResponse = res;

            if (res.statusCode === 401 || res.statusCode === 403) {
                res.resume();
                failJob('Доступ запрещён (HTTP ' + res.statusCode + ') — проверьте HF Token');
                return;
            }
            if (res.statusCode >= 400) {
                res.resume();
                failJob(`Сервер вернул HTTP ${res.statusCode}`);
                return;
            }

            hfDownloadState.status = 'downloading';
            hfDownloadState.totalBytes = parseInt(res.headers['content-length'] || '0', 10) || 0;

            const ws = fs.createWriteStream(partPath);
            activeHfWriteStream = ws;

            let magicChecked = false;
            let headChunk = Buffer.alloc(0);
            const startMs = Date.now();

            ws.on('drain', () => {
                if (hfDownloadState.active && !res.destroyed) res.resume();
            });

            res.on('data', (chunk) => {
                if (!hfDownloadState.active) return;
                let dataToWrite = chunk;
                if (!magicChecked) {
                    headChunk = Buffer.concat([headChunk, chunk]);
                    if (headChunk.length < 4) return;
                    magicChecked = true;
                    if (headChunk.toString('ascii', 0, 4) !== 'GGUF') {
                        res.destroy();
                        failJob('Скачиваемый файл не является GGUF (получен HTML или другой формат)');
                        return;
                    }
                    dataToWrite = headChunk;
                    headChunk = null;
                }
                hfDownloadState.downloadedBytes += dataToWrite.length;
                if (hfDownloadState.totalBytes > 0) {
                    hfDownloadState.percent = Math.min(99, Math.round((hfDownloadState.downloadedBytes / hfDownloadState.totalBytes) * 100));
                }
                const elapsedSec = Math.max(0.25, (Date.now() - startMs) / 1000);
                hfDownloadState.speedMBps = +((hfDownloadState.downloadedBytes / (1024 * 1024)) / elapsedSec).toFixed(2);
                if (!ws.write(dataToWrite)) {
                    res.pause();
                }
            });

            res.on('end', () => {
                if (!hfDownloadState.active) return;
                if (!magicChecked) {
                    failJob('Скачанный файл слишком мал или не является GGUF');
                    return;
                }
                ws.end();
            });

            ws.on('finish', () => {
                if (!hfDownloadState.active) return;
                try {
                    if (!magicChecked || fs.statSync(partPath).size < 16) {
                        failJob('Файл повреждён или слишком мал');
                        return;
                    }
                    fs.renameSync(partPath, safeTarget.fullPath);
                    hfDownloadState.active = false;
                    hfDownloadState.status = 'completed';
                    hfDownloadState.percent = 100;
                    hfDownloadState.finishedAt = getLocalTimestamp();
                    activeHfRequest = null;
                    activeHfResponse = null;
                    activeHfWriteStream = null;
                    activeHfPartPath = '';
                    writeLog('INFO', `[HF-Download] Модель ${safeTarget.filename} успешно скачана в models/`);
                } catch (finErr) {
                    failJob(finErr.message);
                }
            });

            res.on('error', (err) => failJob(err.message));
            ws.on('error', (err) => failJob(err.message));
        } catch (err) {
            failJob(err.message);
        }
    })();

    return getPublicHfDownloadState();
}

async function listLocalGgufModels() {
    ensureDirExists(MODELS_DIR, 'models');
    const [installedModels, hwSpecs] = await Promise.all([
        getLocalOllamaModels(),
        getSystemHardwareSpecs()
    ]);
    const installedNames = new Set((installedModels || []).map(m => String(m.name || '').toLowerCase()));

    const entries = fs.readdirSync(MODELS_DIR, { withFileTypes: true });
    const files = [];

    for (const entry of entries) {
        if (!entry.isFile() || !entry.name.toLowerCase().endsWith('.gguf')) continue;
        const safe = resolveSafeModelFilePath(entry.name);
        if (!safe) continue;
        const stat = fs.statSync(safe.fullPath);
        const sizeGB = +(stat.size / (1024 ** 3)).toFixed(2);
        const ggufMeta = readGgufMetadata(safe.fullPath);
        const detectedTemplate = detectTemplateFamily(safe.filename, ggufMeta);
        const suggestedOllamaName = suggestOllamaModelName(safe.filename);
        const numCtx = ggufMeta.contextLength && ggufMeta.contextLength >= 2048
            ? Math.min(ggufMeta.contextLength, detectedTemplate === 'gemma2' ? 8192 : 32768)
            : (GGUF_TEMPLATES[detectedTemplate]?.defaultNumCtx || 32768);

        files.push({
            filename: safe.filename,
            sizeBytes: stat.size,
            sizeGB,
            modifiedAt: stat.mtime.toISOString(),
            ggufMeta: {
                validMagic: ggufMeta.validMagic,
                version: ggufMeta.version,
                architecture: ggufMeta.architecture || '',
                modelName: ggufMeta.modelName || '',
                hasEmbeddedTemplate: Boolean(ggufMeta.chatTemplate),
                contextLength: ggufMeta.contextLength || null
            },
            detectedTemplate,
            detectedTemplateLabel: GGUF_TEMPLATES[detectedTemplate]?.label || detectedTemplate,
            suggestedOllamaName,
            recommendedNumCtx: numCtx,
            alreadyImported: installedNames.has(suggestedOllamaName.toLowerCase()),
            compatibility: computeHardwareCompatibility(sizeGB, hwSpecs),
            modelfilePreview: buildModelfileContent({
                ggufFilename: safe.filename,
                ggufAbsPath: safe.fullPath,
                templateFamily: detectedTemplate,
                numCtx,
                useRelativePath: true
            })
        });
    }

    files.sort((a, b) => new Date(b.modifiedAt) - new Date(a.modifiedAt));
    return {
        modelsDir: MODELS_DIR,
        files,
        templates: Object.values(GGUF_TEMPLATES).map(t => ({ id: t.id, label: t.label, defaultNumCtx: t.defaultNumCtx })),
        hardware: hwSpecs,
        hfDownload: getPublicHfDownloadState()
    };
}

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

    // 7. Web: Поиск в интернете (DuckDuckGo Search)
    if (req.url === '/api/web/search' && req.method === 'POST') {
        readJsonBody(req, async (err, body) => {
            if (err || !body || !body.query) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется поле query' }));
                return;
            }

            const query = String(body.query).trim();
            const maxResults = parseInt(body.maxResults) || 5;
            const deepFetch = body.deepFetch !== false;

            writeLog('INFO', `[WebSearch] Запрос: "${query.substring(0, 50)}" (глубокое чтение: ${deepFetch ? 'да' : 'нет'})...`);

            try {
                const results = await performWebSearch(query, maxResults, deepFetch);
                writeLog('INFO', `[WebSearch] Успешно найдено ${results.length} результатов для "${query.substring(0, 40)}"`);
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({
                    status: 'ok',
                    query: query,
                    results: results
                }));
            } catch (searchErr) {
                writeLog('ERROR', `[WebSearch] Ошибка поиска: ${searchErr.message}`);
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: searchErr.message, results: [] }));
            }
        });
        return;
    }

    // 8. Web: Прямое чтение страницы по URL
    if (req.url === '/api/web/fetch' && req.method === 'POST') {
        readJsonBody(req, async (err, body) => {
            if (err || !body || !body.url) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется поле url' }));
                return;
            }

            const targetUrl = String(body.url).trim();
            const maxChars = parseInt(body.maxChars) || 3000;

            writeLog('INFO', `[WebFetch] Загрузка текста страницы: "${targetUrl}"...`);

            try {
                const text = await fetchPageContent(targetUrl, maxChars);
                writeLog('INFO', `[WebFetch] Извлечено ${text.length} символов с "${targetUrl}"`);
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({
                    status: 'ok',
                    url: targetUrl,
                    text: text
                }));
            } catch (fetchErr) {
                writeLog('ERROR', `[WebFetch] Ошибка загрузки страницы: ${fetchErr.message}`);
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: fetchErr.message }));
            }
        });
        return;
    }

    // 8.1 Agent: Автономный ReAct-агент глубокого исследования (Deep Research Loop)
    if (req.url === '/api/agent/research' && req.method === 'POST') {
        readJsonBody(req, async (err, body) => {
            if (err || !body || !body.query) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется поле query' }));
                return;
            }

            let chatId = body.chat_id || null;
            const model = body.model || 'qwen3.5:9b';
            const query = String(body.query).trim();
            const requestedPersonaId = body.persona_id || null;
            const requestedPersonaPrompt = body.persona_prompt ? String(body.persona_prompt).trim() : '';
            const requestedSteps = Math.min(Math.max(parseInt(body.max_steps) || 4, 2), 6);
            const maxSteps = extractLocalPathFromQuery(query) ? 6 : requestedSteps;
            const today = new Date().toISOString().split('T')[0];

            let currentOllamaReq = null;
            let isAborted = false;

            res.on('close', () => {
                if (!res.writableEnded) {
                    isAborted = true;
                    writeLog('WARN', `[Agent] Клиент разорвал соединение для "${query.substring(0, 30)}", прерываем Ollama...`);
                    if (currentOllamaReq) {
                        try { currentOllamaReq.destroy(); } catch(e){}
                    }
                }
            });

            if (!chatId || !db.getChatById(chatId)) {
                const newChat = db.createChat('Новый диалог', model);
                chatId = newChat.id;
            }

            try {
                db.addMessage(chatId, 'user', query);
            } catch (dbErr) {
                writeLog('WARN', `[Agent] Ошибка сохранения сообщения пользователя: ${dbErr.message}`);
            }

            res.writeHead(200, {
                'Content-Type': 'application/x-ndjson; charset=utf-8',
                'Transfer-Encoding': 'chunked',
                'Cache-Control': 'no-cache, no-transform',
                'Connection': 'keep-alive'
            });

            const sendEvent = (obj) => {
                if (isAborted || res.writableEnded) return;
                try {
                    res.write(JSON.stringify(obj) + '\n');
                } catch (writeErr) {}
            };

            sendEvent({
                type: 'init',
                chatId: chatId,
                model: model,
                query: query,
                maxSteps: maxSteps
            });

            writeLog('INFO', `[Agent] Запуск Deep Research [${model}] для "${query.substring(0, 50)}" (макс. шагов: ${maxSteps})...`);

            const systemPrompt = `Ты — автономный исследовательский AI-агент ReAct (Reason + Act) внутри приложения Qwen AI Studio, запущенного локально на ПК пользователя.
Твоя цель — глубоко изучить вопрос пользователя (сведения из интернета и/или файлы проекта на компьютере) и подготовить материал для исчерпывающего аналитического отчета.
Сегодняшняя дата: ${today}.

ТЕБЕ ДОСТУПНЫ СЛЕДУЮЩИЕ ДЕЙСТВИЯ:
1. Action: search("поисковый запрос") — поиск в интернете (DuckDuckGo): заголовки, выдержки и ссылки.
2. Action: fetch("URL") — загружает и очищает текст конкретной веб-страницы.
3. Action: list_dir("C:/путь/к/папке") — показывает содержимое папки на ПК пользователя (только чтение).
4. Action: read_file("C:/путь/к/файлу.txt") — читает текстовый файл (README, исходный код, конфиги). Только чтение.
5. Action: finish — завершает сбор информации, когда собрано достаточно фактов.

Разрешённые папки для list_dir и read_file: ${AGENT_FS_ROOTS.join('; ')}. Остальные закрыты.

ПРАВИЛА И СТРОГИЙ ФОРМАТ КАЖДОГО ОТВЕТА:
- Ты должен выводить СТРОГО:
Thought: [Твои рассуждения: что уже известно, чего не хватает, какой шаг сделать сейчас]
Action: search("...") ИЛИ fetch("...") ИЛИ list_dir("...") ИЛИ read_file("...") ИЛИ finish
- За один шаг вызывай строго ОДНО действие. Параметр пиши полностью на одной строке.
- Для исследования проекта на ПК: сначала list_dir корня, затем прочитай README/документацию, затем главные файлы (package.json, точки входа, конфиги). Не читай node_modules, .git и бинарные файлы.
- Для вопросов про интернет используй search, затем fetch по самой полезной ссылке.
- Не выдумывай факты и содержимое файлов. Используй только то, что получил в Observation.
- Если данных уже достаточно, не затягивай — пиши Action: finish.
- Все мысли (Thought) пиши на русском языке.`;

            const agentMessages = [
                { role: 'system', content: systemPrompt },
                { role: 'user', content: `Вопрос для исследования: "${query}"\nНачни исследование с шага 1 (Thought + Action):` }
            ];

            const collectedSources = [];
            const seenUrls = new Set();
            let researchTrace = `Исследование вопроса: "${query}"\n`;
            let totalPromptTokens = 0;
            let totalEvalTokens = 0;

            // Файловый режим: в запросе пользователя указан локальный путь
            const localPathHint = extractLocalPathFromQuery(query);
            const fileMode = !!localPathHint;
            const stepCtx = fileMode ? 12288 : 8192;

            // Выполнение файловых действий (list_dir / read_file). Возвращает текст наблюдения и событие для UI
            const runFsAction = (act, param) => {
                const res = act === 'list_dir' ? agentListDir(param) : agentReadFile(param, 4000);
                if (!res.ok) {
                    return {
                        observationText: `Ошибка: ${res.error}`,
                        event: { type: 'observation', action: act, path: param, error: res.error }
                    };
                }
                const srcId = collectedSources.length + 1;
                collectedSources.push({
                    id: srcId,
                    kind: 'file',
                    title: (act === 'list_dir' ? 'Папка: ' : 'Файл: ') + res.path,
                    url: '',
                    snippet: res.text.substring(0, 250),
                    pageContent: res.text
                });
                if (act === 'list_dir') {
                    return {
                        observationText: `Содержимое папки ${res.path} (${res.count} элементов):\n${res.text}`,
                        event: { type: 'observation', action: act, path: res.path, entries: res.count }
                    };
                }
                return {
                    observationText: `Содержимое файла ${res.path}${res.truncated ? ' (обрезано)' : ''}:\n${res.text}`,
                    event: { type: 'observation', action: act, path: res.path, contentLength: res.chars }
                };
            };

            // Детерминированный шаг 1: сразу показываем модели то, что лежит по указанному пути
            let startStep = 1;
            if (fileMode) {
                const isDir = (() => { try { return fs.statSync(localPathHint).isDirectory(); } catch (e) { return false; } })();
                const preAction = isDir ? 'list_dir' : 'read_file';
                const preThought = `Пользователь указал локальный путь «${localPathHint}». Сначала изучу ${isDir ? 'структуру папки' : 'файл'}.`;
                writeLog('INFO', `[Agent] Файловый режим: ${preAction}("${localPathHint}")`);
                sendEvent({ type: 'step', step: 1, totalSteps: maxSteps, thought: preThought, action: preAction, actionParam: localPathHint });
                const pre = runFsAction(preAction, localPathHint);
                sendEvent({ ...pre.event, step: 1 });
                researchTrace += `\n[Шаг 1]\n💭 Мысль: ${preThought}\n⚙️ Действие: ${preAction} ("${localPathHint}")\n📄 Наблюдение: ${pre.observationText.substring(0, 300)}...\n`;
                agentMessages.push({ role: 'assistant', content: `Thought: ${preThought}\nAction: ${preAction}("${localPathHint}")` });
                agentMessages.push({ role: 'user', content: `Observation: ${pre.observationText}\nСледующий шаг:` });
                startStep = 2;
            }

            for (let step = startStep; step <= maxSteps; step++) {
                if (isAborted) break;
                writeLog('INFO', `[Agent] Шаг ${step}/${maxSteps} рассуждения модели...`);

                let stepResponseText = '';
                try {
                    const ollamaRes = await callOllamaChat(model, agentMessages, {
                        temperature: 0.2,
                        num_ctx: stepCtx,

                        stop: ['\nObservation:', 'Observation:']
                    }, (r) => { currentOllamaReq = r; });

                    if (ollamaRes.prompt_eval_count) totalPromptTokens += ollamaRes.prompt_eval_count;
                    if (ollamaRes.eval_count) totalEvalTokens += ollamaRes.eval_count;

                    stepResponseText = (ollamaRes.message && ollamaRes.message.content) ? ollamaRes.message.content.trim() : '';
                } catch (modelErr) {
                    writeLog('ERROR', `[Agent] Ошибка на шаге ${step}: ${modelErr.message}`);
                    sendEvent({ type: 'step_error', step: step, error: modelErr.message });
                    break;
                }

                const parsed = parseReActOutput(stepResponseText);
                const thought = parsed.thought || 'Анализирую собранные сведения...';
                const action = parsed.action || 'finish';
                const actionParam = parsed.actionParam || '';

                researchTrace += `\n[Шаг ${step}]\n💭 Мысль: ${thought}\n⚙️ Действие: ${action}${actionParam ? ` ("${actionParam}")` : ''}\n`;

                sendEvent({
                    type: 'step',
                    step: step,
                    totalSteps: maxSteps,
                    thought: thought,
                    action: action,
                    actionParam: actionParam
                });

                if (action === 'finish') {
                    writeLog('INFO', `[Agent] Агент решил завершить сбор данных на шаге ${step}.`);
                    break;
                }

                let observationText = '';
                if (action === 'search') {
                    const searchQuery = actionParam || query;
                    writeLog('INFO', `[Agent] Шаг ${step}: Поиск в сети: "${searchQuery}"...`);
                    try {
                        const searchResults = await performWebSearch(searchQuery, 4, false);
                        if (searchResults && searchResults.length > 0) {
                            const newSourcesForStep = [];
                            const obsLines = [];
                            for (const item of searchResults) {
                                if (!seenUrls.has(item.url)) {
                                    seenUrls.add(item.url);
                                    const sourceIdx = collectedSources.length + 1;
                                    const srcObj = {
                                        id: sourceIdx,
                                        title: item.title,
                                        url: item.url,
                                        snippet: item.snippet
                                    };
                                    collectedSources.push(srcObj);
                                    newSourcesForStep.push(srcObj);
                                }
                                obsLines.push(`- [${item.title}] (${item.url}): ${item.snippet}`);
                            }
                            observationText = `Результаты поиска (${searchResults.length}):\n` + obsLines.join('\n');
                            sendEvent({
                                type: 'observation',
                                step: step,
                                action: 'search',
                                resultsCount: searchResults.length,
                                sources: newSourcesForStep
                            });
                        } else {
                            observationText = 'По данному запросу результатов не найдено.';
                            sendEvent({
                                type: 'observation',
                                step: step,
                                action: 'search',
                                resultsCount: 0,
                                sources: []
                            });
                        }
                    } catch (searchErr) {
                        observationText = `Ошибка поиска: ${searchErr.message}`;
                        sendEvent({ type: 'observation', step: step, action: 'search', error: searchErr.message });
                    }
                } else if (action === 'fetch') {
                    const targetUrl = actionParam;
                    writeLog('INFO', `[Agent] Шаг ${step}: Чтение страницы: "${targetUrl}"...`);
                    try {
                        const pageText = await fetchPageContent(targetUrl, 1800);
                        if (pageText && pageText.length > 50) {
                            observationText = `Текст со страницы ${targetUrl} (извлечено ${pageText.length} символов):\n${pageText}`;
                            const existingSrc = collectedSources.find(s => s.url === targetUrl);
                            if (existingSrc) {
                                existingSrc.pageContent = pageText;
                            } else {
                                const newSrc = {
                                    id: collectedSources.length + 1,
                                    title: targetUrl,
                                    url: targetUrl,
                                    snippet: pageText.substring(0, 250) + '...',
                                    pageContent: pageText
                                };
                                collectedSources.push(newSrc);
                            }
                            sendEvent({
                                type: 'observation',
                                step: step,
                                action: 'fetch',
                                url: targetUrl,
                                contentLength: pageText.length
                            });
                        } else {
                            observationText = `Не удалось извлечь читаемый текст с ${targetUrl}.`;
                            sendEvent({ type: 'observation', step: step, action: 'fetch', url: targetUrl, contentLength: 0 });
                        }
                    } catch (fetchErr) {
                        observationText = `Ошибка загрузки страницы: ${fetchErr.message}`;
                        sendEvent({ type: 'observation', step: step, action: 'fetch', error: fetchErr.message });
                    }
                } else if (action === 'list_dir' || action === 'read_file') {
                    writeLog('INFO', `[Agent] Шаг ${step}: ${action}("${actionParam}")`);
                    const fsRes = runFsAction(action, actionParam);
                    observationText = fsRes.observationText;
                    sendEvent({ ...fsRes.event, step: step });
                } else {
                    observationText = `Неизвестное действие "${action}". Используй search("запрос"), fetch("URL"), list_dir("путь"), read_file("путь") или finish.`;
                }

                researchTrace += `📄 Наблюдение: ${observationText.substring(0, 300)}...\n`;

                agentMessages.push({ role: 'assistant', content: stepResponseText });
                agentMessages.push({ role: 'user', content: `Observation: ${observationText}\nСледующий шаг:` });
            }

            if (isAborted) {
                try { res.end(); } catch(e){}
                return;
            }

            writeLog('INFO', `[Agent] Начало финального синтеза отчета по "${query.substring(0, 40)}"...`);
            sendEvent({ type: 'final_start', totalSourcesFound: collectedSources.length });

            const materialsText = collectedSources.map((s, idx) => {
                const content = s.pageContent || s.snippet || '';
                const loc = s.kind === 'file' ? '' : ` (${s.url})`;
                return `[${idx + 1}] "${s.title}"${loc}\nФрагмент:\n${content}`;
            }).join('\n\n');

            const noMaterialsNote = fileMode
                ? 'Файлы по указанному пути прочитать не удалось. Честно сообщи об этом пользователю и НЕ выдумывай содержимое проекта.'
                : 'Прямых ссылок не найдено, используй свои экспертные знания.';

            let activePersonaNote = '';
            try {
                const resolvedPersonaText = requestedPersonaPrompt || (() => {
                    const chatRow = chatId ? db.getChatById(chatId) : null;
                    const pid = requestedPersonaId || (chatRow && chatRow.persona_id) || db.getSetting('active_persona_id', 'default');
                    const p = pid && pid !== 'default' ? db.getPersonaById(pid) : null;
                    return p ? `${p.name}: ${p.system_prompt}` : '';
                })();
                if (resolvedPersonaText) {
                    activePersonaNote = `\nУЧТИ СТИЛЬ И РОЛЬ ВЫБРАННОЙ ПЕРСОНЫ:\n${resolvedPersonaText}\n`;
                }
            } catch (e) {}

            const finalPrompt = `Ты — ведущий эксперт-аналитик и технический исследователь.
Твоя задача — составить подробный, исчерпывающий и структурированный аналитический отчет по запросу: "${query}".
${activePersonaNote}
СОБРАННЫЕ В ХОДЕ АВТОНОМНОГО ИССЛЕДОВАНИЯ МАТЕРИАЛЫ (Дата: ${today}):
${materialsText || noMaterialsNote}

ОБЯЗАТЕЛЬНЫЕ ТРЕБОВАНИЯ К ОТЧЕТУ:
1. Пиши на русском языке с четкой структурой (краткое резюме / суть, подробный разбор по пунктам, сравнительные аспекты, выводы и рекомендации).
2. Обязательно делай ссылки на факты из материалов выше с помощью сносок [1], [2] и т.д.
3. Не лей воду: давай точные цифры, технические характеристики, даты и аргументы.
4. Если изучались файлы или папки проекта: опиши назначение проекта, структуру, ключевые файлы и технологии, но ТОЛЬКО на основе реально прочитанного содержимого. Если что-то не прочитано — так и скажи.
5. Ответ должен быть полным, чтобы пользователю не требовалось искать дополнительные данные.`;


            const finalMessages = [
                { role: 'system', content: finalPrompt },
                { role: 'user', content: `Подготовь итоговый детальный отчет по вопросу: "${query}"` }
            ];

            let fullFinalAnswer = '';
            let finalPromptTokens = 0;
            let finalEvalTokens = 0;

            streamOllamaChat(
                model,
                finalMessages,
                (chunk) => {
                    if (isAborted) return;
                    if (chunk.message && chunk.message.content) {
                        fullFinalAnswer += chunk.message.content;
                        sendEvent({ type: 'token', content: chunk.message.content });
                    }
                    if (chunk.prompt_eval_count) finalPromptTokens = chunk.prompt_eval_count;
                    if (chunk.eval_count) finalEvalTokens = chunk.eval_count;
                },
                () => {
                    if (isAborted) return;
                    totalPromptTokens += finalPromptTokens;
                    totalEvalTokens += finalEvalTokens;

                    try {
                        db.addMessage(chatId, 'assistant', fullFinalAnswer, researchTrace, totalPromptTokens, totalEvalTokens);

                        const curChat = db.getChatById(chatId);
                        if (curChat && curChat.title === 'Новый диалог') {
                            let cleanTitle = query.replace(/[#*`\n]/g, ' ').trim();
                            if (cleanTitle.length > 36) cleanTitle = cleanTitle.substring(0, 36).trim() + '...';
                            if (cleanTitle) {
                                db.updateChat(chatId, { title: cleanTitle });
                                writeLog('INFO', `[Agent] Чат ${chatId} переименован в "${cleanTitle}"`);
                            }
                        }
                    } catch (dbErr) {
                        writeLog('WARN', `[Agent] Ошибка сохранения ответа агента: ${dbErr.message}`);
                    }

                    writeLog('INFO', `[Agent] Исследование завершено (${fullFinalAnswer.length} симв., ${collectedSources.length} источников).`);
                    sendEvent({
                        type: 'final_done',
                        fullContent: fullFinalAnswer,
                        allSources: collectedSources,
                        promptTokens: totalPromptTokens,
                        evalTokens: totalEvalTokens
                    });
                    try { res.end(); } catch(e){}
                },
                (streamErr) => {
                    if (isAborted) return;
                    writeLog('ERROR', `[Agent] Сбой финального стриминга: ${streamErr.message}`);
                    sendEvent({ type: 'error', message: streamErr.message });
                    try { res.end(); } catch(e){}
                },
                { num_ctx: 16384 },
                (r) => { currentOllamaReq = r; }
            );
        });
        return;
    }

    // 8.9 Stage 3: Локальные GGUF модели (models/), генерация Modelfile, ollama create и Hugging Face
    const reqPathname = (req.url || '/').split('?')[0];

    if (reqPathname === '/api/models/gguf' && req.method === 'GET') {
        listLocalGgufModels()
            .then(data => {
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', ...data }));
            })
            .catch(err => {
                writeLog('ERROR', `[GGUF] Ошибка сканирования папки models/: ${err.message}`);
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: err.message, files: [] }));
            });
        return;
    }

    if (reqPathname === '/api/models/gguf/upload' && req.method === 'POST') {
        let rawName = '';
        try {
            const parsedReqUrl = new URL(req.url, 'http://127.0.0.1');
            rawName = parsedReqUrl.searchParams.get('filename') || '';
        } catch (_) {}
        if (!rawName && req.headers['x-filename']) {
            try { rawName = decodeURIComponent(String(req.headers['x-filename'])); } catch (_) { rawName = String(req.headers['x-filename']); }
        }
        const safeTarget = resolveSafeModelFilePath(rawName);
        if (!safeTarget) {
            req.resume();
            res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: 'Разрешена загрузка только файлов с расширением .gguf и безопасным именем.' }));
            return;
        }

        ensureDirExists(MODELS_DIR, 'models');
        const partPath = `${safeTarget.fullPath}.part_${Date.now()}`;
        const ws = fs.createWriteStream(partPath);
        let magicChecked = false;
        let headBuf = Buffer.alloc(0);
        let aborted = false;
        let receivedBytes = 0;

        const failUpload = (code, msg) => {
            if (aborted) return;
            aborted = true;
            destroyWriteStreamAndCleanup(ws, partPath);
            if (!res.headersSent) {
                res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Connection': 'close' });
                res.end(JSON.stringify({ error: msg }));
            }
            setTimeout(() => { try { if (!req.destroyed) req.destroy(); } catch (_) {} }, 50);
        };

        req.on('data', (chunk) => {
            if (aborted) return;
            receivedBytes += chunk.length;
            let dataToWrite = chunk;
            if (!magicChecked) {
                headBuf = Buffer.concat([headBuf, chunk]);
                if (headBuf.length < 4) return;
                magicChecked = true;
                if (headBuf.toString('ascii', 0, 4) !== 'GGUF') {
                    failUpload(400, 'Файл не содержит заголовок GGUF (проверьте, что это настоящая .gguf модель).');
                    return;
                }
                dataToWrite = headBuf;
                headBuf = null;
            }
            if (!ws.write(dataToWrite)) {
                req.pause();
            }
        });

        ws.on('drain', () => {
            if (!aborted) req.resume();
        });

        req.on('aborted', () => failUpload(499, 'Загрузка файла прервана'));
        req.on('error', (err) => failUpload(500, err.message));
        ws.on('error', (err) => failUpload(500, err.message));

        req.on('end', () => {
            if (aborted) return;
            if (!magicChecked || receivedBytes < 16) {
                failUpload(400, 'Файл слишком мал или не является валидным GGUF');
                return;
            }
            ws.end(() => {
                if (aborted) return;
                try {
                    fs.renameSync(partPath, safeTarget.fullPath);
                    const meta = readGgufMetadata(safeTarget.fullPath);
                    const detectedTemplate = detectTemplateFamily(safeTarget.filename, meta);
                    writeLog('INFO', `[GGUF] Загружен файл ${safeTarget.filename} (${(receivedBytes / (1024 ** 2)).toFixed(1)} MB, template=${detectedTemplate})`);
                    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                    res.end(JSON.stringify({
                        status: 'ok',
                        filename: safeTarget.filename,
                        sizeBytes: receivedBytes,
                        detectedTemplate,
                        ggufMeta: meta
                    }));
                } catch (finErr) {
                    failUpload(500, finErr.message);
                }
            });
        });
        return;
    }

    if (reqPathname.startsWith('/api/models/gguf/') && req.method === 'DELETE') {
        let rawName = '';
        try { rawName = decodeURIComponent(reqPathname.slice('/api/models/gguf/'.length)); } catch (_) {}
        const safeTarget = resolveSafeModelFilePath(rawName);
        if (!safeTarget) {
            res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: 'Некорректное имя .gguf файла' }));
            return;
        }
        if (!fs.existsSync(safeTarget.fullPath)) {
            res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: 'Файл не найден в папке models/' }));
            return;
        }
        try {
            fs.unlinkSync(safeTarget.fullPath);
            const mfPath = safeTarget.fullPath.replace(/\.gguf$/i, '.Modelfile');
            if (fs.existsSync(mfPath)) {
                try { fs.unlinkSync(mfPath); } catch (_) {}
            }
            writeLog('INFO', `[GGUF] Удалён локальный файл ${safeTarget.filename}`);
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', deleted: safeTarget.filename }));
        } catch (delErr) {
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: delErr.message }));
        }
        return;
    }

    if (reqPathname === '/api/models/gguf/modelfile-preview' && req.method === 'POST') {
        readJsonBody(req, (err, body) => {
            if (err || !body || !body.filename) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется поле filename' }));
                return;
            }
            const safeTarget = resolveSafeModelFilePath(body.filename);
            if (!safeTarget) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Некорректное имя .gguf файла' }));
                return;
            }
            const meta = fs.existsSync(safeTarget.fullPath) ? readGgufMetadata(safeTarget.fullPath) : null;
            const detectedTemplate = detectTemplateFamily(safeTarget.filename, meta);
            const chosenFamily = (body.templateFamily && GGUF_TEMPLATES[body.templateFamily]) ? body.templateFamily : detectedTemplate;
            const modelfile = buildModelfileContent({
                ggufFilename: safeTarget.filename,
                ggufAbsPath: safeTarget.fullPath,
                templateFamily: chosenFamily,
                numCtx: body.numCtx,
                temperature: body.temperature,
                systemPrompt: body.systemPrompt,
                useRelativePath: true
            });
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', detectedTemplate, templateFamily: chosenFamily, modelfile }));
        });
        return;
    }

    if (reqPathname === '/api/models/gguf/import' && req.method === 'POST') {
        readJsonBody(req, (err, body) => {
            if (err || !body || !body.filename) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется поле filename' }));
                return;
            }
            const safeTarget = resolveSafeModelFilePath(body.filename);
            if (!safeTarget || !fs.existsSync(safeTarget.fullPath)) {
                res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Указанный .gguf файл не найден в папке models/' }));
                return;
            }

            const ollamaModelName = sanitizeOllamaModelName(body.modelName, safeTarget.filename);
            const meta = readGgufMetadata(safeTarget.fullPath);
            const detectedTemplate = detectTemplateFamily(safeTarget.filename, meta);
            const templateFamily = (body.templateFamily && GGUF_TEMPLATES[body.templateFamily]) ? body.templateFamily : detectedTemplate;
            const modelfileContent = buildModelfileContent({
                ggufFilename: safeTarget.filename,
                ggufAbsPath: safeTarget.fullPath,
                templateFamily,
                numCtx: body.numCtx,
                temperature: body.temperature,
                systemPrompt: body.systemPrompt,
                useRelativePath: true
            });

            const modelfileName = safeTarget.filename.replace(/\.gguf$/i, '.Modelfile');
            const modelfilePath = path.join(MODELS_DIR, modelfileName);
            try {
                fs.writeFileSync(modelfilePath, modelfileContent, 'utf8');
            } catch (mfErr) {
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Не удалось сохранить Modelfile: ' + mfErr.message }));
                return;
            }

            res.writeHead(200, {
                'Content-Type': 'application/x-ndjson; charset=utf-8',
                'Cache-Control': 'no-cache',
                'Connection': 'keep-alive'
            });

            const sendLine = (obj) => {
                if (!res.writableEnded) {
                    res.write(JSON.stringify(obj) + '\n');
                }
            };

            const ollamaPath = getOllamaExecutablePath();
            writeLog('INFO', `[GGUF-Import] Запуск "ollama create ${ollamaModelName} -f ${modelfileName}" (template=${templateFamily})`);
            sendLine({ status: `Создание Modelfile (${templateFamily}) и запуск ollama create ${ollamaModelName}...`, percent: 5 });

            const child = spawn(ollamaPath, ['create', ollamaModelName, '-f', modelfileName], {
                cwd: MODELS_DIR,
                stdio: ['ignore', 'pipe', 'pipe'],
                env: { ...process.env, NO_COLOR: '1' }
            });

            let clientClosed = false;
            let lastOutputLines = [];

            const handleOutput = (chunk) => {
                const cleaned = chunk
                    .toString('utf8')
                    .replace(/\x1B\[[0-9;?]*[A-Za-z]/g, '')
                    .replace(/[\u2800-\u28FF]/g, '')
                    .trim();
                if (!cleaned) return;
                const parts = cleaned.split(/[\r\n]+/).map(s => s.trim()).filter(Boolean);
                for (const line of parts) {
                    lastOutputLines.push(line);
                    if (lastOutputLines.length > 10) lastOutputLines.shift();
                    const pctMatch = line.match(/(\d{1,3})%/);
                    const pct = pctMatch ? Math.min(99, parseInt(pctMatch[1], 10)) : null;
                    sendLine({ status: line, ...(pct !== null ? { percent: pct } : {}) });
                }
            };

            child.stdout.on('data', handleOutput);
            child.stderr.on('data', handleOutput);

            res.on('close', () => {
                if (!res.writableEnded) {
                    clientClosed = true;
                    try { child.kill(); } catch (_) {}
                }
            });

            child.on('error', (spawnErr) => {
                writeLog('ERROR', `[GGUF-Import] Сбой запуска ollama create: ${spawnErr.message}`);
                sendLine({ error: `Не удалось запустить ollama create: ${spawnErr.message}` });
                if (!res.writableEnded) res.end();
            });

            child.on('close', (code) => {
                if (!clientClosed) {
                    if (code === 0) {
                        writeLog('INFO', `[GGUF-Import] Модель "${ollamaModelName}" успешно импортирована в Ollama!`);
                        sendLine({
                            status: `✅ Модель "${ollamaModelName}" создана и готова к работе!`,
                            percent: 100,
                            done: true,
                            modelName: ollamaModelName,
                            templateFamily
                        });
                    } else {
                        const reason = lastOutputLines.slice(-3).join(' | ') || `Код выхода ${code}`;
                        writeLog('ERROR', `[GGUF-Import] Ошибка создания "${ollamaModelName}": ${reason}`);
                        sendLine({ error: `Ошибка ollama create: ${reason}` });
                    }
                }
                if (!res.writableEnded && !res.destroyed) res.end();
            });
        });
        return;
    }

    if (reqPathname === '/api/models/hf/inspect' && req.method === 'GET') {
        let input = '';
        try {
            const parsedReqUrl = new URL(req.url, 'http://127.0.0.1');
            input = parsedReqUrl.searchParams.get('input') || '';
        } catch (_) {}
        const hfToken = req.headers['x-hf-token'] ? String(req.headers['x-hf-token']) : '';
        inspectHuggingFaceInput(input, hfToken)
            .then(data => {
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', ...data }));
            })
            .catch(err => {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: err.message }));
            });
        return;
    }

    if (reqPathname === '/api/models/hf/download' && req.method === 'POST') {
        readJsonBody(req, async (err, body) => {
            if (err || !body || (!body.url && !body.input)) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Требуется ссылка url или input' }));
                return;
            }
            try {
                const job = await startHfDownloadJob(body.url || body.input, body.filename || '', body.hfToken || '');
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', job }));
            } catch (dlErr) {
                const code = dlErr.message.includes('Уже идёт') ? 409 : 400;
                res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: dlErr.message }));
            }
        });
        return;
    }

    if (reqPathname === '/api/models/hf/status' && req.method === 'GET') {
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ status: 'ok', job: getPublicHfDownloadState() }));
        return;
    }

    if (reqPathname === '/api/models/hf/cancel' && req.method === 'POST') {
        const cancelled = cancelActiveHfDownload();
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ status: 'ok', cancelled, job: getPublicHfDownloadState() }));
        return;
    }

    // 9. Модели: Проверка наличия новых поколений и обновлений в Ollama Library
    if (req.url.startsWith('/api/models/check-updates') && req.method === 'GET') {
        const forceRefresh = req.url.includes('force=1');
        writeLog('INFO', `[ModelUpdates] Запрос проверки обновлений моделей (принудительно: ${forceRefresh ? 'да' : 'нет'})...`);
        checkModelUpdates(forceRefresh).then(data => {
            writeLog('INFO', `[ModelUpdates] Найдено обновлений: ${data.updatesCount} (установлено: ${data.installedCount})`);
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify(data));
        }).catch(err => {
            writeLog('ERROR', `[ModelUpdates] Ошибка проверки: ${err.message}`);
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: err.message, updates: [] }));
        });
        return;
    }

    // 9.0 Каталог и древо поколений Qwen
    if (req.url === '/api/models/catalog' && req.method === 'GET') {
        Promise.all([
            getLocalOllamaModels(),
            getSystemHardwareSpecs()
        ]).then(([models, hwSpecs]) => {
            const catalog = buildQwenCatalog(models, hwSpecs);
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', catalog: catalog, hardware: hwSpecs }));
        }).catch(err => {
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: err.message, catalog: [] }));
        });
        return;
    }

    // 9.1 Системные характеристики ПК (Hardware Advisor)
    if (req.url === '/api/system/hardware' && req.method === 'GET') {
        getSystemHardwareSpecs().then(specs => {
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', hardware: specs }));
        }).catch(err => {
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: err.message }));
        });
        return;
    }

    // 9.2 Обогащенный список установленных моделей (классы, мощности, даты, совместимость)
    if (req.url === '/api/tags' && req.method === 'GET') {
        Promise.all([
            getLocalOllamaModels(),
            getSystemHardwareSpecs(),
            fetchOllamaQwenLibrary()
        ]).then(([models, hwSpecs]) => {
            const enrichedModels = models.map(m => {
                const classification = classifyModel(m.name, m.size, hwSpecs);
                return {
                    ...m,
                    classification: classification
                };
            });
            // Хронологическая сортировка: от старых годов выпуска к новым (2024 -> 2025 -> 2026)
            enrichedModels.sort((a, b) => getModelSortScore(a) - getModelSortScore(b));
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ models: enrichedModels }));
        }).catch(err => {
            writeLog('ERROR', `Ошибка получения моделей /api/tags: ${err.message}`);
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ models: [] }));
        });
        return;
    }

    // Вспомогательная функция для генерации саммари (сжатия) через Ollama
    function fetchOllamaSummary(model, promptText) {
        return new Promise((resolve, reject) => {
            const payload = JSON.stringify({
                model: model,
                messages: [
                    { role: 'user', content: promptText }
                ],
                stream: false,
                options: {
                    temperature: 0.3,
                    num_ctx: 16384
                }
            });

            const reqSummary = http.request({
                hostname: '127.0.0.1',
                port: OLLAMA_PORT,
                path: '/api/chat',
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Content-Length': Buffer.byteLength(payload)
                }
            }, (resSummary) => {
                let data = '';
                resSummary.on('data', chunk => { data += chunk; });
                resSummary.on('end', () => {
                    try {
                        const json = JSON.parse(data);
                        if (json.message && json.message.content) {
                            resolve(json.message.content.trim());
                        } else if (json.error) {
                            reject(new Error(json.error));
                        } else {
                            reject(new Error('Пустой ответ от модели'));
                        }
                    } catch (e) {
                        reject(e);
                    }
                });
            });

            reqSummary.on('error', reject);
            reqSummary.write(payload);
            reqSummary.end();
        });
    }

    // 10. Лимиты контекстного окна моделей (Context Window Info)
    if (req.url === '/api/models/context-limits' && req.method === 'GET') {
        const limitsMap = {
            'qwen3.8:27b': 32768,
            'qwen3.6:27b': 32768,
            'qwen3.5:9b': 32768,
            'qwen3.5:4b': 32768,
            'qwen3.5:2b': 32768,
            'qwen3.5:0.8b': 32768,
            'qwen3:30b': 32768,
            'qwen3:8b': 32768,
            'qwen3:4b': 32768,
            'qwen3:1.7b': 32768,
            'qwen3:0.6b': 32768,
            'qwen3-vl:8b': 16384,
            'qwen3-vl:4b': 16384,
            'qwen3-vl:2b': 16384,
            'qwen3-coder:7b': 32768,
            'qwen3-embedding:0.6b': 32768,
            'qwen2.5:32b': 32768,
            'qwen2.5:14b': 32768,
            'qwen2.5:7b': 32768,
            'qwen2.5:3b': 32768,
            'qwen2.5:1.5b': 32768,
            'qwen2.5:0.5b': 32768,
            'qwen2.5-coder:7b': 32768,
            'dolphin3:8b': 32768,
            'dolphin-mistral:7b': 32768,
            'dolphin-llama3:8b': 8192,
            'CognitiveComputations/dolphin-2.9.2-qwen2-7b:Q4_K_M': 32768,
            'llama3.2:3b': 32768,
            'llama3.1:8b': 32768,
            'deepseek-r1:7b': 32768,
            'deepseek-r1:8b': 32768,
            'deepseek-r1:14b': 32768,
            'mistral:7b': 32768,
            'mixtral:8x7b': 32768,
            'gemma2:2b': 8192,
            'gemma2:9b': 8192
        };
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({
            status: 'ok',
            defaultNumCtx: 32768,
            recommendedNumCtx: 32768,
            limits: limitsMap,
            reasons: {
                vision: 'Мультимодальные модели (Vision): лимит 16K оптимизирован для видеопамяти 8 GB VRAM, исключая OOM при кодировании картинок.',
                chat: 'Текстовые модели (Chat): полноценный лимит 32K (8K для Gemma 2 / Llama 3) на полной скорости GPU.'
            }
        }));
        return;
    }

    // 10.1 Профили системных инструкций (Personas API)
    if (req.url === '/api/personas' && req.method === 'GET') {
        try {
            const personas = db.getPersonas();
            const activePersonaId = db.getSetting('active_persona_id', 'default') || 'default';
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', personas, activePersonaId }));
        } catch (err) {
            writeLog('ERROR', `[DB] Ошибка получения профилей Persona: ${err.message}`);
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: err.message }));
        }
        return;
    }

    if (req.url === '/api/personas/active' && req.method === 'POST') {
        readJsonBody(req, (err, body) => {
            try {
                const personaId = body && body.persona_id ? String(body.persona_id).trim() : 'default';
                const persona = db.getPersonaById(personaId);
                if (!persona) {
                    res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
                    res.end(JSON.stringify({ error: 'Профиль не найден' }));
                    return;
                }
                db.setSetting('active_persona_id', personaId);
                if (body && body.chat_id) {
                    db.updateChat(body.chat_id, { persona_id: personaId });
                }
                writeLog('INFO', `[Persona] Активирован профиль: "${persona.name}" (${persona.id})`);
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', activePersonaId: personaId, persona }));
            } catch (activeErr) {
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: activeErr.message }));
            }
        });
        return;
    }

    if (req.url === '/api/personas' && req.method === 'POST') {
        readJsonBody(req, (err, body) => {
            try {
                const created = db.createPersona(body || {});
                writeLog('INFO', `[Persona] Создан новый профиль: "${created.name}" (${created.id})`);
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', persona: created }));
            } catch (createErr) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: createErr.message }));
            }
        });
        return;
    }

    if (req.url.match(/^\/api\/personas\/[a-zA-Z0-9_-]+$/) && req.method === 'PATCH') {
        const personaId = req.url.split('/')[3];
        readJsonBody(req, (err, body) => {
            try {
                const updated = db.updatePersona(personaId, body || {});
                if (!updated) {
                    res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
                    res.end(JSON.stringify({ error: 'Профиль не найден' }));
                    return;
                }
                writeLog('INFO', `[Persona] Обновлён профиль: "${updated.name}" (${updated.id})`);
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', persona: updated }));
            } catch (updateErr) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: updateErr.message }));
            }
        });
        return;
    }

    if (req.url.match(/^\/api\/personas\/[a-zA-Z0-9_-]+$/) && req.method === 'DELETE') {
        const personaId = req.url.split('/')[3];
        try {
            const result = db.deletePersona(personaId);
            if (!result.deleted) {
                const code = result.reason === 'builtin' ? 400 : 404;
                const msg = result.reason === 'builtin' ? 'Встроенные профили нельзя удалять (создайте копию или свой профиль)' : 'Профиль не найден';
                res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: msg }));
                return;
            }
            writeLog('INFO', `[Persona] Удалён профиль ${personaId}`);
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', deleted: true, activePersonaId: db.getSetting('active_persona_id', 'default') }));
        } catch (delErr) {
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: delErr.message }));
        }
        return;
    }

    // 11. База данных SQLite: Список всех диалогов
    if (req.url === '/api/chats' && req.method === 'GET') {
        try {
            let chats = db.getChats();
            if (chats.length === 0) {
                const initialChat = db.createChat('Добро пожаловать в Qwen Studio', 'qwen3.5:9b');
                db.addMessage(initialChat.id, 'assistant', 'Привет! Я локальная нейросеть на твоем компьютере. Выбирай модель, отправляй запросы или загружай файлы в базу знаний.');
                chats = db.getChats();
            }
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', chats }));
        } catch (err) {
            writeLog('ERROR', `[DB] Ошибка получения чатов: ${err.message}`);
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: err.message }));
        }
        return;
    }

    // 12. База данных SQLite: Создание нового диалога
    if (req.url === '/api/chats' && req.method === 'POST') {
        readJsonBody(req, (err, body) => {
            try {
                const title = (body && body.title) ? String(body.title).trim() : 'Новый диалог';
                const model = (body && body.model) ? String(body.model).trim() : 'qwen3.5:9b';
                const personaId = (body && body.persona_id) ? String(body.persona_id).trim() : null;
                const chat = db.createChat(title, model, personaId);
                writeLog('INFO', `[DB] Создан новый диалог: "${chat.title}" (${chat.id})`);
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', chat }));
            } catch (createErr) {
                writeLog('ERROR', `[DB] Ошибка создания чата: ${createErr.message}`);
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: createErr.message }));
            }
        });
        return;
    }

    // 13. База данных SQLite: Получение конкретного диалога и сообщений
    if (req.url.match(/^\/api\/chats\/[a-zA-Z0-9_-]+$/) && req.method === 'GET') {
        const chatId = req.url.split('/')[3];
        try {
            const chat = db.getChatById(chatId);
            if (!chat) {
                res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Диалог не найден' }));
                return;
            }
            const messages = db.getMessages(chatId, true);
            const summary = db.getLatestSummary(chatId);
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', chat, messages, summary }));
        } catch (err) {
            writeLog('ERROR', `[DB] Ошибка загрузки диалога ${chatId}: ${err.message}`);
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: err.message }));
        }
        return;
    }

    // 14. База данных SQLite: Переименование / закрепление диалога
    if (req.url.match(/^\/api\/chats\/[a-zA-Z0-9_-]+$/) && req.method === 'PATCH') {
        const chatId = req.url.split('/')[3];
        readJsonBody(req, (err, body) => {
            try {
                const updated = db.updateChat(chatId, body || {});
                if (!updated) {
                    res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
                    res.end(JSON.stringify({ error: 'Диалог не найден' }));
                    return;
                }
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ status: 'ok', chat: updated }));
            } catch (updateErr) {
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: updateErr.message }));
            }
        });
        return;
    }

    // 15. База данных SQLite: Удаление диалога
    if (req.url.match(/^\/api\/chats\/[a-zA-Z0-9_-]+$/) && req.method === 'DELETE') {
        const chatId = req.url.split('/')[3];
        try {
            const deleted = db.deleteChat(chatId);
            writeLog('INFO', `[DB] Диалог ${chatId} удален.`);
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', deleted }));
        } catch (err) {
            writeLog('ERROR', `[DB] Ошибка удаления диалога ${chatId}: ${err.message}`);
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: err.message }));
        }
        return;
    }

    // 16. Волшебная кнопка: Интеллектуальное сжатие контекста диалога
    if (req.url.match(/^\/api\/chats\/[a-zA-Z0-9_-]+\/compact$/) && req.method === 'POST') {
        const chatId = req.url.split('/')[3];
        const chat = db.getChatById(chatId);
        if (!chat) {
            res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ error: 'Диалог не найден' }));
            return;
        }

        const uncompacted = db.getMessages(chatId, false);
        if (uncompacted.length < 3) {
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({
                status: 'skipped',
                message: 'В диалоге слишком мало сообщений для сжатия (минимум 3).'
            }));
            return;
        }

        const transcript = uncompacted.map(m => {
            const roleName = m.role === 'user' ? 'Пользователь' : (m.role === 'summary' ? 'Предыдущая выжимка' : 'Ассистент');
            return `${roleName}: ${m.content}`;
        }).join('\n\n');

        const tokensBefore = uncompacted.reduce((acc, m) => acc + (m.prompt_tokens || 0) + (m.eval_tokens || 0), 0) || Math.round(transcript.length / 3);

        const summarizerPrompt = `Ты — высокоточный архиватор контекста для искусственного интеллекта.
Твоя задача — сжать историю переписки в плотную структурированную выжимку, отбросив все приветствия, вежливости и воду.

ОБЯЗАТЕЛЬНАЯ СТРУКТУРА ВЫЖИМКИ:
🎯 ЦЕЛЬ: что именно делает пользователь
💡 ПРИНЯТЫЕ РЕШЕНИЯ: архитектура, технологии, алгоритмы, согласованный стек
📁 СУЩНОСТИ И ПУТИ: имена файлов, функций, портов, конфигураций
⚡ ТЕКУЩИЙ СТАТУС И СЛЕДУЮЩИЙ ШАГ: на чем конкретно остановились

ПЕРЕПИСКА ДЛЯ СЖАТИЯ:
${transcript.substring(0, 14000)}

Сформируй плотную выжимку на русском языке строго по структуре выше:`;

        writeLog('INFO', `[Compactor] Запуск сжатия диалога ${chatId} (${uncompacted.length} сообщений, ~${tokensBefore} токенов)...`);

        const fastModel = 'qwen3:0.6b';
        fetchOllamaSummary(fastModel, summarizerPrompt).then(summaryText => {
            const tokensAfter = Math.round(summaryText.length / 3);
            const savedTokens = Math.max(0, tokensBefore - tokensAfter);
            const savedPercent = tokensBefore > 0 ? Math.round((savedTokens / tokensBefore) * 100) : 0;

            db.markMessagesCompacted(chatId);
            db.addSummary(chatId, summaryText, tokensBefore, tokensAfter, uncompacted.length);

            writeLog('INFO', `[Compactor] Диалог ${chatId} успешно сжат: освобождено ${savedTokens} токенов (${savedPercent}% сжатия)`);
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({
                status: 'ok',
                summary: summaryText,
                tokensBefore: tokensBefore,
                tokensAfter: tokensAfter,
                savedTokens: savedTokens,
                savedPercent: savedPercent,
                compactedCount: uncompacted.length
            }));
        }).catch(err => {
            writeLog('WARN', `[Compactor] Сбой сжатия через ${fastModel}: ${err.message}. Пробуем модель чата ${chat.model}...`);
            fetchOllamaSummary(chat.model, summarizerPrompt).then(summaryText => {
                const tokensAfter = Math.round(summaryText.length / 3);
                const savedTokens = Math.max(0, tokensBefore - tokensAfter);
                const savedPercent = tokensBefore > 0 ? Math.round((savedTokens / tokensBefore) * 100) : 0;

                db.markMessagesCompacted(chatId);
                db.addSummary(chatId, summaryText, tokensBefore, tokensAfter, uncompacted.length);

                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({
                    status: 'ok',
                    summary: summaryText,
                    tokensBefore: tokensBefore,
                    tokensAfter: tokensAfter,
                    savedTokens: savedTokens,
                    savedPercent: savedPercent,
                    compactedCount: uncompacted.length
                }));
            }).catch(fallbackErr => {
                writeLog('ERROR', `[Compactor] Сбой сжатия: ${fallbackErr.message}`);
                res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: fallbackErr.message }));
            });
        });
        return;
    }

    // 17. Чат с поддержкой сохранения сессий в SQLite и подсчетом токенов
    if (req.url === '/api/chat' && req.method === 'POST') {
        readJsonBody(req, (err, payload) => {
            if (err || !payload) {
                res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Некорректный JSON запрос' }));
                return;
            }

            const chatId = payload.chat_id || null;
            const messages = payload.messages || [];
            const lastUserMsg = [...messages].reverse().find(m => m.role === 'user');

            // Сохраняем сообщение пользователя в SQLite
            if (chatId && lastUserMsg) {
                try {
                    if (!db.getChatById(chatId)) {
                        db.createChat('Новый диалог', payload.model || 'qwen3.5:9b', payload.persona_id || null);
                    }
                    db.addMessage(chatId, 'user', lastUserMsg.content);
                } catch (dbErr) {
                    writeLog('WARN', `[DB] Ошибка сохранения user сообщения: ${dbErr.message}`);
                }
            }

            // Удаляем служебные поля перед передачей в Ollama (Ollama их не ожидает)
            delete payload.chat_id;
            delete payload.persona_id;

            // Обеспечиваем оптимальный лимит контекстного окна под архитектуру модели
            if (!payload.options) payload.options = {};
            if (!payload.options.num_ctx) {
                const mLower = (payload.model || '').toLowerCase();
                const isVision = mLower.includes('-vl') || mLower.includes('vl:');
                const is8kModel = mLower.startsWith('gemma2') || mLower.startsWith('dolphin-llama3');
                payload.options.num_ctx = is8kModel ? 8192 : (isVision ? 16384 : 32768);
            }

            const ollamaPayload = JSON.stringify(payload);
            const proxyReq = http.request({
                hostname: '127.0.0.1',
                port: OLLAMA_PORT,
                path: '/api/chat',
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Content-Length': Buffer.byteLength(ollamaPayload)
                }
            }, (proxyRes) => {
                res.writeHead(proxyRes.statusCode, {
                    'Content-Type': proxyRes.headers['content-type'] || 'application/x-ndjson',
                    'Transfer-Encoding': 'chunked'
                });

                let fullContent = '';
                let fullThinking = '';
                let promptEvalCount = 0;
                let evalCount = 0;

                proxyRes.on('data', (chunk) => {
                    res.write(chunk);

                    const str = chunk.toString();
                    const lines = str.split('\n');
                    for (const line of lines) {
                        if (!line.trim()) continue;
                        try {
                            const parsed = JSON.parse(line);
                            if (parsed.message) {
                                if (parsed.message.content) fullContent += parsed.message.content;
                                if (parsed.message.thinking) fullThinking += parsed.message.thinking;
                            }
                            if (parsed.prompt_eval_count) promptEvalCount = parsed.prompt_eval_count;
                            if (parsed.eval_count) evalCount = parsed.eval_count;
                        } catch (e) {}
                    }
                });

                proxyRes.on('end', () => {
                    res.end();

                    // Сохраняем ответ ассистента и метрики токенов в SQLite
                    if (chatId && (fullContent || fullThinking)) {
                        try {
                            db.addMessage(chatId, 'assistant', fullContent, fullThinking, promptEvalCount, evalCount);

                            // Авто-название для чата, если оно все еще "Новый диалог"
                            const currentChat = db.getChatById(chatId);
                            if (currentChat && currentChat.title === 'Новый диалог' && lastUserMsg) {
                                let cleanTitle = lastUserMsg.content.trim().replace(/[#*`\n]/g, ' ');
                                if (cleanTitle.length > 36) {
                                    cleanTitle = cleanTitle.substring(0, 36).trim() + '...';
                                }
                                if (cleanTitle) {
                                    db.updateChat(chatId, { title: cleanTitle });
                                    writeLog('INFO', `[DB] Чат ${chatId} автоматически переименован в: "${cleanTitle}"`);
                                }
                            }
                        } catch (dbErr) {
                            writeLog('WARN', `[DB] Ошибка сохранения assistant сообщения: ${dbErr.message}`);
                        }
                    }
                });
            });

            proxyReq.on('error', (err) => {
                writeLog('ERROR', `Ошибка связи с Ollama (/api/chat): ${err.message}`);
                if (err.code === 'ECONNREFUSED') {
                    ensureOllamaRunning();
                }
                res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ error: 'Служба Ollama еще запускается. Повторите запрос через 3 сек.' }));
            });

            proxyReq.write(ollamaPayload);
            proxyReq.end();
        });
        return;
    }

    // 18. Проксирование остальных запросов к Ollama (/api/pull, /api/delete, /api/embeddings и т.д.)
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

    // Отдача иконки и фавикона
    if (req.url === '/icon.png' || req.url === '/favicon.ico') {
        const iconFile = req.url === '/favicon.ico' ? 'icon.ico' : 'icon.png';
        const iconPath = path.join(__dirname, iconFile);
        if (fs.existsSync(iconPath)) {
            res.writeHead(200, {
                'Content-Type': req.url === '/favicon.ico' ? 'image/x-icon' : 'image/png',
                'Cache-Control': 'public, max-age=86400'
            });
            fs.createReadStream(iconPath).pipe(res);
            return;
        }
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

    // Открытие окна приложения: если не передан флаг --no-open
    if (!process.argv.includes('--no-open') && process.env.NO_AUTO_OPEN !== '1') {
        const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
        if (fs.existsSync(edgePath)) {
            writeLog('INFO', 'Запуск в отдельном окне приложения (Microsoft Edge App Mode)...');
            exec(`"${edgePath}" --app="http://localhost:${PORT}" --window-size=1300,880`);
        } else {
            writeLog('INFO', 'Запуск в браузере по умолчанию...');
            exec(`start http://localhost:${PORT}`);
        }
    }
});
