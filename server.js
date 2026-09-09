const http = require('http');
const https = require('https');
const { URL } = require('url');
const fs = require('fs');
const path = require('path');
const { exec, spawn, execSync, execFile } = require('child_process');

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
            tierName = '🪶 Ультралегкий';
            approxVramGB = 0.6;
        } else {
            tier = 'flagship';
            tierName = '⚡ Продвинутый';
            approxVramGB = 4.0;
        }
        generation = '3.0';
        releaseDate = OLLAMA_DATES_CACHE.dates['qwen3-embedding'] || 'Сентябрь 2025';
        lineage = 'Официальные эмбеддинги Qwen для базы знаний (контекст 32K)';
    } else if (name.includes('-vl') || name.includes('vl:')) {
        modelClass = 'vision';
        className = '👁️ Зрение (Vision)';
        if (name.includes('2b')) {
            tier = 'light';
            tierName = '🪶 Компактный';
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
            tierName = '🪶 Компактный';
            approxVramGB = 2.0;
        } else if (name.includes('7b')) {
            tier = 'flagship';
            tierName = '⚡ Флагман';
            approxVramGB = 4.7;
        } else if (name.includes('14b')) {
            tier = 'flagship';
            tierName = '⚡ Продвинутый';
            approxVramGB = 9.5;
        } else {
            tier = 'expert';
            tierName = '🧠 Экспертный';
            approxVramGB = 18.0;
        }
        generation = name.includes('qwen3') ? '3.0' : '2.5';
        releaseDate = name.includes('qwen3') ? (OLLAMA_DATES_CACHE.dates['qwen3-coder'] || 'Сентябрь 2025') : (OLLAMA_DATES_CACHE.dates['qwen2.5'] || 'Сентябрь 2024');
        lineage = 'Специализированная модель для разработки и рефакторинга';
    } else {
        modelClass = 'chat';
        className = '💬 Чат и логика';

        if (name.includes('0.5b') || name.includes('0.6b') || name.includes('1.5b') || name.includes('1.7b') || name.includes(':2b') || name.includes('0.8b')) {
            tier = 'light';
            tierName = '🪶 Ультралегкий';
            approxVramGB = 1.8;
        } else if (name.includes(':3b') || name.includes(':4b')) {
            tier = 'balanced';
            tierName = '🚀 Сбалансированный';
            approxVramGB = 3.4;
        } else if (name.includes(':7b') || name.includes(':8b') || name.includes(':9b') || name.includes(':14b')) {
            tier = 'flagship';
            tierName = '⚡ Флагман';
            approxVramGB = 6.6;
        } else {
            tier = 'expert';
            tierName = '🧠 Экспертный';
            approxVramGB = 18.0;
        }

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

    const vramLimit = hwSpecs && hwSpecs.vramGB ? hwSpecs.vramGB : 8.0;
    const ramLimit = hwSpecs && hwSpecs.ramTotalGB ? hwSpecs.ramTotalGB : 16.0;

    let compatibility = {
        status: 'gpu_ready',
        color: '#2ea043',
        badge: '🟢 100% GPU',
        text: `Влезает в видеопамять (${vramLimit} GB VRAM) — максимальная скорость`
    };

    if (approxVramGB > (vramLimit - 0.5)) {
        if (approxVramGB <= (ramLimit - 2.0)) {
            compatibility = {
                status: 'ram_offload',
                color: '#d29922',
                badge: '🟡 RAM Offload',
                text: `Требует оперативную память (${ramLimit} GB RAM) — умеренная скорость`
            };
        } else {
            compatibility = {
                status: 'oom_risk',
                color: '#da3633',
                badge: '🔴 Недостаточно памяти',
                text: 'Слишком тяжелая модель для текущей конфигурации'
            };
        }
    }

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
            description: 'Свежая компактная модель: скорость 50+ токенов/сек, мультимодальность, идеально для повседневных задач.',
            vramFit: '🟢 100% GPU (Занимает меньше половины VRAM)'
        },
        {
            check: (name) => name.startsWith('qwen3:0.6b') || name.startsWith('qwen3:1.7b'),
            modelClass: 'chat',
            className: '💬 Чат и логика',
            tier: 'light',
            tierName: '🪶 УЛЬТРАЛЕГКИЙ',
            targetFamily: 'qwen3.5',
            targetTag: 'qwen3.5:2b',
            title: 'Qwen3.5 2B (Компактная)',
            sizeApprox: '~1.8 GB',
            oldGeneration: 'Qwen3 1.7B / 0.6B (Октябрь 2025)',
            newGeneration: 'Qwen3.5 2B (Сентябрь 2026)',
            targetReleaseDate: OLLAMA_DATES_CACHE.dates['qwen3.5'] || 'Сентябрь 2026',
            evolution: 'Qwen3 0.6B/1.7B (Окт 2025) ➔ Qwen3.5 2B (Сент 2026) 🚀',
            description: 'Миниатюрная модель 2026 года нового поколения с блоками размышлений.',
            vramFit: '🟢 100% GPU (Минимальная нагрузка)'
        },
        {
            check: (name) => name.startsWith('nomic-embed-text'),
            modelClass: 'embedding',
            className: '⚡ Эмбеддинги (RAG)',
            tier: 'light',
            tierName: '🪶 СТАНДАРТ RAG',
            targetFamily: 'qwen3-embedding',
            targetTag: 'qwen3-embedding:0.6b',
            title: 'Qwen3-Embedding 0.6B (RAG)',
            sizeApprox: '~600 MB',
            oldGeneration: 'nomic-embed-text (Базовая)',
            newGeneration: 'Qwen3-Embedding 0.6B (Сентябрь 2025)',
            targetReleaseDate: OLLAMA_DATES_CACHE.dates['qwen3-embedding'] || 'Сентябрь 2025',
            evolution: 'nomic-embed-text ➔ Qwen3-Embedding 0.6B ✨',
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
                description: 'Новая модель со зрением: распознавание изображений, графиков, скриншотов и документов.',
                vramFit: '🟢 Идеально для 8 GB VRAM'
            });
        }
    }

    const result = {
        status: 'ok',
        hardware: hwSpecs,
        installedCount: installed.length,
        updatesCount: updates.length,
        updates: updates,
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
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ models: enrichedModels }));
        }).catch(err => {
            writeLog('ERROR', `Ошибка получения моделей /api/tags: ${err.message}`);
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ models: [] }));
        });
        return;
    }

    // 10. Проксирование остальных запросов к Ollama (/api/chat, /api/pull, /api/delete и т.д.)
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
