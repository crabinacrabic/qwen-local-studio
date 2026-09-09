/**
 * db.js — Нативная база данных SQLite для Qwen Local AI Studio
 * Использует встроенный модуль node:sqlite из Node.js 24+ (Zero-dependency)
 */

const { DatabaseSync } = require('node:sqlite');
const path = require('path');
const fs = require('fs');

const DB_FILE = path.join(__dirname, 'qwen_studio.db');

// Инициализация базы данных
const db = new DatabaseSync(DB_FILE);

// Включаем WAL (Write-Ahead Logging) для максимальной производительности и надежности
db.exec('PRAGMA journal_mode = WAL;');
db.exec('PRAGMA synchronous = NORMAL;');
db.exec('PRAGMA foreign_keys = ON;');

// Создание таблиц
db.exec(`
    CREATE TABLE IF NOT EXISTS chats (
        id TEXT PRIMARY KEY,
        title TEXT NOT NULL,
        model TEXT NOT NULL,
        pinned INTEGER DEFAULT 0,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chat_id TEXT NOT NULL,
        role TEXT NOT NULL,
        content TEXT NOT NULL,
        thinking TEXT DEFAULT '',
        prompt_tokens INTEGER DEFAULT 0,
        eval_tokens INTEGER DEFAULT 0,
        is_compacted INTEGER DEFAULT 0,
        created_at TEXT NOT NULL,
        FOREIGN KEY (chat_id) REFERENCES chats(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS chat_summaries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chat_id TEXT NOT NULL,
        summary TEXT NOT NULL,
        tokens_before INTEGER NOT NULL,
        tokens_after INTEGER NOT NULL,
        compacted_messages_count INTEGER NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY (chat_id) REFERENCES chats(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS user_memories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT DEFAULT 'general',
        fact TEXT NOT NULL,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    );
`);

// Вспомогательная функция локального времени
function getLocalTimestamp() {
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// ==========================================
// МЕТОДЫ УПРАВЛЕНИЯ ЧАТАМИ (CHATS)
// ==========================================

function createChat(title = 'Новый диалог', model = 'qwen3.5:9b') {
    const id = 'chat_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
    const now = getLocalTimestamp();
    const stmt = db.prepare(`
        INSERT INTO chats (id, title, model, pinned, created_at, updated_at)
        VALUES (?, ?, ?, 0, ?, ?)
    `);
    stmt.run(id, title, model, now, now);
    return getChatById(id);
}

function getChats() {
    const stmt = db.prepare(`
        SELECT c.*, 
               (SELECT COUNT(*) FROM messages m WHERE m.chat_id = c.id AND m.is_compacted = 0) as active_messages_count,
               (SELECT content FROM messages m WHERE m.chat_id = c.id ORDER BY m.id DESC LIMIT 1) as last_message_preview,
               (SELECT created_at FROM messages m WHERE m.chat_id = c.id ORDER BY m.id DESC LIMIT 1) as last_message_time
        FROM chats c
        ORDER BY c.pinned DESC, c.updated_at DESC
    `);
    return stmt.all();
}

function getChatById(id) {
    const stmt = db.prepare(`SELECT * FROM chats WHERE id = ?`);
    const chat = stmt.get(id);
    return chat || null;
}

function updateChat(id, updates = {}) {
    const chat = getChatById(id);
    if (!chat) return null;

    const title = updates.title !== undefined ? updates.title : chat.title;
    const model = updates.model !== undefined ? updates.model : chat.model;
    const pinned = updates.pinned !== undefined ? (updates.pinned ? 1 : 0) : chat.pinned;
    const now = getLocalTimestamp();

    const stmt = db.prepare(`
        UPDATE chats 
        SET title = ?, model = ?, pinned = ?, updated_at = ?
        WHERE id = ?
    `);
    stmt.run(title, model, pinned, now, id);
    return getChatById(id);
}

function touchChat(id) {
    const now = getLocalTimestamp();
    const stmt = db.prepare(`UPDATE chats SET updated_at = ? WHERE id = ?`);
    stmt.run(now, id);
}

function deleteChat(id) {
    // Каскадное удаление сообщений и саммари
    db.prepare(`DELETE FROM messages WHERE chat_id = ?`).run(id);
    db.prepare(`DELETE FROM chat_summaries WHERE chat_id = ?`).run(id);
    const stmt = db.prepare(`DELETE FROM chats WHERE id = ?`);
    const res = stmt.run(id);
    return res.changes > 0;
}

// ==========================================
// МЕТОДЫ УПРАВЛЕНИЯ СООБЩЕНИЯМИ (MESSAGES)
// ==========================================

function addMessage(chatId, role, content, thinking = '', promptTokens = 0, evalTokens = 0, isCompacted = 0) {
    const now = getLocalTimestamp();
    const stmt = db.prepare(`
        INSERT INTO messages (chat_id, role, content, thinking, prompt_tokens, eval_tokens, is_compacted, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const result = stmt.run(chatId, role, content, thinking, promptTokens, evalTokens, isCompacted, now);
    touchChat(chatId);
    return {
        id: Number(result.lastInsertRowid),
        chat_id: chatId,
        role,
        content,
        thinking,
        prompt_tokens: promptTokens,
        eval_tokens: evalTokens,
        is_compacted: isCompacted,
        created_at: now
    };
}

function getMessages(chatId, includeCompacted = false) {
    if (includeCompacted) {
        const stmt = db.prepare(`
            SELECT * FROM messages 
            WHERE chat_id = ? 
            ORDER BY id ASC
        `);
        return stmt.all(chatId);
    } else {
        const stmt = db.prepare(`
            SELECT * FROM messages 
            WHERE chat_id = ? AND is_compacted = 0 
            ORDER BY id ASC
        `);
        return stmt.all(chatId);
    }
}

function getCompactedMessages(chatId) {
    const stmt = db.prepare(`
        SELECT * FROM messages 
        WHERE chat_id = ? AND is_compacted = 1 
        ORDER BY id ASC
    `);
    return stmt.all(chatId);
}

function markMessagesCompacted(chatId, messageIds = []) {
    if (!messageIds || messageIds.length === 0) {
        const stmt = db.prepare(`
            UPDATE messages 
            SET is_compacted = 1 
            WHERE chat_id = ? AND role NOT IN ('system', 'summary')
        `);
        return stmt.run(chatId);
    }

    const placeholders = messageIds.map(() => '?').join(',');
    const stmt = db.prepare(`
        UPDATE messages 
        SET is_compacted = 1 
        WHERE chat_id = ? AND id IN (${placeholders})
    `);
    return stmt.run(chatId, ...messageIds);
}

// ==========================================
// МЕТОДЫ САММАРИЗАЦИИ И СЖАТИЯ КОНТЕКСТА
// ==========================================

function addSummary(chatId, summary, tokensBefore, tokensAfter, compactedCount) {
    const now = getLocalTimestamp();
    const stmt = db.prepare(`
        INSERT INTO chat_summaries (chat_id, summary, tokens_before, tokens_after, compacted_messages_count, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    `);
    const result = stmt.run(chatId, summary, tokensBefore, tokensAfter, compactedCount, now);
    
    // Также сохраняем саммари как системно-видимую карточку в таблицу сообщений
    addMessage(chatId, 'summary', summary, '', 0, 0, 0);

    return {
        id: Number(result.lastInsertRowid),
        chat_id: chatId,
        summary,
        tokens_before: tokensBefore,
        tokens_after: tokensAfter,
        compacted_messages_count: compactedCount,
        created_at: now
    };
}

function getLatestSummary(chatId) {
    const stmt = db.prepare(`
        SELECT * FROM chat_summaries 
        WHERE chat_id = ? 
        ORDER BY id DESC 
        LIMIT 1
    `);
    return stmt.get(chatId) || null;
}

// ==========================================
// ДОЛГОСРОЧНАЯ ПАМЯТЬ О ПОЛЬЗОВАТЕЛЕ (MEMORIES)
// ==========================================

function getMemories() {
    const stmt = db.prepare(`SELECT * FROM user_memories ORDER BY id ASC`);
    return stmt.all();
}

function addMemory(category, fact) {
    const now = getLocalTimestamp();
    const stmt = db.prepare(`
        INSERT INTO user_memories (category, fact, created_at)
        VALUES (?, ?, ?)
    `);
    const result = stmt.run(category || 'general', fact, now);
    return {
        id: Number(result.lastInsertRowid),
        category: category || 'general',
        fact,
        created_at: now
    };
}

function deleteMemory(id) {
    const stmt = db.prepare(`DELETE FROM user_memories WHERE id = ?`);
    const res = stmt.run(id);
    return res.changes > 0;
}

// ==========================================
// НАСТРОЙКИ (SETTINGS)
// ==========================================

function getSetting(key, defaultValue = null) {
    const stmt = db.prepare(`SELECT value FROM settings WHERE key = ?`);
    const row = stmt.get(key);
    return row ? row.value : defaultValue;
}

function setSetting(key, value) {
    const stmt = db.prepare(`
        INSERT INTO settings (key, value) 
        VALUES (?, ?) 
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
    `);
    stmt.run(key, String(value));
}

module.exports = {
    db,
    createChat,
    getChats,
    getChatById,
    updateChat,
    deleteChat,
    addMessage,
    getMessages,
    getCompactedMessages,
    markMessagesCompacted,
    addSummary,
    getLatestSummary,
    getMemories,
    addMemory,
    deleteMemory,
    getSetting,
    setSetting
};
