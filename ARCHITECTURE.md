# Project Architecture & Tech Stack Guide (for Future Agents) 🤖📚

This document is created for future AI agents and developers working on **Qwen AI Studio (Local)**.

---

## 1. Hardware Stack

- **CPU**: AMD Ryzen 7 5700X (8 cores, 16 threads).
- **GPU**: AMD Radeon RX 6600 (8 GB VRAM GDDR6). Compute runs on the GPU compute engine (ROCm / DirectML / Vulkan support via Ollama).
- **RAM**: 16 GB DDR4 (under peak load, 8B models and cache occupy up to 14–15 GB).
- **OS**: Windows (Shell: PowerShell / CMD).

---

## 2. Software Stack

- **Inference Service**: [Ollama](https://ollama.com/)
  - Executable path: `%LOCALAPPDATA%\Programs\Ollama\ollama.exe`
  - Port: `127.0.0.1:11434`
  - Added to Windows registry startup: `HKCU:\Software\Microsoft\Windows\CurrentVersion\Run`
  - Health-check endpoint: `GET http://127.0.0.1:11434/api/tags`
- **Backend**: Node.js (v24.x)
  - Entry point: `server.js`
  - URL: `http://localhost:3000`
  - Zero-dependency architecture: built-in Node.js modules only (`http`, `fs`, `path`, `child_process`, etc.).
  - Logging: written with local timestamps (not UTC!) to `app.log` and `stdout`.
- **Frontend**:
  - File: `chat.html` (served by the backend at `http://localhost:3000`).
  - Pure HTML5, CSS3 (adaptive dark theme in GitHub Dark style), native Vanilla JavaScript without heavy frameworks.
  - Streaming response handling (`fetch` + `ReadableStreamDefaultReader` + `TextDecoder`).
  - Real-time reasoning display (`Thinking...` block).
- **Version Control**:
  - Git repository: `https://github.com/crabinacrabic/qwen-local-studio`
  - Branch: `main`

---

## 3. Neural Network Models in the System

### A. Generative Text Models (for Chat)
1. **`qwen3:8b` (Flagship)**: 8B parameters, Q4_K_M (~5.2 GB). Primary thinking model (`<thought>`), strong Russian responses, analyzes complex documents.
2. **`qwen3:4b`**: 4B parameters (~2.5 GB). Great speed/quality balance for fast responses.
3. **`qwen3:1.7b`**: 1.7B parameters (~1.4 GB). Fast compact model.
4. **`qwen3:0.6b`**: 0.6B parameters (~522 MB). Ultra-lightweight model for quick tests and simple replies.

### B. Embedding Models (Internal Utility, RAG ONLY)
- **`nomic-embed-text`**: 137M parameters (~274 MB).
  - Generates 768-dimensional semantic embedding vectors (`POST /api/embeddings`).
  - **IMPORTANT**: Filtered out of the chat model selector and model cards in `chat.html` so users do not try to chat with it directly (it cannot generate chat dialogue).

---

## 4. Local RAG Architecture (Knowledge Base)

All processing runs **100% locally and offline**:

```
[User Document: .txt, .md, .docx, source code]
                       │
                       ▼
  [Unpacking & Text Parsing]
    - For .docx: system tar.exe extracts word/document.xml,
      the script strips XML tags and extracts clean UTF-8 text.
    - For plain text/code: direct UTF-8 read.
                       │
                       ▼
    [Smart Chunking (chunkText)]
    - ~500-character chunks with 70-character overlap
    - Hard length limit to stay within embedding context bounds
                       │
                       ▼
    [Vectorization via nomic-embed-text]
    - Request to Ollama POST /api/embeddings
    - Euclidean vector norm computation (norm)
                       │
                       ▼
   [Storage: knowledge_index.json & knowledge/]
    - Chunks: { id, filename, chunkIndex, text, embedding, norm, createdAt }
    - Source files: knowledge/ directory (listed in .gitignore)
                       │
                       ▼
    [Semantic Search (POST /api/rag/search)]
    - User query is vectorized via nomic-embed-text
    - Cosine similarity computed against all in-memory chunks
    - Descending sort by score, filter (score >= 0.35), Top-3 results
                       │
                       ▼
     [Prompt Augmentation & Dispatch to Qwen]
    - Chat UI shows widget: "📚 Использовано источников: N (совпадение XX%)"
    - System prompt injects context from matched chunks
    - Qwen responds strictly from document facts with reasoning block
```

---

## 5. Live Web Search & Browsing Architecture 🌐

Live internet access works **without paid APIs, without registration, and without external npm packages**:

```
[User Query / Direct URL]
                 │
                 ▼
     [Analysis & Routing]
    ├── Contains direct URL ──► POST /api/web/fetch (deep page reading)
    └── "🌐 Интернет: ВКЛ" enabled ──► POST /api/web/search (DuckDuckGo Search)
                 │
                 ▼
      [server.js Networking Module]
    - Primary engine: DuckDuckGo HTML (POST https://html.duckduckgo.com/html/)
    - Fallback engine: DuckDuckGo Lite (POST https://lite.duckduckgo.com/lite/)
    - Target link decoding (uddg redirect unwrap)
    - HTML tag and entity cleanup for titles and snippets
                 │
                 ▼
      [Deep Fetch]
    - Loads main text (up to 1600 chars) for Top-2 most relevant sites
    - Strips scripts, styles, headers, footers, and navigation
                 │
                 ▼
      [Context Augmentation for Qwen]
    - Injects current date (сегодня / today: YYYY-MM-DD)
    - Block: === АКТУАЛЬНЫЕ РЕЗУЛЬТАТЫ ПОИСКА В ИНТЕРНЕТЕ ===
    - Citation numbering [1], [2], [3]
                 │
                 ▼
         [Generation & UI]
    - Model cites sources [1], [2] during reasoning and final answer
    - UI renders an interactive collapsible source-cards block with clickable links
```

---

## 6. Auto-Check & Safe Model Update Architecture 🔄

```
[chat.html Startup / Click "Проверить новые модели"]
                         │
                         ▼
        [GET /api/models/check-updates]
        ├── Local models: GET /api/tags
        └── Remote library: GET https://ollama.com/search?q=qwen
                         │
                         ▼
          [Generation & VRAM Matching]
        - qwen3:8b ➔ qwen3.5:9b (flagship, 256K context, ~6.6 GB)
        - qwen3:4b ➔ qwen3.5:4b (fast, ~3.4 GB)
        - nomic-embed-text ➔ qwen3-embedding:0.6b (RAG, 32K context)
        - qwen3-vl:4b (Vision recommendation)
                         │
                         ▼
             [Interactive UI Dialog]
    User chooses a strategy per model:
    🔘 Replace (download new + delete old to save disk space)
    🔘 Keep both models
    🔘 Skip
                         │
                         ▼
          [Safe Update Pipeline]
    1. POST /api/pull (streaming download with 0–100% progress bar)
    2. Completion verification
    3. If "Replace" chosen: DELETE /api/delete for the old model
       (ONLY after 100% success of the new download!)
    4. Automatic switch of chat selector to the new model
```

---

## 7. Key Files & Scripts

| File | Purpose |
| :--- | :--- |
| **`server.js`** | Local Node.js server: Ollama proxy, RAG (`/api/rag/*`), web search (`/api/web/*`), model update checks (`/api/models/check-updates`), autonomous ReAct agent (`/api/agent/research`), logging to `app.log`. |
| **`chat.html`** | Web chat UI: Web Search / RAG / Deep Research toggles, model update modal, downloader, knowledge base panel, session sidebar. |
| **`start_web_chat.bat`** | Launcher: checks/starts Ollama if needed, starts `node server.js`, and opens the browser. |
| **`git_push.bat`** | Convenience 1-click script to commit and push changes to GitHub. |
| **`knowledge_index.json`** | Vector index and document chunks (stored locally, `.gitignore`). |
| **`knowledge/`** | Directory storing copies of user-uploaded files (`.gitignore`). |
| **`app.log`** | Server and UI event log with local timestamps (`.gitignore`). |
| **`.gitignore`** | Excludes logs, model weights, personal documents, and vector database from the public repository. |

---

## 8. Rules for Future Agents

1. **Language of Reasoning, Docs & Reports**: Always think/reason in English and write internal technical documentation (e.g., `ARCHITECTURE.md`, `ROADMAP.md`) in English. Always write user-facing responses and report artifacts (e.g., `implementation_plan.md`, `walkthrough.md`) in Russian.
2. **Data Safety**: Never commit contents of `knowledge/`, `knowledge_index.json`, `app.log`, SQLite DB files, or binary model weights to git.
3. **RX 6600 GPU Compatibility**: Do not run two heavy 14B+ models simultaneously, as VRAM is limited to 8 GB. For general chat `qwen3.5:9b` (or 7B–8B quantized models) fits best; for lightweight tasks use `qwen3.5:4b`.
4. **Model Filtering**: Any new embedding models (containing `embed`, `bge`, `minilm`) must remain hidden from the main chat model selector.
5. **Zero-Dependency Stack**: Do not add heavy npm dependencies if a task can be solved using native Node.js built-ins (`http`, `https`, `url`, `fs`, `child_process`, `node:sqlite`).
6. **Safe Model Deletion**: Never delete a working older model until the replacement model has downloaded 100% and passed verification.
