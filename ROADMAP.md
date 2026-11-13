# Qwen AI Studio (Local) Project Roadmap 🚀

An autonomous local workspace for open-weights AI models built on pure Node.js (Zero-Dependency) and Ollama.

---

## ✅ Implemented (Status: Ready for Use)

### 1. System Core & Stability
- [x] Local web chat with streaming responses and reasoning blocks (`<think>`).
- [x] Automatic background startup of the Ollama service on server boot with auto-reconnect on `ECONNREFUSED`.
- [x] Event logging to `app.log` with local timestamps.
- [x] Full autonomy and clean architecture: **0 external npm dependencies** (native Node.js 24).
- [x] Helper script `git_push.bat` for instant synchronization with GitHub.

### 2. Hardware Monitoring & Model Evolution Tree
- [x] **Native System Scanner (Hardware Advisor)** via PowerShell WMI/CIM:
  - Detected GPU and VRAM (AMD RX 6600 8 GB);
  - System memory capacity (16 GB RAM);
  - Free space on system drive.
- [x] **Model Catalog with Capability Classification**:
  - Categories: `💬 Chat & Logic`, `👁️ Vision`, `💻 Coding`, `⚡ Embeddings (RAG)`;
  - Compute tiers: `🪶 Ultra-Light`, `🚀 Balanced`, `⚡ Flagship (8 GB GPU)`, `🧠 Expert`;
  - PC compatibility check (🟢 *100% GPU*, 🟡 *RAM Offload*, 🔴 *OOM*).
- [x] **Interactive Generation Succession Tree**:
  - Version lineage display (Qwen 2.5 ➔ 3.0 ➔ 3.5 ➔ 3.8);
  - Parameter growth warnings (e.g., upgrading from 2B to 4B with higher VRAM requirements);
  - Automatic update checks against the official Ollama library.

### 3. Local RAG (Knowledge Base)
- [x] Custom zero-dependency vector engine:
  - Text chunking with sliding overlap;
  - Chunk vectorization via local embedding models (`qwen3-embedding` / `nomic-embed-text`);
  - Semantic retrieval via Cosine Similarity;
  - Text extraction from `.txt`, `.md`, `.js`, `.py`, `.json`, `.csv`, `.log`, plus Word `.docx` without third-party npm packages;
  - Source citations with relevance scores in the UI.

### 4. Live Internet Access (Web Search & Fetch)
- [x] Built-in search parser powered by DuckDuckGo (HTML / Lite) with no paid API keys required.
- [x] Direct web page reader and cleaner (strips scripts, styles, navigation, and ads).
- [x] Automatic recognition of search intent and injection of fresh facts with current date and `[1]`, `[2]` footnotes.

### 5. SQLite Database & Session Management (ChatGPT / Claude Style)
- [x] Embedded database using native `node:sqlite` in WAL mode:
  - Tables: `chats`, `messages`, `chat_summaries`, `user_memories`, `settings`.
  - Reliable persistence of full conversation history across page reloads (`F5`).
- [x] Multi-session sidebar:
  - Time-grouped conversations: *Today*, *Yesterday*, *Previous 7 Days*, *Earlier*;
  - New chat shortcut (`Ctrl+N`), sidebar toggle (`Ctrl+B`);
  - Live keyword search across conversations;
  - Title editing and deletion with confirmation;
  - Automatic concise title generation from the first prompt.

### 6. Context Manager & Smart Compaction (Magic Compact)
- [x] **Claude-style Context Window Bar**:
  - Color-coded utilization indicator: 🟢 Green (<60%), 🟡 Amber (60–84%), 🔴 Red (≥85%);
  - Accurate token accounting by capturing `prompt_eval_count` and `eval_count` from the Ollama stream;
  - Display of model limit (32K tokens) and remaining free context space.
- [x] **🪄 Magic "Compact Context" Button**:
  - Fast essence extraction via local `qwen3:0.6b` (1–2 seconds);
  - Structured summary (*Goal, Decisions, Entities & Paths, Current Status*) saving over 80% of tokens;
  - Compacted context card rendered inline in conversation history;
  - Optional **Auto-Compact (>85%)** toggle to prevent degradation in long sessions.

---

## 🧭 Next Milestones

> [!NOTE]
> **Interface Principle**: The system focuses strictly on a **text and visual interface** (voice input and microphone support are excluded to keep speed, clarity, and focus on code and text at maximum).

---

### Stage 1: Autonomous AI Agent (Deep Research & ReAct Loop) 🕵️
* [x] **Multi-Step Agent Loop (ReAct: Reason + Act)**:
  * Implemented a native agent loop on the backend (`/api/agent/research`): the model formulates a hypothesis (`Thought`), calls web search and page fetch tools (`Action`), evaluates retrieved data (`Observation`), and synthesizes a final analytical report.
* [x] **Generation Mode Switcher in UI**:
  * Header toggle `🕵️ Deep Research: ВКЛ/ВЫКЛ` (ON/OFF) with neon cyberpunk highlight and `localStorage` persistence.
  * `⚡ Quick Answer` — standard single-pass generation (or quick web search/RAG);
  * `🕵️ Deep Research` — launches the autonomous multi-source agent to compile a detailed report.
* [x] **Interactive Agent Reasoning Log**:
  * Step-by-step trace displayed in chat: model hypotheses, generated search queries, fetched URLs, and extracted text size.
  * Reasoning chain persisted in SQLite for replay when revisiting past conversations.

---

### Stage 2: Support for Alternative & Uncensored Models (Uncensored & Open-Weights) 🔓 ✅ (Completed)
* [x] **Uncensored Model Lineups**:
  * Added verified catalog entries and classification for the **Dolphin** series (`dolphin3:8b`, `dolphin-mistral:7b`, `CognitiveComputations/dolphin-2.9.2-qwen2-7b:Q4_K_M`) alongside `dolphin-llama3` support;
  * Added **`abliterated`** / uncensored classification, filtering, and hardware compatibility checks preserving full capability without refusal guardrails.
* [x] **Expanded Third-Party Model Families (`🌐 Open-Weights`)**:
  * **Llama 3.1 / 3.2** (`llama3.2:3b`, `llama3.1:8b`);
  * **DeepSeek-R1** (`deepseek-r1:7b`, `deepseek-r1:8b`, `deepseek-r1:14b` with RAM Offload detection);
  * **Mistral / Mixtral** (`mistral:7b`, `mixtral:8x7b` flagged as `oom_risk` on 8 GB VRAM + 16 GB RAM) and **Google Gemma 2** (`gemma2:2b`, `gemma2:9b` with native `8192` context limit);
  * Unified hardware compatibility evaluation (`computeHardwareCompatibility`) across installed model classification and the download catalog.
* [x] **System Prompt Profiles (Persona & System Prompts)**:
  * Added SQLite `personas` table, per-chat `persona_id` persistence, 5 built-in profiles (*Default Assistant*, *Strict Programmer*, *Uncensored / No Lectures*, *Data Analyst*, *Reverse Engineer*), REST CRUD endpoints (`/api/personas*`), and UI selector/management modal wired into standard chat, Web/RAG prompts, and Deep Research synthesis.

---

### Stage 2.5: Smart Chat & Agent File Tools (Read-Only Sandbox) 📁
* [x] **Model- and Mode-Aware System Prompt (`buildSystemPrompt`)**:
  * Normal chat receives a dynamic `system` prompt specifying the active Ollama model name and an honest statement of mode capabilities/limitations (pointing users to `🕵️ Deep Research` or a `C:\...` path for project inspection).
* [x] **Smart Web Search Gate (`shouldWebSearch`)**:
  * Greetings, small talk, and meta-questions about the assistant no longer trigger unnecessary DuckDuckGo searches.
* [x] **Sandboxed ReAct Agent File Tools (`list_dir` and `read_file`)**:
  * Added read-only `list_dir` and `read_file` tools inside `/api/agent/research` restricted by allowed roots (`AGENT_FS_ROOTS`, defaulting to `Desktop`), validated via `realpathSync`, skipping binary files, heavy directories, and sensitive files.
  * Fixed `parseReActOutput` parser to capture `Action:` arguments accurately and raised the Ollama timeout to 120s.
* [x] **Local Path Auto-Routing & File UI (`hasLocalWindowsPath`)**:
  * Prompts containing a local Windows path automatically route to the autonomous ReAct agent; inspected directories and local files are displayed in the reasoning steps and sources block.

---

### Stage 3: GGUF & Hugging Face Model Importer 📦 ✅ (Completed)
* [x] **`models/` Directory Support for `.gguf` Files**:
  * Added local `models/` directory scanning (`GET /api/models/gguf`), streaming multi-GB `.gguf` drag-and-drop uploads (`POST /api/models/gguf/upload` with `GGUF` magic-header validation and zero full-file RAM buffering via `XMLHttpRequest` progress), hardware compatibility evaluation, and safe deletion (`DELETE /api/models/gguf/:filename`).
* [x] **Automatic `Modelfile` Generator & Ollama Importer**:
  * Implemented lightweight binary GGUF v2/v3 metadata header reader (`readGgufMetadata`), automatic prompt template detection (`detectTemplateFamily`: `chatml`, `llama3`, `mistral`, `gemma2`, `deepseek`, `auto`), interactive `Modelfile` preview (`POST /api/models/gguf/modelfile-preview`), and streaming `ollama create <name> -f <Modelfile>` execution (`POST /api/models/gguf/import`).
* [x] **Direct URL Downloader (Hugging Face)**:
  * Added Hugging Face repo/URL inspector (`GET /api/models/hf/inspect` with `/blob/` -> `/resolve/` rewrite, recursive quantization discovery, and private/loopback SSRF blocking) plus background streaming `.gguf` downloader with live progress/speed tracking and cancellation (`POST /api/models/hf/download`, `GET /api/models/hf/status`, `POST /api/models/hf/cancel`).

---

### Stage 4: Local Code Execution Sandbox 💻
* [ ] **Safe JavaScript / Python Interpreter**:
  * Isolated execution of model-generated calculation and data-processing scripts.
* [ ] **Inline Result Visualization**:
  * Charts (SVG / Canvas), mathematical computations, CSV/JSON table parsing, and instant inline rendering.

---

### Stage 5: Advanced Multimodal Support for Vision Models 👁️
* [ ] **Drag-and-Drop Images & Screenshots**:
  * Clipboard image paste (`Ctrl+V`) or drag-and-drop directly into the chat input box.
* [ ] **Document & Diagram Parsing via `qwen3-vl:4b`**:
  * Recognition of charts, scanned tables, flowcharts, and architecture diagrams into structured text or code.

---

## 🛠️ Project Architectural Principles

1. **Zero-Dependencies**: No bloated external npm packages, Python wrappers, or heavy frameworks — only pure native Node.js and Vanilla JS.
2. **100% Local & Private**: No user data leaves the local machine for third-party cloud APIs. All inference, embeddings, and databases run locally.
3. **Hardware Efficiency**: Optimized for consumer 8 GB VRAM GPUs (AMD Radeon RX 6600) through smart context management (32K), lightweight helper models (0.6B), and memory discipline.
4. **Text-First Focus**: Clean, fast, responsive interface dedicated strictly to text, knowledge, search, and code.
