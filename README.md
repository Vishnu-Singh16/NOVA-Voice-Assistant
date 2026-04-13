# 🎙️ NOVA – Voice-Controlled AI Desktop Assistant

> A secure, AI-powered desktop assistant for Windows that enables hands-free system control, intelligent automation, and real-time web & media interaction using natural language.

---

## 📌 Overview

NOVA is a voice-controlled AI assistant designed to bridge the gap between human language and system-level operations. It allows users to interact with their computer using speech or text to perform tasks like file management, navigation, application control, web search, and media playback.

Built with a modular architecture and powered by LLM-based intent routing, NOVA combines usability, automation, and security into a single intelligent system.

---

## 🚀 Key Features

### 🎤 Voice Control & Wake Word
- Natural language command execution via speech
- Always-on wake word detection: **“OK NOVA”**
- Background listening with cooldown handling
- Works with both voice and typed input

---

### 🧠 AI-Powered Command Routing
- Uses Groq LLM for intent detection and argument extraction
- Dynamically maps user commands to system functions
- Handles both structured commands and casual conversations

---

### 📂 File Management
- Create, delete, and rename files via voice
- Interactive prompts for file name and extension
- Safe execution with confirmation layer

---

### 📁 Smart Directory Navigation
- Navigate folders using natural language
- Fuzzy matching for folder names
- Global folder search with teleportation
- Supports shortcuts like Desktop, Downloads, Documents

---

### 🚀 Application & Document Control
- Launch and close applications by name
- Open documents (PDF, DOCX, PPT, TXT, etc.)
- Intelligent matching with multiple suggestions

---

### 🌐 Web Search & Browser Automation
- Perform Google searches using voice commands
- Opens results in Brave or default browser

**Examples:**
- “Search Cricbuzz”
- “Google Python tutorials”

---

### 🎵 YouTube & Media Playback
- Search and play videos directly from voice
- Automatically opens YouTube results

**Examples:**
- “Play Believer song”
- “Play lofi beats”

---

### ⏭️ Media Control (System-Wide)
- Control playback across YouTube, Spotify, VLC, etc.

**Supported Commands:**
- Play / Pause  
- Next / Previous  
- Volume Up / Down  
- Mute  

**Examples:**
- “Play first video”  
- “Next video”  
- “Increase volume by 10”  
- “Mute”  

---

### 🔧 Environment Variable Manager
- Set system environment variables via guided interaction
- Automatically updates PATH using Windows Registry

---

### ⚡ Power Control
- Shutdown / Restart system
- Schedule delayed actions
- Cancel pending operations

---

### 🖥️ Premium Desktop GUI
- Built with CustomTkinter
- Dark-themed modern interface
- Chat-style interaction UI
- Sidebar with quick actions and system status

---

### 🔍 Background System Indexing
- Indexes apps, folders, and documents on startup
- Enables fast search and intelligent matching
- Runs asynchronously without blocking UI

---

## 🧱 Architecture

GUI (CustomTkinter)
↓
Core Logic (assistant_core.py)
↓
AI Routing (Groq API)
↓
System Execution (OS / Windows APIs)


---

## 🛠️ Tech Stack

- **Language:** Python  
- **AI/NLP:** Groq API (LLM-based intent routing)  
- **Speech Recognition:** Google Speech API  
- **Text-to-Speech:** Windows SAPI (pyttsx3)  
- **GUI:** CustomTkinter  
- **System Control:** os, subprocess, winreg, ctypes  
- **Audio Handling:** PyAudio  
- **Utilities:** difflib, send2trash  

---

## ▶️ How to Run

```bash
# 1. Create virtual environment
python -m venv venv

# 2. Activate it
venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run GUI
python gui.py