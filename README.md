# Office AI Add-in

An AI assistant that runs inside Microsoft Word, Excel, and PowerPoint.
Supports **local AI** (Ollama) and **cloud AI** (OpenAI, Anthropic, Gemini).

---

## Project Structure

```
office-ai-addin/
├── backend/
│   ├── main.py            ← Python FastAPI server (the brain)
│   └── requirements.txt   ← Python dependencies
└── addin/
    ├── taskpane.html      ← The UI inside Office
    └── manifest.xml       ← Tells Office about your add-in
```

---

## Quick Start

### 1. Install Python dependencies

```bash
cd backend
pip install -r requirements.txt
```

### 2. Start the backend server

```bash
cd backend
python main.py
```

The server runs at **http://localhost:8000**
- Visit http://localhost:8000/health to confirm it's working
- The add-in UI is served at http://localhost:8000/addin/taskpane.html

### 3. (Optional) Set up Local AI with Ollama

Download Ollama from https://ollama.com and then run:

```bash
ollama pull llama3        # download a model (~4GB)
ollama serve              # start the local AI server
```

The backend will auto-detect it. You can switch between local/cloud in the add-in UI.

### 4. Load the add-in into Office

**For development (easiest):**
1. Open Word (or Excel/PowerPoint)
2. Go to **Insert → Add-ins → My Add-ins → Upload My Add-in**
3. Select `addin/manifest.xml`
4. The "Office AI" panel will appear in the right sidebar

**Alternative — sideload via network share:**
- Put the `addin/` folder on a shared network path
- In Office → Trust Center → Trusted Add-in Catalogs → add the path

---

## Using the Add-in

### Chat Tab
- Type any prompt and hit **Send**
- Use **"Use Selection"** to pull selected text from your document into the prompt
- Switch between **Cloud AI** and **Local AI** with the toggle at the top

### Tools Tab
One-click actions on selected text:
- **Summarize** — get a quick summary
- **Fix Grammar** — clean up spelling/grammar
- **Make Formal / Casual** — change tone
- **Simplify / Expand** — rewrite for clarity or detail
- **Translate** — translate to Spanish (edit the code to change language)
- **Bullet Points** — extract key points

After any result, use **Insert ↓** to replace your selection in the document.

### Settings Tab
1. Pick your **AI provider** (OpenAI, Anthropic, or Gemini)
2. Paste your **API key**
3. Optionally override the **model** (e.g. `gpt-4o`, `claude-opus-4-6`)
4. Hit **Save Settings**

For local AI, click **Detect Local Models** to see what Ollama models you have installed.

---

## API Endpoints (for reference)

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Check if backend is running |
| POST | `/ai/generate` | Generate AI response |
| GET | `/settings` | Get current settings |
| POST | `/settings` | Save settings |
| GET | `/local/models` | List Ollama models |

---

## Common Issues

**"Ollama not running"** — run `ollama serve` in a terminal first.

**"No API key configured"** — go to the Settings tab and paste your key.

**Office won't load the add-in** — make sure `python main.py` is running, then reload the add-in.

**CORS errors in browser** — the backend has CORS open for development; it's fine for localhost.

---

## Next Steps

- [ ] Package with **PyInstaller** so users don't need Python installed
- [ ] Add HTTPS with `mkcert` for production Office deployment
- [ ] Build a website / landing page for distribution
- [ ] Publish to Microsoft AppSource (requires Microsoft Partner account)
