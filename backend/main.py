"""
Office AI Add-in - Python Backend
Run with: uvicorn main:app --reload --port 8000
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
import httpx
import json
import os
import sys
import subprocess
import importlib.util

app = FastAPI(title="Office AI Add-in Backend", version="1.0.0")

# Allow the Office add-in (localhost) to call this backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, lock this to your add-in's origin
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Models ────────────────────────────────────────────────────────────────────

class AIRequest(BaseModel):
    prompt: str
    mode: str = "cloud"          # "cloud" or "local"
    provider: str = "openai"     # "openai" | "anthropic" | "gemini"
    api_key: Optional[str] = None
    model: Optional[str] = None  # override default model
    system_prompt: Optional[str] = None
    max_tokens: int = 1024

class SettingsPayload(BaseModel):
    provider: str
    api_key: str
    model: Optional[str] = None

# ─── In-memory settings store (replace with SQLite or a config file later) ────

settings_store: dict = {
    "provider": "openai",
    "api_key": "",
    "model": "",
    "local_model": "llama3",     # default Ollama model
}

# ─── Health check ──────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok", "version": "1.0.0"}

# ─── Settings endpoints ────────────────────────────────────────────────────────

@app.get("/settings")
def get_settings():
    # Never send the full API key back to the client
    safe = settings_store.copy()
    key = safe.get("api_key", "")
    safe["api_key"] = f"{'*' * (len(key) - 4)}{key[-4:]}" if len(key) > 4 else ("set" if key else "")
    return safe

@app.post("/settings")
def save_settings(payload: SettingsPayload):
    settings_store["provider"] = payload.provider
    settings_store["api_key"] = payload.api_key
    if payload.model:
        settings_store["model"] = payload.model
    return {"status": "saved"}

# ─── Local AI (Ollama) ─────────────────────────────────────────────────────────

def check_ollama_running() -> bool:
    """Check if Ollama is running locally."""
    try:
        import urllib.request
        urllib.request.urlopen("http://localhost:11434", timeout=2)
        return True
    except Exception:
        return False

async def call_local_ai(prompt: str, system_prompt: str, model: str, max_tokens: int) -> str:
    """Call Ollama local AI server."""
    if not check_ollama_running():
        raise HTTPException(
            status_code=503,
            detail="Ollama is not running. Start it with: ollama serve"
        )

    model = model or settings_store.get("local_model", "llama3")
    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": prompt})

    async with httpx.AsyncClient(timeout=120.0) as client:
        response = await client.post(
            "http://localhost:11434/api/chat",
            json={
                "model": model,
                "messages": messages,
                "stream": False,
                "options": {"num_predict": max_tokens},
            },
        )
        response.raise_for_status()
        data = response.json()
        return data["message"]["content"]

# ─── Cloud AI providers ────────────────────────────────────────────────────────

async def call_openai(prompt: str, system_prompt: str, api_key: str, model: str, max_tokens: int) -> str:
    model = model or "gpt-4o-mini"
    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": prompt})

    async with httpx.AsyncClient(timeout=60.0) as client:
        response = await client.post(
            "https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": model, "messages": messages, "max_tokens": max_tokens},
        )
        if response.status_code == 401:
            raise HTTPException(status_code=401, detail="Invalid OpenAI API key.")
        response.raise_for_status()
        data = response.json()
        return data["choices"][0]["message"]["content"]

async def call_anthropic(prompt: str, system_prompt: str, api_key: str, model: str, max_tokens: int) -> str:
    model = model or "claude-haiku-4-5-20251001"
    payload = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": prompt}],
    }
    if system_prompt:
        payload["system"] = system_prompt

    async with httpx.AsyncClient(timeout=60.0) as client:
        response = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json",
            },
            json=payload,
        )
        if response.status_code == 401:
            raise HTTPException(status_code=401, detail="Invalid Anthropic API key.")
        response.raise_for_status()
        data = response.json()
        return data["content"][0]["text"]

async def call_gemini(prompt: str, system_prompt: str, api_key: str, model: str, max_tokens: int) -> str:
    model = model or "gemini-1.5-flash"
    full_prompt = f"{system_prompt}\n\n{prompt}" if system_prompt else prompt

    async with httpx.AsyncClient(timeout=60.0) as client:
        response = await client.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}",
            json={
                "contents": [{"parts": [{"text": full_prompt}]}],
                "generationConfig": {"maxOutputTokens": max_tokens},
            },
        )
        if response.status_code == 400:
            raise HTTPException(status_code=401, detail="Invalid Gemini API key.")
        response.raise_for_status()
        data = response.json()
        return data["candidates"][0]["content"]["parts"][0]["text"]

# ─── Main AI endpoint ──────────────────────────────────────────────────────────

@app.post("/ai/generate")
async def generate(req: AIRequest):
    """
    Single endpoint for all AI calls.
    mode = "local"  → Ollama
    mode = "cloud"  → uses saved or provided API key + provider
    """
    system = req.system_prompt or "You are a helpful AI assistant inside Microsoft Office. Be concise and clear."

    try:
        if req.mode == "local":
            result = await call_local_ai(req.prompt, system, req.model or "", req.max_tokens)

        else:  # cloud
            api_key = req.api_key or settings_store.get("api_key", "")
            provider = req.provider or settings_store.get("provider", "openai")

            if not api_key:
                raise HTTPException(status_code=400, detail="No API key configured. Go to Settings.")

            if provider == "openai":
                result = await call_openai(req.prompt, system, api_key, req.model or "", req.max_tokens)
            elif provider == "anthropic":
                result = await call_anthropic(req.prompt, system, api_key, req.model or "", req.max_tokens)
            elif provider == "gemini":
                result = await call_gemini(req.prompt, system, api_key, req.model or "", req.max_tokens)
            else:
                raise HTTPException(status_code=400, detail=f"Unknown provider: {provider}")

        return {"result": result, "mode": req.mode}

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ─── Ollama model list ─────────────────────────────────────────────────────────

@app.get("/local/models")
async def list_local_models():
    """Returns available Ollama models on this machine."""
    if not check_ollama_running():
        return {"models": [], "error": "Ollama not running"}
    try:
        async with httpx.AsyncClient(timeout=10.0) as client:
            response = await client.get("http://localhost:11434/api/tags")
            data = response.json()
            models = [m["name"] for m in data.get("models", [])]
            return {"models": models}
    except Exception as e:
        return {"models": [], "error": str(e)}

# ─── Serve the add-in HTML (optional – for local dev) ─────────────────────────

addin_path = os.path.join(os.path.dirname(__file__), "..", "addin")
if os.path.isdir(addin_path):
    app.mount("/addin", StaticFiles(directory=addin_path, html=True), name="addin")

# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
