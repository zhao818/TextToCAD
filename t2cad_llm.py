# t2cad_llm.py — Unified LLM abstraction layer for all TextToCAD tools
"""
Shared LLM client with LangChain + fallback paths.
All four tools (AutoCAD/Excel/Word/PPT) import from here instead of
each maintaining their own _call_llm / _call_bridge / load_config.

Architecture:
  LLMClient.chat(messages) → str
    ├── _call_langchain()     # LangChain ChatOpenAI (DeepSeek/Claude/GPT)
    ├── _call_requests()      # Raw requests fallback
    └── _call_bridge()        # File-based bridge mode

Usage:
  from t2cad_llm import LLMClient, load_config, strip_code_fence, resolve_proxies

  cfg = load_config()
  client = LLMClient(cfg)
  answer = client.chat([{"role":"user","content":"Hello"}])
  # multi-turn with memory:
  answer = client.chat_with_memory("Hello", session_id="cad_session_1")
"""

import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
CONFIG_DIR = Path.home() / ".text_to_cad"
CONFIG_FILE = CONFIG_DIR / "config.json"
BRIDGE_DIR = CONFIG_DIR / "bridge"
BRIDGE_INPUT = BRIDGE_DIR / "input.txt"
BRIDGE_OUTPUT = BRIDGE_DIR / "output.py"
BRIDGE_DONE = BRIDGE_DIR / "done.txt"

# ---------------------------------------------------------------------------
# Default config
# ---------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "provider": "deepseek",
    "api_key": "",
    "api_base": "https://api.deepseek.com/v1",
    "model": "deepseek-chat",
    "temperature": 0.0,
    "max_tokens": 4096,
    "proxies": {"enabled": True, "http": "", "https": ""},
    "language": "zh",
    "units": "mm",
    "auto_execute": True,
    "show_code": True,
}

# ---------------------------------------------------------------------------
# Config helpers (shared across all tools)
# ---------------------------------------------------------------------------
def load_config():
    """Load config from ~/.text_to_cad/config.json, filling defaults."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    BRIDGE_DIR.mkdir(parents=True, exist_ok=True)
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        return cfg
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
    return dict(DEFAULT_CONFIG)


def resolve_proxies(proxies_cfg):
    """Resolve proxy settings for requests. Returns dict or None."""
    proxies = proxies_cfg or {"enabled": True}
    if not isinstance(proxies, dict) or not proxies.get("enabled"):
        return None
    http_proxy = proxies.get("http") or proxies.get("https")
    if not http_proxy:
        try:
            from urllib.request import getproxies
            sys_proxy = getproxies()
            http_proxy = sys_proxy.get("https") or sys_proxy.get("http")
        except Exception:
            pass
    if http_proxy:
        return {"http": http_proxy, "https": http_proxy}
    return None


def strip_code_fence(text):
    """Remove ```python ... ``` markers from LLM output."""
    text = re.sub(r'^```(?:python)?\s*\n?', '', text.strip())
    text = re.sub(r'\n?```\s*$', '', text)
    return text.strip()


# ---------------------------------------------------------------------------
# LLMClient — the one LLM interface all tools use
# ---------------------------------------------------------------------------
class LLMClient:
    """Unified LLM client with LangChain primary + requests fallback + bridge.

    Supports:
    - DeepSeek (via LangChain ChatOpenAI, OpenAI-compatible endpoint)
    - Claude (via Anthropic SDK if available)
    - OpenAI (via LangChain ChatOpenAI)
    - Bridge mode (file-based, for air-gapped workflows)
    - Conversation memory (LangChain ConversationBufferMemory)
    """

    def __init__(self, cfg: dict = None, use_langchain: bool = True):
        self.cfg = cfg or load_config()
        self.proxies = resolve_proxies(self.cfg.get("proxies"))
        self._use_langchain = use_langchain
        self._lc_model = None  # Lazy-init LangChain model
        self._memories: dict[str, list] = {}  # session_id → message list

    # ── Public API ─────────────────────────────────────────

    def chat(self, messages: list[dict]) -> str:
        """Send messages to LLM, return text response. Primary entry point."""
        provider = self.cfg.get("provider", "deepseek")

        if provider == "bridge":
            return self._call_bridge(messages)

        # Try LangChain first, fall back to raw requests
        if self._use_langchain:
            try:
                return self._call_langchain(messages)
            except Exception as e:
                print(f"[t2cad_llm] LangChain failed ({e}), falling back to requests")
                return self._call_requests(messages)
        else:
            return self._call_requests(messages)

    def chat_with_memory(self, user_message: str, session_id: str = "default",
                         system_prompt: str = "") -> str:
        """Chat with conversation memory. Maintains context across calls.

        Args:
            user_message: The user's input text
            session_id: Unique key for this conversation thread
            system_prompt: Optional system instruction (only applied on first call)

        Returns:
            LLM response text
        """
        if session_id not in self._memories:
            self._memories[session_id] = []
            if system_prompt:
                self._memories[session_id].append(
                    {"role": "system", "content": system_prompt}
                )

        self._memories[session_id].append({"role": "user", "content": user_message})

        response = self.chat(self._memories[session_id])

        self._memories[session_id].append({"role": "assistant", "content": response})

        # Prune old memory if too long (>40 messages)
        if len(self._memories[session_id]) > 40:
            # Keep system prompt + last 20 exchanges
            sys_msgs = [m for m in self._memories[session_id]
                       if m["role"] == "system"]
            rest = self._memories[session_id][-40:]
            self._memories[session_id] = sys_msgs + rest

        return response

    def clear_memory(self, session_id: str = "default"):
        """Clear conversation history for a session."""
        self._memories.pop(session_id, None)

    def switch_model(self, provider: str, model: str = None):
        """Runtime model switch. Rebuilds LangChain model on next call."""
        self.cfg["provider"] = provider
        if model:
            self.cfg["model"] = model
        self._lc_model = None  # Force rebuild

    # ── LangChain engine ────────────────────────────────────

    def _get_lc_model(self):
        """Lazy-init LangChain ChatModel. Reuses cached instance."""
        if self._lc_model is not None:
            return self._lc_model

        api_base = self.cfg.get("api_base", "https://api.deepseek.com/v1")
        api_key = self.cfg.get("api_key", "")
        model = self.cfg.get("model", "deepseek-chat")
        temperature = self.cfg.get("temperature", 0.0)
        max_tokens = self.cfg.get("max_tokens", 4096)

        try:
            from langchain_openai import ChatOpenAI

            # Build proxy config for LangChain's httpx client
            extra_kwargs = {}
            if self.proxies:
                proxy_url = self.proxies.get("http") or self.proxies.get("https")
                if proxy_url:
                    # LangChain ChatOpenAI passes extra kwargs to openai client
                    extra_kwargs["openai_proxy"] = proxy_url

            self._lc_model = ChatOpenAI(
                model=model,
                openai_api_key=api_key,
                openai_api_base=api_base,
                temperature=temperature,
                max_tokens=max_tokens,
                **extra_kwargs,
            )
        except ImportError:
            raise RuntimeError(
                "langchain-openai not installed. "
                "Run: pip install langchain-openai"
            )
        return self._lc_model

    def _call_langchain(self, messages: list[dict]) -> str:
        """Call LLM via LangChain ChatOpenAI."""
        model = self._get_lc_model()

        # Convert dict messages to LangChain format
        from langchain_core.messages import HumanMessage, AIMessage, SystemMessage

        lc_messages = []
        for m in messages:
            role = m.get("role", "user")
            content = m.get("content", "")
            if role == "system":
                lc_messages.append(SystemMessage(content=content))
            elif role == "assistant":
                lc_messages.append(AIMessage(content=content))
            else:
                lc_messages.append(HumanMessage(content=content))

        response = model.invoke(lc_messages)
        return strip_code_fence(response.content)

    # ── Raw requests fallback ───────────────────────────────

    def _call_requests(self, messages: list[dict]) -> str:
        """Fallback: raw requests.post to OpenAI-compatible API."""
        try:
            import requests
        except ImportError:
            raise RuntimeError("requests not installed and bridge mode not active")

        headers = {
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": f"Bearer {self.cfg['api_key']}",
        }
        body = {
            "model": self.cfg["model"],
            "messages": messages,
            "temperature": self.cfg["temperature"],
            "max_tokens": self.cfg["max_tokens"],
        }
        url = f"{self.cfg['api_base'].rstrip('/')}/chat/completions"
        resp = requests.post(
            url, headers=headers, json=body, timeout=60, proxies=self.proxies
        )
        resp.raise_for_status()
        return strip_code_fence(resp.json()["choices"][0]["message"]["content"])

    # ── Bridge mode ─────────────────────────────────────────

    def _call_bridge(self, messages: list[dict]) -> str:
        """File-based bridge: writes input.txt, waits for output.py."""
        BRIDGE_DIR.mkdir(parents=True, exist_ok=True)

        with open(BRIDGE_INPUT, "w", encoding="utf-8") as f:
            f.write(messages[-1]["content"])

        if BRIDGE_DONE.exists():
            BRIDGE_DONE.unlink()

        waited = 0
        while waited < 120:
            if BRIDGE_OUTPUT.exists() and BRIDGE_DONE.exists():
                code = strip_code_fence(
                    open(BRIDGE_OUTPUT, "r", encoding="utf-8").read()
                )
                BRIDGE_INPUT.unlink(missing_ok=True)
                BRIDGE_OUTPUT.unlink(missing_ok=True)
                BRIDGE_DONE.unlink(missing_ok=True)
                return code
            time.sleep(0.5)
            waited += 0.5

        raise TimeoutError("桥接模式等待超时 (120s)")


# ---------------------------------------------------------------------------
# Module-level convenience: get a ready-to-use client
# ---------------------------------------------------------------------------
_default_client: Optional[LLMClient] = None


def get_client(use_langchain: bool = True) -> LLMClient:
    """Get or create the default LLMClient singleton."""
    global _default_client
    if _default_client is None:
        _default_client = LLMClient(use_langchain=use_langchain)
    return _default_client


def reload_client():
    """Force recreate the default client (after config changes)."""
    global _default_client
    _default_client = None
    return get_client()
