"""OpenAI-compatible chat client — the seam for "generate from different LLMs".

RTX (the local GPU box) and most self-hosted engines (vLLM, Ollama, LM Studio,
TGI) speak the OpenAI /chat/completions shape, which is already what GigaChat's
client returns too. So this client mirrors GigaChatClient.chat's return contract
({"choices":[{"message":{"content"}}], "usage"}) and the whole text pipeline —
structure-first classification, outline, blocks, style card — drives it unchanged.

Vision is intentionally not implemented here: the improvement loop keeps the JUDGE
on GigaChat (independent of whichever model generated the deck), and generation
uses structure-first text classification, so a text-only backend is sufficient.
If a backend later needs vision, add inline base64 image_url parts here.
"""

import os

import requests
from common.tls import allow_unverified_tls

CHAT_TIMEOUT = (10, 300)  # local models can be slow to first token


class OpenAICompatClient:
    def __init__(self, base_url, api_key=None, default_model=None, verify_ssl=True):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key or os.environ.get("OPENAI_COMPAT_API_KEY", "")
        self.default_model = default_model
        self.verify_ssl = verify_ssl
        if not verify_ssl:
            allow_unverified_tls(self.base_url)

    def _headers(self):
        h = {"Content-Type": "application/json", "Accept": "application/json"}
        if self.api_key:
            h["Authorization"] = f"Bearer {self.api_key}"
        return h

    def chat(self, messages, model=None, **kwargs):
        # GigaChat carries images as a per-message "attachments" list; a plain
        # OpenAI backend rejects that key, so strip it — this client is text-only.
        clean = [{k: v for k, v in m.items() if k != "attachments"} for m in messages]
        resp = requests.post(
            f"{self.base_url}/chat/completions",
            headers=self._headers(),
            json={"model": model or self.default_model, "messages": clean, **kwargs},
            verify=self.verify_ssl,
            timeout=CHAT_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()

    def list_models(self):
        resp = requests.get(f"{self.base_url}/models", headers=self._headers(),
                            verify=self.verify_ssl, timeout=(10, 30))
        resp.raise_for_status()
        return resp.json()

    def get_balance(self):
        # Local/self-hosted models don't meter tokens — report empty so the
        # balance-surfacing code degrades cleanly rather than erroring.
        return {"balance": []}

    def upload_file(self, path, purpose="general"):
        raise NotImplementedError("OpenAICompatClient is text-only; keep the judge on a vision backend")

    def ask_about_image(self, image_path, prompt, model=None, **kwargs):
        raise NotImplementedError("OpenAICompatClient is text-only; keep the judge on a vision backend")
