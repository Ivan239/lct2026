import base64
import os
import time
import uuid

import requests

OAUTH_URL = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
# api.giga.chat hosts the newer model line (GigaChat-2*) — the older
# gigachat.devices.sberbank.ru host doesn't serve them. OAuth stays on the
# ngw host regardless; it's a separate auth service, not tied to API_BASE.
API_BASE = "https://api.giga.chat/v1"

# (connect, read) timeouts. Without them a hung socket — routine on the flaky
# tunnel these endpoints are reached through — blocks the request forever, and
# the retry logic in common/model_fallback never even gets an exception to
# retry on. requests.exceptions.Timeout is already in its transient-error set,
# so timing out here turns a permanent hang into a retried call.
FAST_TIMEOUT = (10, 30)   # oauth, models, balance
CHAT_TIMEOUT = (10, 180)  # completions can legitimately think for a while
UPLOAD_TIMEOUT = (10, 60)


class GigaChatClient:
    def __init__(self, client_id=None, client_secret=None, scope=None, verify_ssl=True):
        self.client_id = client_id or os.environ["GIGACHAT_CLIENT_ID"]
        self.client_secret = client_secret or os.environ["GIGACHAT_CLIENT_SECRET"]
        self.scope = scope or os.environ.get("GIGACHAT_SCOPE", "GIGACHAT_API_PERS")
        self.verify_ssl = verify_ssl
        self._token = None
        self._token_expires_at = 0

    def _get_access_token(self):
        auth_key = base64.b64encode(
            f"{self.client_id}:{self.client_secret}".encode()
        ).decode()
        resp = requests.post(
            OAUTH_URL,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
                "RqUID": str(uuid.uuid4()),
                "Authorization": f"Basic {auth_key}",
            },
            data={"scope": self.scope},
            verify=self.verify_ssl,
            timeout=FAST_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
        self._token = data["access_token"]
        # expires_at is epoch ms in GigaChat's response
        self._token_expires_at = data["expires_at"] / 1000
        return self._token

    def _token_valid(self):
        return self._token and time.time() < self._token_expires_at - 30

    def access_token(self):
        if not self._token_valid():
            self._get_access_token()
        return self._token

    def list_models(self):
        resp = requests.get(
            f"{API_BASE}/models",
            headers={"Accept": "application/json", "Authorization": f"Bearer {self.access_token()}"},
            verify=self.verify_ssl,
            timeout=FAST_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()

    def get_balance(self):
        """Remaining token balance per model tier — a free GET, no tokens spent
        checking it, so it's cheap to surface after every real request."""
        resp = requests.get(
            f"{API_BASE}/balance",
            headers={"Accept": "application/json", "Authorization": f"Bearer {self.access_token()}"},
            verify=self.verify_ssl,
            timeout=FAST_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()

    def chat(self, messages, model="GigaChat", **kwargs):
        resp = requests.post(
            f"{API_BASE}/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Authorization": f"Bearer {self.access_token()}",
            },
            json={"model": model, "messages": messages, **kwargs},
            verify=self.verify_ssl,
            timeout=CHAT_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()

    def upload_file(self, path, purpose="general"):
        mime = "image/png" if path.lower().endswith(".png") else "application/octet-stream"
        with open(path, "rb") as f:
            resp = requests.post(
                f"{API_BASE}/files",
                headers={"Accept": "application/json", "Authorization": f"Bearer {self.access_token()}"},
                data={"purpose": purpose},
                files={"file": (os.path.basename(path), f, mime)},
                verify=self.verify_ssl,
                timeout=UPLOAD_TIMEOUT,
            )
        resp.raise_for_status()
        return resp.json()

    def ask_about_image(self, image_path, prompt, model="GigaChat-2-Pro", **kwargs):
        file_obj = self.upload_file(image_path)
        return self.chat(
            [{"role": "user", "content": prompt, "attachments": [file_obj["id"]]}],
            model=model,
            **kwargs,
        )
