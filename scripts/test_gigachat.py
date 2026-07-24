import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from dotenv import load_dotenv

load_dotenv()

from llm_clients.gigachat import GigaChatClient


def main():
    client = GigaChatClient()
    try:
        models = client.list_models()
    except Exception as e:
        print(f"SSL/connection error with verify=True: {e}")
        print("Retrying with verify_ssl=False (common for GigaChat's cert chain)...")
        client = GigaChatClient(verify_ssl=False)
        models = client.list_models()

    print("Models available:")
    for m in models.get("data", []):
        print(" -", m["id"])

    print("\nTest chat completion:")
    result = client.chat([{"role": "user", "content": "Скажи одно слово: работает"}])
    print(result["choices"][0]["message"]["content"])


if __name__ == "__main__":
    main()
