import glob
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from dotenv import load_dotenv

load_dotenv()

from llm_clients.gigachat import GigaChatClient

PROMPT = """Перед тобой рендер слайда презентации.
Определи архетип слайда (одно из: title, section_divider, bullet_list, stats_kpi,
two_column_comparison, image_caption, quote, agenda, closing, other).
Затем перечисли визуальные регионы на слайде и их роль (заголовок, буллет-лист,
цифра-метрика, колонка сравнения и т.д.) с примерным расположением (верх/центр/низ, лево/право).
Ответь кратко, строго по структуре:
Архетип: ...
Регионы:
- ...
"""

MODELS_TO_TRY = ["GigaChat-2-Max", "GigaChat-2-Pro"]


def main():
    client = GigaChatClient(verify_ssl=False)
    default_dir = os.path.join(os.path.dirname(__file__), "..", "output", "test_slides")
    images_dir = sys.argv[1] if len(sys.argv) > 1 else default_dir
    images = sorted(glob.glob(os.path.join(images_dir, "*.png")))

    working_model = None
    for model in MODELS_TO_TRY:
        try:
            client.ask_about_image(images[0], PROMPT, model=model)
            working_model = model
            break
        except Exception as e:
            print(f"Model {model} failed: {e}")

    if not working_model:
        print("No vision-capable model worked, stopping.")
        return

    print(f"Using model: {working_model}\n")
    for img_path in images:
        print(f"=== {os.path.basename(img_path)} ===")
        result = client.ask_about_image(img_path, PROMPT, model=working_model)
        print(result["choices"][0]["message"]["content"])
        print()


if __name__ == "__main__":
    main()
