import os

from PIL import Image, ImageDraw, ImageFont

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "test_slides")
os.makedirs(OUT_DIR, exist_ok=True)

W, H = 1280, 720
NAVY = (24, 42, 74)
WHITE = (255, 255, 255)
ACCENT = (240, 140, 40)
GRAY = (90, 90, 90)


def font(size, bold=False):
    names = ["Arial Bold.ttf", "Arial.ttf"] if bold else ["Arial.ttf"]
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def title_slide():
    img = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(img)
    d.rectangle([0, H - 12, W, H], fill=ACCENT)
    d.text((100, 280), "Название продукта", font=font(64, bold=True), fill=WHITE)
    d.text((100, 370), "Питч для инвесторов — Q3 2026", font=font(28), fill=(200, 200, 210))
    img.save(os.path.join(OUT_DIR, "1_title.png"))


def bullets_slide():
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 100], fill=NAVY)
    d.text((60, 30), "Ключевые преимущества", font=font(36, bold=True), fill=WHITE)
    bullets = [
        "Автоматизация подготовки презентаций",
        "Сохранение фирменного стиля компании",
        "Сокращение времени с часов до минут",
        "Работа на российском LLM-стеке",
    ]
    y = 180
    for b in bullets:
        d.ellipse([70, y + 8, 82, y + 20], fill=ACCENT)
        d.text((100, y), b, font=font(28), fill=(40, 40, 40))
        y += 80
    img.save(os.path.join(OUT_DIR, "2_bullets.png"))


def stats_slide():
    img = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(img)
    d.text((60, 40), "Метрики за квартал", font=font(36, bold=True), fill=WHITE)
    stats = [("x10", "Ускорение подготовки"), ("3", "Шаблона в пилоте"), ("0", "Часов ручной вёрстки")]
    x = 100
    for num, label in stats:
        d.text((x, 260), num, font=font(72, bold=True), fill=ACCENT)
        d.text((x, 380), label, font=font(22), fill=(210, 210, 220))
        x += 380
    img.save(os.path.join(OUT_DIR, "3_stats.png"))


def two_column_slide():
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 100], fill=NAVY)
    d.text((60, 30), "Продукт vs Конкуренты", font=font(36, bold=True), fill=WHITE)
    d.rectangle([60, 140, 620, 620], outline=ACCENT, width=3)
    d.text((90, 160), "Наш подход", font=font(28, bold=True), fill=NAVY)
    d.rectangle([660, 140, 1220, 620], outline=GRAY, width=3)
    d.text((690, 160), "Обычные AI-тулы", font=font(28, bold=True), fill=GRAY)
    left = ["Учится на вашем шаблоне", "Точная фирменная стилистика", "RU LLM-стек"]
    right = ["Свои generic темы", "Без учёта бренда", "Зарубежные модели"]
    y = 240
    for l, r in zip(left, right):
        d.text((90, y), "- " + l, font=font(22), fill=(40, 40, 40))
        d.text((690, y), "- " + r, font=font(22), fill=(90, 90, 90))
        y += 70
    img.save(os.path.join(OUT_DIR, "4_two_column.png"))


if __name__ == "__main__":
    title_slide()
    bullets_slide()
    stats_slide()
    two_column_slide()
    print("Saved to", OUT_DIR)
