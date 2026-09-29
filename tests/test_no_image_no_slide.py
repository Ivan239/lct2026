"""Нет картинки — нет и слайда под неё.

Сервис не генерирует изображения (задача со звёздочкой, только топ-10), и
слайд «картинка + подпись» без картинки выходит пустой рамкой «ИЗОБРАЖЕНИЕ»:
в девяти колодах сдачи таких было девять. Раньше код ВСТАВЛЯЛ такой слайд в
каждую колоду (модель сама просила его через раз)."""

from content_parser.two_phase import _enforce_outline_rules


def _roles(outline, **kw):
    return [i["role"] for i in _enforce_outline_rules(outline, **kw)]


OUTLINE = [{"role": "title", "theme": "t", "count": None},
           {"role": "bullet_list", "theme": "b", "count": 3},
           {"role": "image_caption", "theme": "скриншот", "count": None},
           {"role": "closing", "theme": "c", "count": None}]


def test_without_images_the_image_slide_is_dropped_and_not_inserted():
    assert _roles(OUTLINE) == ["title", "bullet_list", "closing"]
    assert "image_caption" not in _roles(OUTLINE[:2])


def test_with_images_the_old_rule_stands():
    assert _roles(OUTLINE, images_available=True) == ["title", "bullet_list", "image_caption", "closing"]
    assert "image_caption" in _roles(OUTLINE[:2], images_available=True)
