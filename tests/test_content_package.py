"""Content package import and decomposition (docs/CONTENT_PACKAGE.md).

The VK Tech brief asks for «импорт и декомпозиция … контент-пакетов», and the
organizers ship none, so the format is ours. These tests pin what the rest of
the pipeline will rely on: the parts come back separate (brief, facts, numbers,
tables, images), a zip reads the same as a folder, and a broken package fails
with a message that names the problem instead of an exception three layers
down.
"""

import json
import os
import zipfile

import pytest

from conftest import ROOT

from content_package import (ContentPackageError, extract_numbers, load_package,
                             to_brief_text)

PNG_1PX = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010806000000"
    "1f15c4890000000d49444154789c6360000002000154a24f6e0000000049454e44ae426082")


def _package(root, *, manifest=None, brief="# Платформа «Поток»\nНужно убедить директора.",
             facts=None, tables=None, images=()):
    os.makedirs(root, exist_ok=True)
    if manifest is not None:
        with open(os.path.join(root, "package.json"), "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False)
    if brief is not None:
        with open(os.path.join(root, "brief.md"), "w", encoding="utf-8") as f:
            f.write(brief)
    if facts is not None:
        with open(os.path.join(root, "facts.md"), "w", encoding="utf-8") as f:
            f.write(facts)
    for name, text in (tables or {}).items():
        os.makedirs(os.path.join(root, "data"), exist_ok=True)
        with open(os.path.join(root, "data", name), "w", encoding="utf-8") as f:
            f.write(text)
    for name in images:
        os.makedirs(os.path.join(root, "images"), exist_ok=True)
        with open(os.path.join(root, "images", name), "wb") as f:
            f.write(PNG_1PX)
    return root


def test_a_folder_is_decomposed_into_separate_parts(tmp_path):
    root = _package(
        str(tmp_path / "pkg"),
        manifest={"purpose": "product", "slides": 12, "audience": "коммерческие директора",
                  "images": {"dashboard.png": "Экран единого дашборда"}},
        facts="- Рост конверсии +25 %\n- Экономия 30 часов в месяц\n\n# комментарий\n",
        tables={"sales.csv": "Квартал;Выручка, млн ₽\nQ1;120\nQ2;148\n"},
        images=["dashboard.png", "team_photo.jpg"])
    pkg = load_package(root)
    assert pkg["title"] == "Платформа «Поток»"
    assert pkg["purpose_label"] == "продукт"
    assert pkg["slides"] == 12
    assert pkg["facts"] == ["Рост конверсии +25 %", "Экономия 30 часов в месяц"]
    table = pkg["tables"][0]
    assert table["name"] == "sales"
    assert table["columns"] == ["Квартал", "Выручка, млн ₽"]
    assert table["rows"] == [["Q1", "120"], ["Q2", "148"]]
    assert table["numeric_columns"] == ["Выручка, млн ₽"]
    assert [i["caption"] for i in pkg["images"]] == ["Экран единого дашборда", "team photo"]
    assert {"+25%", "30часов", "120", "148"} <= set(pkg["numbers"])


def test_a_zip_reads_the_same_as_the_folder_it_was_made_from(tmp_path):
    root = _package(str(tmp_path / "pkg"), facts="- Точность прогноза 95%\n",
                    tables={"t.csv": "a,b\n1,2\n"}, images=["x.png"])
    archive = str(tmp_path / "pkg.zip")
    with zipfile.ZipFile(archive, "w") as zf:
        for dirpath, _, files in os.walk(root):
            for name in files:
                full = os.path.join(dirpath, name)
                zf.write(full, os.path.join("pkg", os.path.relpath(full, root)))
    from_dir, from_zip = load_package(root), load_package(archive)
    for key in ("title", "brief", "facts", "numbers", "tables"):
        assert from_zip[key] == from_dir[key], key
    assert os.path.isfile(from_zip["images"][0]["path"])


def test_a_package_without_a_brief_is_refused_with_a_reason(tmp_path):
    root = _package(str(tmp_path / "pkg"), brief=None, facts="- факт\n")
    with pytest.raises(ContentPackageError, match="brief.md"):
        load_package(root)


def test_an_unknown_purpose_is_refused_with_the_allowed_list(tmp_path):
    root = _package(str(tmp_path / "pkg"), manifest={"purpose": "маркетинг"})
    with pytest.raises(ContentPackageError, match="feature"):
        load_package(root)


def test_numbers_keep_their_units_and_normalise_spacing():
    assert extract_numbers("рост на +25 % и экономия 30 часов") == ["+25%", "30часов"]
    assert extract_numbers("бюджет 20 227 000 ₽, доля 3,5 %") == ["20227000₽", "3.5%"]
    assert extract_numbers("версия v2 и Q4") == []


def test_the_flattened_brief_keeps_the_parts_labelled(tmp_path):
    root = _package(str(tmp_path / "pkg"), manifest={"purpose": "initiative", "slides": 10},
                    facts="- Срок пилота 2 недели\n",
                    tables={"kpi.csv": "Метрика,Значение\nNPS,48\n"}, images=["a.png"])
    text = to_brief_text(load_package(root))
    assert "Назначение: инициатива" in text
    assert "Объём: 10 слайдов" in text
    assert "- Срок пилота 2 недели" in text
    assert "Таблица «kpi»: Метрика, Значение" in text
    assert "  NPS | 48" in text
    assert "Нужно убедить директора." in text


def test_a_semicolon_csv_with_decimal_commas_keeps_its_columns(tmp_path):
    """The usual Russian export: «;» between fields, a decimal comma inside
    them. csv.Sniffer chose the comma and split «План, млн ₽» and «6,2»."""
    root = _package(str(tmp_path / "pkg"),
                    tables={"budget.csv": "Месяц;План, млн ₽;Факт, млн ₽\nИюль;6,2;5,8\nАвгуст;6,0;5,6\n"})
    table = load_package(root)["tables"][0]
    assert table["columns"] == ["Месяц", "План, млн ₽", "Факт, млн ₽"]
    assert table["rows"] == [["Июль", "6,2", "5,8"], ["Август", "6,0", "5,6"]]
    assert table["numeric_columns"] == ["План, млн ₽", "Факт, млн ₽"]


def test_the_sample_packages_load_with_facts_numbers_and_a_table():
    """The three packages in samples/ are the loop's content until real ones
    exist: one per purpose from the brief, each with facts and a table."""
    base = os.path.join(ROOT, "samples", "content_packages")
    purposes = set()
    for name in sorted(os.listdir(base)):
        pkg = load_package(os.path.join(base, name))
        purposes.add(pkg["purpose"])
        assert pkg["facts"] and pkg["numbers"], name
        assert pkg["tables"] and all(len(t["columns"]) > 1 for t in pkg["tables"]), name
        assert any(t["numeric_columns"] for t in pkg["tables"]), name
    assert {"feature", "project", "initiative"} <= purposes
