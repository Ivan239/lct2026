"""POST /api/generate/package — a content package as the pipeline's input.

The VK Tech brief asks for «импорт и декомпозиция … контент-пакетов». Until
iter110 a package could be loaded (iter104) but the service took only a brief
string. The endpoint takes the .zip, decomposes it, feeds the labelled brief to
the planner and checks the deck's numbers against the PACKAGE's numbers — the
Appendix 1 question «все цифры со слайда есть в исходных материалах?» — in the
response's warnings.

Called directly, not over HTTP (httpx is not installed), with the planner
stubbed: no LLM, no LibreOffice, no network.
"""

import io
import os
import tempfile
import zipfile

import pytest
from conftest import ROOT

from fastapi import HTTPException, UploadFile

import api.main as api_main
from common.synthesis import SYNTHESIZE
from content_package import load_package

PACKAGE = os.path.join(ROOT, "samples", "content_packages", "feature_smart_search")


def _zip(folder):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for dirpath, _, files in os.walk(folder):
            for name in files:
                full = os.path.join(dirpath, name)
                zf.write(full, os.path.join("pkg", os.path.relpath(full, folder)))
    buf.seek(0)
    return UploadFile(file=buf, filename="pkg.zip")


def _template_id():
    for tid in api_main._template_registry:
        if os.path.exists(os.path.join(api_main.TEMPLATES_DIR, f"{tid}.pptx")):
            return tid
    pytest.skip("no template in the registry")


@pytest.fixture
def stubbed(monkeypatch, tmp_path):
    seen = {}
    fact = next(f for f in load_package(PACKAGE)["facts"] if any(c.isdigit() for c in f))

    def plan(template_id, brief, model=None, slides=None, variant=None):
        seen["brief"] = brief
        seen["slides"] = slides
        seen.setdefault("variants", []).append(variant)
        return [({"type": "title", "title": "Умный поиск", "subtitle": "Питч фичи"}, SYNTHESIZE),
                ({"type": "bullet_list", "title": "Что даёт поиск",
                  "bullets": [fact, "Выручка выросла на 777%"]}, SYNTHESIZE)], []

    monkeypatch.setattr(api_main, "_plan_two_phase", plan)
    monkeypatch.setattr(api_main, "render_pptx_to_pngs", lambda pptx, out: [])
    monkeypatch.setattr(api_main, "_current_balance", lambda: None)
    monkeypatch.setattr(api_main, "_synth_canvas_hints", lambda tid, plan: {})
    monkeypatch.setattr(api_main, "_measured_backgrounds", lambda tid: {})
    monkeypatch.setattr(api_main, "GENERATED_DIR", str(tmp_path / "generated"))
    os.makedirs(tmp_path / "generated")
    workdirs = []
    real_mkdtemp = tempfile.mkdtemp

    def mkdtemp(*args, **kwargs):
        # The same `tempfile` module serves the loader: keep its `dir` (the
        # request's workdir), redirect only the undirected ones into tmp_path.
        kwargs["dir"] = kwargs.get("dir") or str(tmp_path)
        workdirs.append(real_mkdtemp(*args, **kwargs))
        return workdirs[-1]

    monkeypatch.setattr(tempfile, "mkdtemp", mkdtemp)
    seen["workdirs"] = workdirs
    seen["fact"] = fact
    return seen


def test_a_package_reaches_the_planner_and_its_numbers_audit_the_deck(stubbed):
    result = api_main.generate_from_package(template_id=_template_id(), file=_zip(PACKAGE), model=None)

    assert "Факты (используй только их" in stubbed["brief"]
    assert stubbed["fact"] in stubbed["brief"]
    assert stubbed["slides"] == load_package(PACKAGE)["slides"], "the package's deck size is lost"
    assert result["package"]["facts"] > 0 and result["package"]["tables"]

    flagged = [w for w in result["warnings"] if w["kind"] == "numbers_not_in_source"]
    assert [w["slide"] for w in flagged] == [2], result["warnings"]
    assert "777%" in flagged[0]["details"]
    # the package's own fact on the same slide is not flagged
    assert flagged[0]["details"].endswith("777%"), flagged[0]["details"]


def test_the_upload_and_its_extraction_are_removed_after_the_request(stubbed):
    api_main.generate_from_package(template_id=_template_id(), file=_zip(PACKAGE), model=None)
    assert stubbed["workdirs"] and not any(os.path.exists(d) for d in stubbed["workdirs"])


def test_a_package_without_a_brief_is_a_422_that_names_it(stubbed, tmp_path):
    folder = tmp_path / "broken"
    os.makedirs(folder)
    (folder / "facts.md").write_text("- Рост 10%\n", encoding="utf-8")
    with pytest.raises(HTTPException) as err:
        api_main.generate_from_package(template_id=_template_id(), file=_zip(str(folder)), model=None)
    assert err.value.status_code == 422 and "brief.md" in err.value.detail


def test_the_plain_brief_endpoint_audits_numbers_against_the_brief(stubbed):
    """No package: the brief is the whole source. The fact's number is in the
    brief below, «777%» is not."""
    brief = f"Питч умного поиска. {stubbed['fact']}"
    result = api_main.generate_presentation(
        api_main.GenerateRequest(template_id=_template_id(), brief=brief))
    flagged = [w for w in result["warnings"] if w["kind"] == "numbers_not_in_source"]
    assert len(flagged) == 1 and flagged[0]["details"].endswith(": 777%"), result["warnings"]


def test_three_variants_come_back_together_one_per_layout(stubbed):
    """Организаторы: «все 3 презентации генерировались за 5 минут», параллельно.
    Один запрос — три колоды, у каждой свой вариант плана."""
    result = api_main.generate_three_variants(
        api_main.GenerateRequest(template_id=_template_id(), brief="Питч умного поиска"))
    assert [v["variant"] for v in result["variants"]] == list(api_main.VARIANTS)
    assert sorted(stubbed["variants"]) == sorted(api_main.VARIANTS)
    assert len({v["generation_id"] for v in result["variants"]}) == 3
    assert all(v["variant_title"] for v in result["variants"]) and result["seconds"] >= 0


def test_a_package_can_ask_for_all_three(stubbed):
    result = api_main.generate_from_package(template_id=_template_id(), file=_zip(PACKAGE),
                                            model=None, all_variants=True)
    assert len(result["variants"]) == 3 and result["package"]["facts"] > 0
