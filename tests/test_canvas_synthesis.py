"""Canvas synthesis (plan 9.3): a role the template lacks is synthesized on a
CLONE of a template slide — its background art and chrome survive — instead of
a blank white page. Verified structurally (renders are eyeballed in sweeps)."""

import os

from conftest import SURVEY_31, TEMPLATES_DIR, requires

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from common.synthesis import SYNTHESIZE
from generator.generator import generate

TJ = os.path.join(TEMPLATES_DIR, "custom_f496182bb15f42bb.pptx")

PLAN = [
    ({"type": "title", "title": "Т", "subtitle": "П"}, 4),
    ({"type": "closing", "title": "Спасибо!", "subtitle": "Пока"}, SYNTHESIZE),
]


@requires(TJ)
def test_canvas_keeps_background_art(tmp_path):
    out = str(tmp_path / "canvas.pptx")
    generate(TJ, PLAN, out, synth_canvas={1: 10})
    prs = Presentation(out)
    assert len(prs.slides) == 2
    synth = prs.slides[1]
    # The T-Ж folder background is a full-bleed PICTURE — it must survive.
    assert any(s.shape_type == MSO_SHAPE_TYPE.PICTURE for s in synth.shapes)
    texts = " ".join(
        s.text_frame.text for s in synth.shapes if s.has_text_frame
    )
    assert "Спасибо!" in texts
    # The canvas's own content text must NOT leak through.
    assert "что вы есть" not in texts


@requires(TJ)
def test_without_canvas_hint_still_works(tmp_path):
    out = str(tmp_path / "no_canvas.pptx")
    generate(TJ, PLAN, out)  # from-scratch path — package gate must still pass
    assert Presentation(out).slides


@requires(TJ)
def test_image_frame_survives_a_narrow_canvas_and_clears_its_chrome():
    """A canvas is REUSED for a different role, so its own text area is not
    automatically a usable content area. T-Zh study slide 10 is one low line and
    hands back a 1.03in band against the template's 4.04in: the title ate it,
    both candidate bands for the image frame came out NEGATIVE, and the slide
    shipped as a heading alone on a full-page empty card. Every T-Zh template
    has three such canvases, and they are exactly the ones offered for synthesis.

    Falling back to the template-wide band then has to respect the furniture the
    clone still carries — first render after the fallback showed the dashed
    frame crossing the footer rule and boxing in «КОММЕНТАРИЙ»."""
    from pptx.util import Emu

    from generator.deck_style import apply_observed_style, observe_deck_style
    from generator.layout_bounds import infer_content_bounds
    from generator.slide_kit import is_chrome_shape
    from generator.synthesizer import IMAGE_PLACEHOLDER_NAME, synthesize_image_caption
    from template_parser.parser import extract_template, extract_theme

    prs = Presentation(TJ)
    theme = apply_observed_style(extract_theme(TJ), observe_deck_style(prs))
    bounds = infer_content_bounds(extract_template(TJ))

    idx = synthesize_image_caption(
        prs, theme, bounds,
        {"title": "Платформа в работе", "image": "экран дашборда"},
        canvas_idx=10,
    )
    slide = prs.slides[idx]
    frames = [s for s in slide.shapes if s.name == IMAGE_PLACEHOLDER_NAME]
    assert frames, "the image slide shipped as a bare heading"

    frame = frames[0]
    footers = [
        s for s in slide.shapes
        if s is not frame and s.top is not None and s.height is not None
        and is_chrome_shape(s, prs.slide_height)
        and s.top > prs.slide_height // 2
    ]
    assert footers, "fixture changed: this canvas is supposed to carry footer chrome"
    lowest_allowed = min(int(s.top) for s in footers)
    assert int(frame.top + frame.height) <= lowest_allowed, (
        f"frame runs into the canvas footer: ends at "
        f"{Emu(frame.top + frame.height).inches:.2f}in, footer starts at "
        f"{Emu(lowest_allowed).inches:.2f}in")


@requires(TJ)
def test_canvas_placeholders_are_blanked_but_managed_chrome_is_kept():
    """A clone keeps the whole footer band, and nothing on this path blanked it:
    «КОММЕНТАРИЙ» shipped verbatim under EVERY synthesized slide of this deck.
    The native fill path has always cleared exactly this class of text.

    The split matters as much as the clearing — the page number and the running
    topic slot are owned by later passes (iter31) and must survive, or this fix
    just re-breaks what that one repaired."""
    from generator.deck_style import apply_observed_style, observe_deck_style
    from generator.layout_bounds import infer_content_bounds
    from generator.slide_kit import is_chrome_shape
    from generator.synthesizer import synthesize_image_caption
    from template_parser.parser import extract_template, extract_theme

    prs = Presentation(TJ)
    theme = apply_observed_style(extract_theme(TJ), observe_deck_style(prs))
    bounds = infer_content_bounds(extract_template(TJ))

    # Slide 8 carries all three: the topic slot, a page number and the
    # ownerless «КОММЕНТАРИЙ» — the split this test is about.
    canvas = prs.slides[8]
    before = [
        s.text_frame.text.strip() for s in canvas.shapes
        if s.has_text_frame and is_chrome_shape(s, prs.slide_height)
        and s.text_frame.text.strip()
    ]
    assert any("КОММЕНТАРИЙ" in t for t in before), (
        "fixture changed: this canvas is supposed to carry the placeholder")
    assert any(t.isdigit() for t in before), (
        "fixture changed: this canvas is supposed to carry a page number")

    idx = synthesize_image_caption(
        prs, theme, bounds, {"title": "Платформа", "image": "дашборд"}, canvas_idx=8)
    after = [
        s.text_frame.text.strip() for s in prs.slides[idx].shapes
        if s.has_text_frame and is_chrome_shape(s, prs.slide_height)
        and s.text_frame.text.strip()
    ]
    assert not any("КОММЕНТАРИЙ" in t for t in after), f"placeholder shipped: {after}"
    assert any(t.isdigit() for t in after), f"the page number was wiped too: {after}"


@requires(TJ)
def test_image_frame_does_not_reserve_room_for_a_wrap_that_cannot_happen():
    """The clearance under the title carries a safety margin because the height
    estimate can under-count lines (the renderer wraps 2-4% earlier than the
    metrics predict). That margin has to be EARNED: a title that provably fits
    on one line was still given the two-line floor, leaving 0.47in of dead band
    between the heading and the frame — a visible hole on the render.

    Both directions are checked: the short title gets exactly the gap, the long
    one keeps its reserve and the frame never climbs into it."""
    from pptx.util import Emu, Inches

    from fonts.metrics import FontResolver
    from generator.deck_style import apply_observed_style, observe_deck_style
    from generator.layout_bounds import infer_content_bounds
    from generator.synthesizer import IMAGE_PLACEHOLDER_NAME, synthesize_image_caption
    from template_parser.parser import extract_template, extract_theme

    bounds = infer_content_bounds(extract_template(TJ))
    gaps = {}
    for label, title in (
        ("short", "Платформа в работе"),
        ("long", "Как мы перестроили процесс согласования отчётности в компании"),
    ):
        prs = Presentation(TJ)
        theme = apply_observed_style(extract_theme(TJ), observe_deck_style(prs))
        idx = synthesize_image_caption(
            prs, theme, bounds, {"title": title, "image": "дашборд"},
            resolver=FontResolver(TJ, extract_theme(TJ)), canvas_idx=8)
        shapes = prs.slides[idx].shapes
        frame = next(s for s in shapes if s.name == IMAGE_PLACEHOLDER_NAME)
        title_box = next(s for s in shapes if s.name.startswith("TextBox"))
        gaps[label] = int(frame.top) - int(title_box.top + title_box.height)

    assert 0 < gaps["short"] <= int(Inches(0.4)), (
        f"dead band under a title that cannot wrap: {Emu(gaps['short']).inches:.2f}in")
    assert gaps["long"] > 0, "the frame climbed into a title that did wrap"


@requires(SURVEY_31)
def test_synthesized_block_steps_around_the_canvas_decor():
    """The survey-31 canvas keeps a small hand-drawn heart at 4.53-5.46in, and
    the synthesized KPI block landed across it — the render showed the artwork
    sitting on top of «клиентов в месяц».

    Two causes, and neither fix works alone. The canvas's own band is 2.82in on
    a 7.5in slide, while a heading plus three KPI pairs need about 2.9in, so the
    content overflowed the band's bottom — which is exactly where the heart
    lives. Widening the band alone just re-centres the block onto the heart;
    stepping around obstacles alone does nothing, because the 2.82in band offers
    no clear position. Measured in that order, one attempt at a time."""
    from pptx.util import Emu

    from generator.deck_style import apply_observed_style, observe_deck_style
    from generator.layout_bounds import infer_content_bounds
    from generator.synthesizer import synthesize_stats_kpi
    from template_parser.parser import extract_template, extract_theme

    prs = Presentation(SURVEY_31)
    theme = apply_observed_style(extract_theme(SURVEY_31), observe_deck_style(prs))
    bounds = infer_content_bounds(extract_template(SURVEY_31))

    idx = synthesize_stats_kpi(
        prs, theme, bounds,
        {"title": "Результаты",
         "stats": [["-40%", "времени на отчёт"], ["+18", "клиентов в месяц"],
                   ["4 дня", "на внедрение"]]},
        canvas_idx=29)
    slide = prs.slides[idx]

    decor = [s for s in slide.shapes if "PICTURE" in str(s.shape_type)]
    assert decor, "fixture changed: this canvas is supposed to carry decor"
    texts = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    assert len(texts) >= 4, f"expected a title and KPI pairs, got {len(texts)}"

    for art in decor:
        art_top, art_bottom = int(art.top), int(art.top + art.height)
        for box in texts:
            top, bottom = int(box.top), int(box.top + box.height)
            assert not (top < art_bottom and bottom > art_top), (
                f"{box.text_frame.text.strip()[:24]!r} at "
                f"{Emu(top).inches:.2f}-{Emu(bottom).inches:.2f}in runs across decor at "
                f"{Emu(art_top).inches:.2f}-{Emu(art_bottom).inches:.2f}in")


def test_canvas_picker_skips_the_templates_readme_page():
    """"Fewest text boxes" picks the blandest slide, and the blandest slide is
    sometimes the template's own instructions page. The T-Zh universal deck ends
    with a «Технический слайд» telling whoever uses the template which fonts to
    download; it has 3 boxes and no furniture, the branded candidates have 5 and
    9, so the <=3 cut left it as the only option and ALL THREE synthesized
    slides were cloned from it — bare white pages on a template whose identity
    is black with yellow artwork.

    Its background is (248,248,248) on exactly one slide of the deck, so
    singleton backgrounds are dropped first. The rotation must survive that:
    filtering to the deck's MAJORITY background instead was tried and reverted,
    because it collapsed T-Zh mono from alternating blue/white to white only."""
    import glob
    import json
    import os

    from common.synthesis import SYNTHESIZE
    from design_system.style_profile import build_measured_profile
    from evaluation.loop import _synth_canvas_hints

    def hints_for(tid, model, positions=4):
        cache = f"output/loop/{model}/{tid}/archetypes.json"
        pngs = sorted(glob.glob(f"output/loop/{model}/{tid}/rendered/*.png"))
        if not os.path.exists(cache) or not pngs:
            return None
        archetypes = {int(k): v for k, v in json.load(open(cache)).items()}
        profile = build_measured_profile(pngs, archetypes)
        plan = [({"type": "bullet_list"}, SYNTHESIZE)] * positions
        return _synth_canvas_hints(f"output/templates/{tid}.pptx", plan, profile)

    universal = hints_for("custom_47dfd8952eb47583", "GigaChat-3-Ultra")
    if universal:
        assert 11 not in universal.values(), (
            f"the readme page is being cloned as a canvas: {universal}")

    mono = hints_for("custom_838830368dac3116", "GigaChat-2")
    if mono:
        assert len(set(mono.values())) > 1, (
            f"colour rotation collapsed to a single canvas: {mono}")
