"""A synthesized cover/closing is placed like the template's own cover.

Both were centred unconditionally in a box inset 10% from each edge. Measured on
the title slide of every real template, EVERY one is left aligned:

    universal 0.39in   mono 0.19in   study 0.39in   survey-31 1.10in   survey-69 1.10in

so a centred block starting at 1.33in matched none of them, and the closing read
as coming from a different template than the slides before it — visible on the
render of the mono deck, whose own cover sets its title hard against the left
edge while ours sat centred.
"""

from conftest import SURVEY_31, TJ_MONO, TJ_UNIVERSAL, requires

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu

from fonts.metrics import FontResolver
from generator.deck_style import apply_observed_style, observe_deck_style
from generator.generator import _pick_title_shape, _template_cover_pt
from generator.layout_bounds import infer_content_bounds
from generator.slide_kit import is_chrome_shape
from generator.synthesizer import synthesize_title
from template_parser.parser import extract_template, extract_theme

COVER = {"title": "Запустим пилот в вашем подразделении",
         "subtitle": "Две недели на подключение и обучение команды"}


def _synthesized_cover(template):
    prs = Presentation(template)
    theme = apply_observed_style(extract_theme(template), observe_deck_style(prs))
    # As generate() assembles it — the measured values live in the theme. Read
    # here straight off the template rather than through the new helper: an
    # import of something HEAD lacks aborts the module, and «the helper is
    # missing» says far less than «the cover sits at 1.33in instead of 0.19in».
    theme["cover_pt"] = _template_cover_pt(prs)
    own = _pick_title_shape(list(prs.slides)[0], set())
    theme["cover_align"] = own.text_frame.paragraphs[0].alignment if own is not None else None
    theme["cover_left"] = int(own.left) if own is not None and own.left is not None else None
    theme["cover_top"] = int(own.top) if own is not None and own.top is not None else None
    below = [x for x in prs.slides[0].shapes
             if x.has_text_frame and x.text_frame.text.strip() and own is not None
             and x.shape_id != own.shape_id and x.top is not None and x.top >= own.top
             and not is_chrome_shape(x, prs.slide_height)]
    theme["cover_sub_top"] = int(max(x.top for x in below)) if below else None
    idx = synthesize_title(prs, theme, infer_content_bounds(extract_template(template)), COVER,
                           resolver=FontResolver(template, extract_theme(template)))
    slide = prs.slides[idx]
    boxes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    title = next(s for s in boxes if s.text_frame.text.strip() == COVER["title"])
    subtitle = next(s for s in boxes if s.text_frame.text.strip() == COVER["subtitle"])
    return title, subtitle


@requires(TJ_MONO)
@requires(SURVEY_31)
def test_cover_takes_the_left_margin_and_alignment_from_the_template():
    """Two templates with very different margins — 0.19in and 1.10in — so a
    hard-coded value of any kind fails one of them."""
    for template in (TJ_MONO, SURVEY_31):
        source = Presentation(template)
        own_title = _pick_title_shape(list(source.slides)[0], set())
        title, subtitle = _synthesized_cover(template)

        assert int(title.left) == int(own_title.left), (
            f"{template}: cover at {Emu(title.left).inches:.2f}in, the template's "
            f"own title at {Emu(own_title.left).inches:.2f}in")
        assert title.text_frame.paragraphs[0].alignment != PP_ALIGN.CENTER
        assert int(subtitle.left) == int(title.left), "subtitle must share the title's edge"


@requires(TJ_MONO)
@requires(SURVEY_31)
def test_cover_takes_its_top_from_the_template_too():
    """A flat 38% of the slide height matched exactly one template of five.
    Measured tops: 26 / 37 / 4 / 22 / 26% — mono opens hard against the top edge
    at 0.24in, survey-31 at 1.94in, so no constant satisfies both."""
    for template in (TJ_MONO, SURVEY_31):
        source = Presentation(template)
        own_title = _pick_title_shape(list(source.slides)[0], set())
        title, _ = _synthesized_cover(template)
        assert int(title.top) == int(own_title.top), (
            f"{template}: cover top {Emu(title.top).inches:.2f}in, the template's own "
            f"{Emu(own_title.top).inches:.2f}in")


@requires(TJ_MONO)
def test_the_subtitle_sits_where_the_template_puts_its_own():
    """All three T-Zh covers anchor the subtitle near the BOTTOM edge — 0.19 to
    0.35in above it — while the gap they leave under the title is 1.21 / 2.55 /
    2.28in, i.e. no gap at all, just whatever is left. Tucking it under the
    title left mono's closing with two empty thirds below."""
    source = Presentation(TJ_MONO)
    own_title = _pick_title_shape(list(source.slides)[0], set())
    own_sub = max((x for x in list(source.slides)[0].shapes
                   if x.has_text_frame and x.text_frame.text.strip()
                   and x.shape_id != own_title.shape_id and x.top is not None
                   and not is_chrome_shape(x, source.slide_height)),
                  key=lambda x: x.top)
    _, subtitle = _synthesized_cover(TJ_MONO)
    assert int(subtitle.top) == int(own_sub.top), (
        f"subtitle at {Emu(subtitle.top).inches:.2f}in, the template's own at "
        f"{Emu(own_sub.top).inches:.2f}in")


@requires(TJ_MONO)
def test_a_long_title_still_pushes_the_subtitle_below_itself():
    """The measured position may only move the subtitle DOWN: iter59's collision
    (a subtitle printed through an overflowing title) must stay impossible."""
    long_cover = {"title": " ".join(["Запустим пилот в вашем подразделении"] * 3),
                  "subtitle": COVER["subtitle"]}
    prs = Presentation(TJ_MONO)
    theme = apply_observed_style(extract_theme(TJ_MONO), observe_deck_style(prs))
    own = _pick_title_shape(list(prs.slides)[0], set())
    theme["cover_pt"] = _template_cover_pt(prs)
    theme["cover_align"] = own.text_frame.paragraphs[0].alignment
    theme["cover_left"], theme["cover_top"] = int(own.left), int(own.top)
    theme["cover_sub_top"] = int(own.top)  # deliberately ABOVE where the title ends
    idx = synthesize_title(prs, theme, infer_content_bounds(extract_template(TJ_MONO)),
                           long_cover, resolver=FontResolver(TJ_MONO, extract_theme(TJ_MONO)))
    boxes = [s for s in prs.slides[idx].shapes if s.has_text_frame and s.text_frame.text.strip()]
    title = next(s for s in boxes if s.text_frame.text.strip() == long_cover["title"])
    subtitle = next(s for s in boxes if s.text_frame.text.strip() == long_cover["subtitle"])
    assert int(subtitle.top) >= int(title.top + title.height)


@requires(TJ_UNIVERSAL)
def test_the_cover_still_fits_the_slide_from_its_new_left_edge():
    """Moving the block right must not push it off the other side: the width is
    taken from the remaining room, not kept at 80% of the slide."""
    title, _ = _synthesized_cover(TJ_UNIVERSAL)
    prs = Presentation(TJ_UNIVERSAL)
    assert int(title.left + title.width) <= int(prs.slide_width), (
        f"cover runs to {Emu(title.left + title.width).inches:.2f}in on a "
        f"{Emu(prs.slide_width).inches:.2f}in slide")


@requires(SURVEY_31)
def test_a_template_without_a_readable_title_keeps_the_old_placement():
    """The fallback must survive: no measurement, no change in behaviour."""
    prs = Presentation(SURVEY_31)
    theme = apply_observed_style(extract_theme(SURVEY_31), observe_deck_style(prs))
    theme["cover_align"] = theme["cover_left"] = None
    theme["cover_top"] = theme["cover_sub_top"] = None
    idx = synthesize_title(prs, theme, infer_content_bounds(extract_template(SURVEY_31)), COVER)
    title = next(s for s in prs.slides[idx].shapes
                 if s.has_text_frame and s.text_frame.text.strip() == COVER["title"])
    assert int(title.left) == int(prs.slide_width * 0.1)
    assert int(title.top) == int(prs.slide_height * 0.38)
