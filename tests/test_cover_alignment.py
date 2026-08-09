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
    theme["cover_align"], theme["cover_left"] = None, None
    idx = synthesize_title(prs, theme, infer_content_bounds(extract_template(SURVEY_31)), COVER)
    title = next(s for s in prs.slides[idx].shapes
                 if s.has_text_frame and s.text_frame.text.strip() == COVER["title"])
    assert int(title.left) == int(prs.slide_width * 0.1)
