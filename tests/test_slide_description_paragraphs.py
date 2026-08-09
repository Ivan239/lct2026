"""The slide description the classifier sees must show paragraph structure.

survey-31's parse labels ONE of its 31 slides and shrugs at the other 30, and
the reason is in the description, not in the model: its content slides are
"title + one text box", and that box holds 4, 5, 7 or 8 separate paragraphs —
unmistakable bullet lists — which the description joined into a single
80-character run-on. The model was asked to tell a list from prose with the
list-ness removed.

Whether the classification actually improves can only be seen with the network;
what is verified here is that the decisive fact now reaches the prompt.
"""

from conftest import SURVEY_31, TJ_UNIVERSAL, requires

from design_system.extractor import _describe_slide
from template_parser.parser import extract_template

LIST_SLIDE = 24  # «Top 7» — seven marked lines in one box


@requires(SURVEY_31)
def test_a_multi_paragraph_box_is_described_as_a_list():
    template = extract_template(SURVEY_31)
    body = [s for s in template["slides"][LIST_SLIDE]["shapes"] if s["text"]]
    counts = [len([p for p in s["text"] if any(r["text"].strip() for r in p)]) for s in body]
    assert max(counts) == 7, f"fixture changed: paragraph counts {counts}"

    described = _describe_slide(template["slides"][LIST_SLIDE], template["slide_size_in"])
    assert "7 paragraphs" in described, described
    # …and more than the first one, so the model can see they are peers rather
    # than one sentence that happens to be long.
    assert described.count(" | ") >= 2, described


@requires(SURVEY_31)
def test_a_single_paragraph_box_is_unchanged():
    template = extract_template(SURVEY_31)
    described = _describe_slide(template["slides"][LIST_SLIDE], template["slide_size_in"])
    assert '"Top 7"' in described, "a one-paragraph title must keep its plain form"
    assert "1 paragraphs" not in described


@requires(TJ_UNIVERSAL)
def test_the_change_adds_no_noise_to_a_template_without_lists():
    """The T-Zh grid slide has thirteen one-line boxes; none of them may gain a
    paragraph note, or the description grows without saying anything."""
    template = extract_template(TJ_UNIVERSAL)
    described = _describe_slide(template["slides"][2], template["slide_size_in"])
    assert described.count("paragraphs") == 1, described  # only the hand-broken title
