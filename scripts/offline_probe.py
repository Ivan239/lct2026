"""Generate + render + score a deck on EVERY real template, without the network.

Why this exists: the loop's own runner needs GigaChat, and the improvement work
does not stop when the network does. This drives the same pipeline blocks —
build_spec -> plan_from_outline -> _synth_canvas_hints -> generate — with content
written here instead of by a model, so every layout decision is exercised and the
renders can be looked at.

Two things it deliberately does that the earlier throwaway probes got wrong:

1. It honours the plan's `count`. A probe that passes a fixed three-item block
   ignores the capacity the matcher just computed, so a capacity fix looks like a
   no-op — that happened for real (iter42: a one-big-number stats slide kept
   showing three lines after the fix, because the probe never asked for one).

2. It defaults to LONG content. Short, convenient bullets stop finding defects:
   the same five decks scored 99.4-100 and looked clean on short content, and the
   moment the text became realistic the checks produced three flags and the
   renders showed 8pt microtext (iter51) and decor sitting on a KPI figure
   (iter52). `--content short` is kept for comparing the two.

3. It can vary the ITEM COUNT. The plan's capacity match decides which template
   slide a role lands on, so a fixed brief only ever exercises one branch: with
   three bullets the universal template hands the list to its prose slide
   (capacity 3) and its designed 6-cell numbered grid — a whole layout — had
   never been rendered by this probe. `--items 5` reaches it.

Usage:
    .venv/bin/python3 scripts/offline_probe.py [--content long|short] [--only NAME]
                                               [--items N]
"""

import argparse
import datetime
import glob
import json
import os
import shutil
import statistics
import subprocess
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from pptx import Presentation  # noqa: E402
from pptx.util import Emu  # noqa: E402

from common.synthesis import SYNTHESIZE  # noqa: E402
from content_parser.two_phase import _enforce_text_budgets, bullet_char_budget  # noqa: E402
from design_system.style_profile import build_measured_profile  # noqa: E402
from evaluation.evaluate import evaluate_deck  # noqa: E402
from evaluation.loop import _synth_canvas_hints  # noqa: E402
from fonts.metrics import FontResolver  # noqa: E402
from generator.text_fit import (SINGLE_LINE_SAFETY, estimate_block_height_in,  # noqa: E402
                                estimate_wrapped_lines, horizontal_margins_in,
                                vertical_insets_emu)
from template_parser.parser import extract_theme  # noqa: E402
from generator.generator import _pick_title_shape, generate, is_chrome_shape  # noqa: E402
from matcher.matcher import plan_from_outline  # noqa: E402
from rendering.render import render_pptx_to_pngs  # noqa: E402
from template_spec.builder import build_spec  # noqa: E402

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TEMPLATES = os.path.join(BASE, "output", "templates")
LOOP_ROOT = os.path.join(BASE, "output", "loop")
OUT_ROOT = os.path.join(BASE, "output", "offline_probe")

REAL_TEMPLATES = [
    ("universal", "custom_47dfd8952eb47583"),
    ("mono", "custom_838830368dac3116"),
    ("study", "custom_f496182bb15f42bb"),
    ("survey-31", "custom_30e96c06e2d47ec3"),
    ("survey-69", "custom_78dc579e05d11399"),
]

OUTLINE = [
    {"role": "title", "theme": "платформа", "count": None},
    {"role": "bullet_list", "theme": "что мешало", "count": 3},
    {"role": "stats_kpi", "theme": "результаты", "count": 3},
    {"role": "two_column_comparison", "theme": "было и стало", "count": 2},
    {"role": "image_caption", "theme": "платформа в работе", "count": None},
    {"role": "closing", "theme": "пилот", "count": None},
]

LONG = {
    "title": {"type": "title", "title": "Платформа управленческой отчётности «Поток»",
              "subtitle": "Итоги пилотного внедрения в трёх подразделениях за квартал"},
    "bullet_list": {"type": "bullet_list", "title": "Что мешало собирать отчётность вовремя",
                    "bullets": ["Ручной сбор показателей из семи независимых систем",
                                "Разные форматы выгрузок у каждого подразделения",
                                "Согласование занимало до двух недель"]},
    "stats_kpi": {"type": "stats_kpi", "title": "Результаты пилотного внедрения",
                  "stats": [["-40%", "времени на подготовку регулярной отчётности"],
                            ["+18", "подключённых клиентов ежемесячно"],
                            ["4 дня", "на полное внедрение в подразделении"]]},
    "two_column_comparison": {"type": "two_column_comparison", "title": "Как изменился процесс",
                              "left_heading": "Было",
                              "left_points": ["Ручной сбор показателей", "Разрозненные файлы в почте"],
                              "right_heading": "Стало",
                              "right_points": ["Автоматический сбор по расписанию", "Единая витрина данных"]},
    "image_caption": {"type": "image_caption", "title": "Как выглядит витрина данных",
                      "image": "экран дашборда с графиком отгрузок по неделям"},
    "closing": {"type": "closing", "title": "Запустим пилот в вашем подразделении",
                "subtitle": "Две недели на подключение и обучение команды"},
}

# A third brief, with the TITLES long. The last two defects — a cover title
# running through the balloon artwork and a running header cut to «ЕДИНАЯ
# ПЛАТФОРМА УПРАВЛЕНЧЕСКОЙ» — were both found on decks assembled by hand,
# because LONG and SHORT differ in the length of their BODY text and keep the
# titles short. Anything that depends on a title's length was invisible here.
WORDY = {
    "title": {"type": "title",
              "title": "Единая платформа управленческой отчётности и аналитики для розничной сети",
              "subtitle": "Итоги пилотного внедрения в трёх подразделениях за четвёртый квартал"},
    "bullet_list": {"type": "bullet_list",
                    "title": "Что мешало собирать управленческую отчётность вовремя и без ручной сверки",
                    "bullets": ["Ручной сбор показателей из семи независимых систем",
                                "Разные форматы выгрузок у каждого подразделения",
                                "Согласование занимало до двух недель"]},
    "stats_kpi": {"type": "stats_kpi",
                  "title": "Результаты пилотного внедрения в трёх подразделениях за квартал",
                  "stats": [["-40%", "времени на подготовку регулярной отчётности"],
                            ["+18", "подключённых клиентов ежемесячно"],
                            ["4 дня", "на полное внедрение в подразделении"]]},
    "two_column_comparison": {"type": "two_column_comparison",
                              "title": "Как изменился процесс подготовки управленческой отчётности",
                              "left_heading": "Было",
                              "left_points": ["Ручной сбор показателей", "Разрозненные файлы в почте"],
                              "right_heading": "Стало",
                              "right_points": ["Автоматический сбор по расписанию", "Единая витрина данных"]},
    "image_caption": {"type": "image_caption",
                      "title": "Как выглядит витрина данных в интерфейсе руководителя подразделения",
                      "image": "экран дашборда с графиком отгрузок по неделям"},
    "closing": {"type": "closing",
                "title": "Запустим пилот в вашем подразделении уже в этом квартале",
                "subtitle": "Две недели на подключение и обучение команды"},
}


SHORT = {
    "title": {"type": "title", "title": "Платформа Поток", "subtitle": "Итоги квартала"},
    "bullet_list": {"type": "bullet_list", "title": "Что мешало",
                    "bullets": ["Ручные отчёты", "Разные форматы", "Долгие согласования"]},
    "stats_kpi": {"type": "stats_kpi", "title": "Результаты",
                  "stats": [["-40%", "на отчёт"], ["+18", "клиентов"], ["4 дня", "внедрение"]]},
    "two_column_comparison": {"type": "two_column_comparison", "title": "Было и стало",
                              "left_heading": "Было", "left_points": ["Ручной сбор", "Файлы"],
                              "right_heading": "Стало", "right_points": ["Автосбор", "Витрина"]},
    "image_caption": {"type": "image_caption", "title": "Платформа в работе",
                      "image": "экран дашборда"},
    "closing": {"type": "closing", "title": "Запустим пилот", "subtitle": "Две недели"},
}


def _classified(archetypes):
    return sum(1 for role in archetypes.values() if role != "other")


def _cached_parse(template_id):
    """The archetype map + renders any earlier loop run left behind. Classifying
    needs the network; reusing a cached parse is what makes this offline."""
    best = None
    for model in sorted(os.listdir(LOOP_ROOT)) if os.path.isdir(LOOP_ROOT) else []:
        path = os.path.join(LOOP_ROOT, model, template_id, "archetypes.json")
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8") as f:
            archetypes = {int(k): v for k, v in json.load(f).items()}
        # Skip a parse that is nothing but "other": that is a classification
        # that failed (loop._is_failed_parse), and probing against it measures
        # a deck with no native slides for reasons that have nothing to do with
        # the code under test. Measured: 9 of 20 cached parses were like that.
        if set(archetypes.values()) <= {"other"}:
            continue
        pngs = sorted(glob.glob(os.path.join(LOOP_ROOT, model, template_id, "rendered", "*.png")))
        best = (archetypes, pngs) if best is None or _classified(archetypes) > _classified(best[0]) else best
    return best if best is not None else (None, None)


def _stamp(deck_path, content, items):
    """Write what produced this deck into the file's own properties.

    Twice now a measurement has read a deck from an older run and drawn a
    conclusion from it: once because a skipped template kept its output (fixed
    in iter72 by clearing the directory), and once because an --out directory
    from an earlier session was still on disk with a preset that no longer
    matched. Clearing helps only the directories a run actually touches; a stamp
    travels with the file, so a reader can always ask what it is looking at.
    """
    try:
        sha = subprocess.run(["git", "-C", BASE, "rev-parse", "--short", "HEAD"],
                             capture_output=True, text=True, timeout=10).stdout.strip()
    except Exception:
        sha = "?"
    prs = Presentation(deck_path)
    prs.core_properties.comments = (
        f"{STAMP_MARKER} content={content} items={items or 'plan'} "
        f"at {datetime.datetime.now().isoformat(timespec='seconds')} commit={sha}"
    )
    prs.save(deck_path)


STAMP_MARKER = "offline_probe"


def describe(deck_path):
    """The stamp, for a measurement that wants to be sure what it is reading.

    Matched on the marker rather than on emptiness: python-pptx fills comments
    with «generated using python-pptx» of its own accord, so "not empty" would
    have called a foreign deck stamped."""
    comments = Presentation(deck_path).core_properties.comments or ""
    return comments if comments.startswith(STAMP_MARKER) else "(no stamp)"


def build_plan(spec, outline, blocks):
    """(block, slide_idx) pairs the way the real pipeline builds them.

    Public because every ad-hoc script that assembles a plan by hand has got it
    wrong the same way: iter42 ignored the plan's `count`, iter56 skipped the
    character budgets, and an experiment in iter85 skipped them again and read
    the renderer's own autofit shrink as a product defect. The correct path has
    to be the easy one.
    """
    item_chars = {s["idx"]: s.get("item_chars")
                  for fam in spec["families"] for s in fam["slides"]}
    plan = []
    seen_role = {}
    for item, slide_idx, count in plan_from_outline(outline, spec)[0]:
        seen_role[item["role"]] = seen_role.get(item["role"], -1) + 1
        block = vary_block(blocks[item["role"]], seen_role[item["role"]])
        if count:  # honour the capacity the matcher computed
            for field in ("bullets", "stats", "left_points", "right_points"):
                if field in block:
                    block[field] = list(block[field])[:count]
        block = _enforce_text_budgets(
            block, item["role"], bullet_char_budget(item_chars.get(slide_idx)))
        plan.append((block, slide_idx))
    return plan


# A repeated role must carry DIFFERENT content, or the deck is duplicate by
# construction and every duplicate-related criterion fires on the fixture rather
# than on the product: --repeat 3 reported «3 повторяющихся заголовка» and
# «смысловые дубли» on all four templates, which said nothing except that the
# probe fed the same block three times.
_REPEAT_THEMES = ["что мешало", "что сделали", "что дальше", "что проверили", "что осталось"]
_ITEM_POOL = [
    "Ручной сбор показателей из семи независимых систем",
    "Разные форматы выгрузок у каждого подразделения",
    "Согласование занимало до двух недель",
    "Автоматический сбор по расписанию каждую ночь",
    "Единая витрина данных для всех подразделений",
    "Проверка расхождений до публикации отчёта",
    "Обучение команды за две недели без отрыва",
    "Права доступа настраиваются по ролям",
    "История версий отчёта хранится целиком",
]


def vary_block(block, n):
    """The same block, told about a different part of the story."""
    if n == 0:
        return dict(block)
    varied = dict(block)
    theme = _REPEAT_THEMES[n % len(_REPEAT_THEMES)]
    if varied.get("title"):
        varied["title"] = f"{varied['title']} — {theme}"
    # The items are REPLACED, not suffixed. A suffix is the first thing the
    # per-slot character budget trims away, and then the slides carry identical
    # bullets under different headings — a fixture that merely stops tripping
    # the duplicate check instead of actually varying.
    for field in ("bullets", "left_points", "right_points"):
        if field in varied:
            varied[field] = [_ITEM_POOL[(n * 7 + i) % len(_ITEM_POOL)]
                             for i in range(len(varied[field]))]
    if "stats" in varied:
        varied["stats"] = [[num, _ITEM_POOL[(n * 5 + i) % len(_ITEM_POOL)]]
                           for i, (num, _) in enumerate(varied["stats"])]
    return varied


def repeat_outline(outline, times):
    """The outline with its CONTENT roles repeated — the axis no preset covers.

    A role used twice sends the matcher back to the same template slide, which
    is the clone path, where the nastiest bugs of this project lived (orphaned
    sldId entries and a notesSlide claimed by several slides, both of which made
    PowerPoint demand repair). One deck of each role never exercises it."""
    if times <= 1:
        return outline
    out = []
    for item in outline:
        out.append(item)
        if item["role"] in ("bullet_list", "stats_kpi", "two_column_comparison"):
            out.extend(dict(item) for _ in range(times - 1))
    return out


def boxes_over_at_render_wrap(deck_path, source_path=None, plan=None):
    """How many boxes of OUR deck are predicted to overflow once the renderer's
    earlier wrap is allowed for.

    enforce_text_fits measures against the full box width; the renderer wraps
    2-4% sooner (CLAUDE.md), so a title measured at two lines is drawn on three
    and spills out of its box. Measured on the T-Zh mono repeat deck: the
    heading needs 2.45in in a 1.88in box, and LibreOffice either autofits it
    smaller — silently undoing the size we computed — or lets it run over.

    Counts OUR text only. A template's own text exceeds its frame 27-43 times per
    deck by intent (iter61), so a raw count over the whole deck says nothing
    about us — the first version printed 23 for a six-slide deck and most of it
    belonged to the designer. Text that also appears on the source slide is
    skipped.

    Counts the boxes the EARLY WRAP breaks, and only those. The first version
    counted every box whose estimated text was taller than its frame, and that
    is a different — much commoner, and mostly harmless — thing: the T-Zh
    universal deck scored 15, of which 13 were one-liners missing their frame by
    0.01in (an 8pt «2025» in a 0.12in box, a 24pt «-40%» in a 0.36in one),
    because a designer sizes a box to the cap height, not to the font's full
    line box. Those cannot be made better by any wrap margin — nothing wraps.
    So a box counts only when the safety-reduced width costs it a LINE that the
    full width did not, which is exactly what a margin in enforce_text_fits
    would prevent. Same rule as 1.1 (CLAUDE.md): "estimate taller than frame" is
    not a defect measure.

    Every paragraph is measured at ITS OWN size. Charging them all the box's
    largest run is what produced most of the rest: on a KPI card the 11pt
    caption was wrapped at 24pt, the number's size.

    Reported, not fixed: adding the margin inside enforce_text_fits would shrink
    the designer's boxes too. This number says how often it happens to us, so a
    change can be judged rather than guessed at.
    """
    template_texts = set()
    if source_path and plan:
        source = Presentation(source_path)
        slides = list(source.slides)
        for _, idx in plan:
            if isinstance(idx, int) and 0 <= idx < len(slides):
                template_texts.update(
                    " ".join(sh.text_frame.text.split())
                    for sh in slides[idx].shapes if sh.has_text_frame)

    prs = Presentation(deck_path)
    metrics_for = FontResolver(deck_path, extract_theme(deck_path)).metrics_for
    over = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.width or not shape.height:
                continue
            paragraphs = [(p, next((r.font.size.pt for r in p.runs
                                    if r.font.size and r.text.strip()), None))
                          for p in shape.text_frame.paragraphs if p.text.strip()]
            paragraphs = [(p, size) for p, size in paragraphs if size]
            font = next((r.font.name for p in shape.text_frame.paragraphs
                         for r in p.runs if r.font.name), None)
            metrics = metrics_for(font) if font else None
            if not paragraphs or metrics is None:
                continue
            if " ".join(shape.text_frame.text.split()) in template_texts:
                continue  # the designer's own text, in the designer's own box
            margins = horizontal_margins_in(shape)
            width_in = Emu(shape.width).inches
            costs_a_line = any(
                estimate_wrapped_lines(p.text, width_in * SINGLE_LINE_SAFETY, size,
                                       metrics=metrics, margins_in=margins)
                > estimate_wrapped_lines(p.text, width_in, size,
                                         metrics=metrics, margins_in=margins)
                for p, size in paragraphs)
            if not costs_a_line:
                continue
            top_inset, bottom_inset = vertical_insets_emu(shape)
            usable = Emu(max(0, shape.height - top_inset - bottom_inset)).inches
            drawn = sum(
                estimate_block_height_in(
                    [p.text], width_in * SINGLE_LINE_SAFETY, size, metrics=metrics,
                    line_spacing=p.line_spacing, margins_in=margins)
                for p, size in paragraphs)
            if drawn > usable:
                over += 1
    return over


def canvas_slots_used(deck_path, source_path, plan):
    """The emptiest reused canvas: (filled, offered, slide number).

    A template slide is a GRID of text slots, and we fill a slide by picking
    shapes from it. When we pick one and blank the rest, the slide keeps the
    designer's proportions and loses the design: T-Zh mono's numbered-list slide
    offers six slots — three texts on the right, «01»/«02» beside them, a title
    — and our bullet_list deck filled two of them, so the render is a heading
    and one stray paragraph in an otherwise empty slide.

    Nothing in the harness saw that. Geometry was clean (nothing overflows,
    nothing collides), the evenness criteria exempt sparse slides, and the score
    stayed at 99.5. A number that nobody prints is a number nobody checks, so
    this one is printed as a fact next to "из шаблона" — with the slide named,
    since "2/6" alone is the kind of bare ratio that got misread for twenty
    iterations.

    Chrome (page numbers, the running header, the year) is excluded: it is
    furniture on both sides and would flatter the ratio.
    """
    if not source_path or not plan:
        return None
    source = Presentation(source_path)
    source_slides = list(source.slides)
    deck = Presentation(deck_path)
    deck_slides = list(deck.slides)
    height = deck.slide_height

    def _content_texts(slide, slide_height):
        return sum(1 for sh in slide.shapes
                   if sh.has_text_frame and sh.text_frame.text.strip()
                   and not is_chrome_shape(sh, slide_height))

    worst = None
    for pos, (_, idx) in enumerate(plan):
        if not isinstance(idx, int) or not 0 <= idx < len(source_slides):
            continue  # synthesized: there is no canvas to leave empty
        if pos >= len(deck_slides):
            continue
        offered = _content_texts(source_slides[idx], source.slide_height)
        filled = _content_texts(deck_slides[pos], height)
        if not offered:
            continue
        if worst is None or filled / offered < worst[0] / worst[1]:
            worst = (filled, offered, pos + 1)
    # Only an UNDER-filled canvas is worth reporting. survey-31's reused slide
    # offers one slot and gets two (we add a subtitle shape of our own), and
    # "2/1" as the headline number would read as a defect where there is none.
    return worst if worst and worst[0] < worst[1] else None


def _ink_rows(img, box, size):
    """Rows of the crop that contain ink, as (top, bottom) spans in pixels.

    Ink is "darker or lighter than this crop's own background by a margin", so
    it works on a white slide and on a dark canvas alike. Rows closer together
    than a third of the tallest span belong to the same line — Cyrillic
    descenders and diacritics («ё», «у») otherwise split one line in two.
    """
    from PIL import Image  # noqa: F401  (imported here: the probe runs headless too)

    w, h = img.size
    x0 = max(0, int(w * box[0] / size[0]))
    x1 = min(w, int(w * (box[0] + box[2]) / size[0]))
    y0 = max(0, int(h * box[1] / size[1]))
    y1 = min(h, int(h * (box[1] + box[3]) / size[1]))
    if x1 - x0 < 4 or y1 - y0 < 4:
        return []
    crop = img.crop((x0, y0, x1, y1)).convert("L")
    pixels = crop.tobytes()  # one byte per pixel in "L"; getdata() is deprecated
    cw, ch = crop.size
    background = max(set(pixels), key=pixels.count)
    rows = []
    for row in range(ch):
        line = pixels[row * cw:(row + 1) * cw]
        inked = sum(1 for value in line if abs(value - background) > 40)
        rows.append(inked > cw * 0.005)

    spans = []
    start = None
    for row, has_ink in enumerate(rows + [False]):
        if has_ink and start is None:
            start = row
        elif not has_ink and start is not None:
            spans.append((start, row))
            start = None
    if not spans:
        return []
    merge_gap = max(2, (max(b - a for a, b in spans)) // 3)
    merged = [spans[0]]
    for a, b in spans[1:]:
        if a - merged[-1][1] <= merge_gap:
            merged[-1] = (merged[-1][0], b)
        else:
            merged.append((a, b))
    return merged


def title_ink_cut_by_its_box(deck_path, renders, plan=None):
    """Slides where the title's ink runs into the bottom edge of its own box.

    The harness has been blind exactly where the generator is blind: both count
    lines with the same fontTools model, so when the model was wrong — bold
    titles measured with the regular face, 7-8% narrow on Cyrillic (iter96) —
    every check agreed with the mistake, the deck scored 99.5, and a bullet was
    printed through the title's second line.

    Pixels do not share that model, but they do not come labelled either, and
    two earlier shapes of this measure failed on that:

    * counting ink rows INSIDE the box found nothing — the unpredicted line is
      outside the box by definition;
    * looking BELOW the box fired on every healthy deck, because the next box's
      text is there and pixels cannot say whose line it is.

    What pixels can say without attribution is whether the ink reaches the box's
    own bottom edge. A box sized for the text it holds leaves the last line
    clear of the edge; a box sized for one line while the renderer draws two has
    the second line starting inside it and running off. That is exactly the
    condition that makes everything positioned after the title wrong.

    SYNTHESIZED slides only. On a native slide the title box is the designer's,
    and a designer's text exceeds its frame by intent — the first run of this
    flagged two untouched T-Zh study slides for exactly that, the same false
    alarm CLAUDE.md records three times over.

    Returns [(slide, clearance_in)] — reported, not scored.
    """
    if not renders:
        return []
    from PIL import Image

    synthesized = {pos for pos, (_, idx) in enumerate(plan or [])
                   if not isinstance(idx, int)}
    prs = Presentation(deck_path)
    size = (prs.slide_width, prs.slide_height)
    out = []
    for index, slide in enumerate(prs.slides):
        if index >= len(renders):
            break
        if index not in synthesized:
            continue
        boxes = [s for s in slide.shapes
                 if s.has_text_frame and s.text_frame.text.strip()
                 and s.top is not None and s.width and s.height
                 and not is_chrome_shape(s, prs.slide_height)]
        if len(boxes) < 2:
            continue  # nothing is positioned after this title
        title = min(boxes, key=lambda s: int(s.top))
        with Image.open(renders[index]) as raw:
            img = raw.convert("RGB")
            spans = _ink_rows(img, (int(title.left), int(title.top),
                                    int(title.width), int(title.height)), size)
            crop_px = max(1.0, img.size[1] * int(title.height) / size[1])
        if not spans:
            continue
        clearance = Emu(int(title.height)).inches * (1 - spans[-1][1] / crop_px)
        # Antialiasing and descenders reach a few thousandths past the glyphs;
        # a box that is genuinely cutting a line leaves nothing at all.
        if clearance < 0.02:
            out.append((index + 1, round(clearance, 3)))
    return out


# A cover is deliberately louder than a content slide — that is what cover_pt
# exists for — and a closing echoes it. Comparing them against the content mode
# reported survey-31's 72pt cover as a defect on the first run.
_NOT_CONTENT_ROLES = ("title", "closing")


def title_sizes(deck_path, source_path=None, plan=None):
    """Title point sizes across the deck, and how far they stray from its mode.

    Criterion 9.3 («одинаковые размеры заголовков») is scored by eye alone, and
    the eye is right to be bothered: the universal deck runs 28 / 24 / 29 / 42 /
    42 / 42, so «Что мешало собирать отчётность вовремя» reads half the size of
    «Как изменился процесс» two slides later.

    The number is printed WITH the source slide's own size, because the
    difference is inherited, not invented: slide 2 sits on the template's
    numbered-grid canvas, whose heading really is 24pt in the template. Whether
    that is the designer's decision or a size we should normalise is exactly
    what the pair of numbers is for — the same reason the title gap is printed
    against the template's own gap rather than judged alone.

    Returns (sizes, mode, [(slide, ours, theirs)]) for the slides that stray.
    """
    content = {pos for pos, (block, _) in enumerate(plan or [])
               if block.get("type") not in _NOT_CONTENT_ROLES}
    prs = Presentation(deck_path)
    sizes = []
    for index, slide in enumerate(prs.slides):
        if plan and index not in content:
            sizes.append(None)
            continue
        title = _pick_title_shape(slide, set())
        found = None
        if title is not None and title.has_text_frame:
            points = [r.font.size.pt for p in title.text_frame.paragraphs
                      for r in p.runs if r.font.size and r.text.strip()]
            found = max(points) if points else None
        sizes.append(found)
    known = [s for s in sizes if s]
    if not known:
        return None
    mode = max(set(known), key=known.count)

    source_sizes = {}
    if source_path and plan:
        source = Presentation(source_path)
        slides = list(source.slides)
        for pos, (_, idx) in enumerate(plan):
            if isinstance(idx, int) and 0 <= idx < len(slides):
                title = _pick_title_shape(slides[idx], set())
                if title is not None and title.has_text_frame:
                    points = [r.font.size.pt for p in title.text_frame.paragraphs
                              for r in p.runs if r.font.size and r.text.strip()]
                    if points:
                        source_sizes[pos] = max(points)

    stray = [(i + 1, size, source_sizes.get(i))
             for i, size in enumerate(sizes)
             if size and abs(size - mode) > mode * 0.25]
    return known, mode, stray


def title_length_vs_template(deck_path, source_path):
    """(our median title length, the template's) in characters.

    Why it is printed next to the sizes: no sizing rule can make a title look
    native when it is twice the length the template's own headings are. Measured
    over the corpus — designers' medians 13 / 16 / 17 / 29 characters, ours 33 on
    every template — and that gap is what turns every attempt at keeping the
    designer's type size into a choice between microtype and a wall of words
    (iter101: universal's wide box improved at 42pt, mono's closing became three
    lines of 61pt).

    The fix this points at is the TITLE BUDGET, per template, taken from the
    designer's own headings — not another rule about sizes.
    """
    if not source_path:
        return None

    def _median_len(path):
        lengths = []
        for slide in Presentation(path).slides:
            title = _pick_title_shape(slide, set())
            if title is None or not title.has_text_frame:
                continue
            text = " ".join(title.text_frame.text.split())
            if text:
                lengths.append(len(text))
        return round(statistics.median(lengths)) if lengths else None

    ours, theirs = _median_len(deck_path), _median_len(source_path)
    if ours is None or theirs is None:
        return None
    return ours, theirs


def _median_title_gap_pct(path):
    """Median gap between a slide's title and the content under it, as a share
    of the slide height. None when no slide offers two content boxes."""
    prs = Presentation(path)
    height = prs.slide_height
    gaps = []
    for slide in prs.slides:
        boxes = []
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.text_frame.text.strip():
                continue
            if shape.top is None or shape.height is None:
                continue
            if is_chrome_shape(shape, height):
                continue
            boxes.append((int(shape.top), int(shape.top) + int(shape.height)))
        if len(boxes) < 2:
            continue
        boxes.sort()
        gaps.append((boxes[1][0] - boxes[0][1]) / height * 100)
    return round(statistics.median(gaps)) if gaps else None


def title_gap_vs_template(deck_path, source_path):
    """(ours, template's) median title-to-content gap, in percent of the slide.

    An OPEN question, printed rather than decided. Synthesized slides centre
    their block in the band that is left, which on short content opens a hole:
    ours run 24-30% of the slide height on survey-31 where the template's own
    slides run 0-8%. Медианы of the five real templates are 0 / 1 / 4 / 4 / 20
    percent — so "hug the title" would be as flat a rule as "always centre",
    and mono genuinely does hold its content low.

    Capping the drop at the template's own gap was tried (iter95) and reverted:
    it contradicts test_synth_center, which records the opposite decision made
    on a real deck — a comparison glued to the title with the lower half of the
    slide empty. Both are true at once, and that is the point of printing the
    two numbers: the designer's content is BOTH glued to the title and dense,
    while ours is sparse, so neither rule alone settles it.
    """
    if not source_path:
        return None
    ours = _median_title_gap_pct(deck_path)
    theirs = _median_title_gap_pct(source_path)
    if ours is None or theirs is None:
        return None
    return ours, theirs


def probe(name, template_id, blocks, out_root, outline=OUTLINE, stamp=None):
    src = os.path.join(TEMPLATES, f"{template_id}.pptx")
    archetypes, pngs = _cached_parse(template_id)
    if archetypes is None:
        # A skipped template must leave NOTHING behind. survey-69 has been
        # skipped since iter53 (its cached parses are poisoned) while its deck
        # and eval.json from an older run sat in the output directory — and were
        # read as current during a later measurement. Same trap as iter69, where
        # a crashed run left the previous deck in place and it looked like the
        # fix had done nothing: an artifact that may not have been rewritten is
        # not evidence.
        stale = os.path.join(out_root, name)
        if os.path.isdir(stale):
            shutil.rmtree(stale)
        return {"template": name, "skipped": "нет разобранного шаблона в output/loop"}

    profile = build_measured_profile(pngs, archetypes) if pngs else None
    spec = build_spec(src, archetypes, style_profile=profile)
    plan = build_plan(spec, outline, blocks)

    # Cleared BEFORE generating for the same reason it is cleared on a skip: a
    # run that dies half-way leaves the previous deck and renders in place, and
    # they read as the result of the run that just failed. That happened for
    # real in iter69 — a SyntaxError killed the probe, the old renders were
    # inspected, and the conclusion "the change did nothing" was right only by
    # accident.
    out_dir = os.path.join(out_root, name)
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    deck = os.path.join(out_dir, "deck.pptx")
    generate(src, plan, deck, synth_canvas=_synth_canvas_hints(src, plan, profile),
             canvas_backgrounds=(profile or {}).get("backgrounds"))
    _stamp(deck, *(stamp or ("?", None)))
    renders = render_pptx_to_pngs(deck, os.path.join(out_dir, "render"))
    # WITH the roles, as evaluation/loop.py does. Without them the evenness
    # criteria treat the cover and the closing as content slides and score the
    # deck uneven for having a cover — the very false positive the exemption
    # exists to prevent. The probe measured differently from the real loop for
    # thirty-five iterations, which made every dop_distribution/dop_pacing
    # number in those reports pessimistic.
    slide_roles = {pos: block["type"] for pos, (block, _) in enumerate(plan)}
    result = evaluate_deck(deck, brief="offline probe", label=f"probe_{name}",
                           slide_roles=slide_roles,
                           render_dir=os.path.join(out_dir, "render"),
                           out_json=os.path.join(out_dir, "eval.json"))
    native = sum(1 for _, idx in plan if idx != SYNTHESIZE)
    # How much of the TEMPLATE the parse actually classified. Without this the
    # summary's "из шаблона 1/6" reads as "this template offers one usable
    # slide", and that is how it was read for twenty iterations — survey-31's
    # best cached parse labels 1 of its 31 slides and calls the other 30
    # "other", i.e. the classification all but failed. loop._is_failed_parse
    # only rejects a parse where EVERY slide is "other", so this one slips
    # through and every content slide gets synthesized.
    classified = _classified(archetypes)
    weak = {cid: v["score"] for cid, v in result["scores"].items()
            if v.get("score") is not None and v["score"] <= 3}
    return {
        "template": name, "deck": deck, "renders": renders,
        "total_100": result["total_100"],
        "native": f"{native}/{len(plan)}",
        "parse": f"{classified}/{len(archetypes)}",
        "over_at_wrap": boxes_over_at_render_wrap(deck, src, plan),
        "slots": canvas_slots_used(deck, src, plan),
        "title_gap": title_gap_vs_template(deck, src),
        "title_ink": title_ink_cut_by_its_box(deck, renders, plan),
        "title_sizes": title_sizes(deck, src, plan),
        "title_len": title_length_vs_template(deck, src),
        "weak": weak,
        "unmeasured": [i + 1 for i, n in
                       enumerate(result["contrast"].get("unmeasured_boxes") or []) if n],
        "unmeasured_why": sorted({w for ws in (result["contrast"].get("unmeasured_why") or [])
                                  for w in ws}),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--content", choices=("long", "short", "wordy"), default="long")
    ap.add_argument("--only", default=None, help="probe one template by name")
    ap.add_argument("--repeat", type=int, default=1,
                    help="repeat each content role N times — a longer deck, and the only "
                         "way this probe reaches the clone path")
    ap.add_argument("--items", type=int, default=None,
                    help="how many bullets the brief asks for (default: the outline's 3) "
                         "— changes which template slide the matcher picks")
    ap.add_argument("--out", default=OUT_ROOT)
    args = ap.parse_args()

    blocks = {"long": LONG, "short": SHORT, "wordy": WORDY}[args.content]
    outline = OUTLINE
    if args.items:
        blocks = {k: dict(v) for k, v in blocks.items()}
        extra = ["Отчёты собирались вручную в конце месяца",
                 "Данные расходились между системами"]
        pool = list(blocks["bullet_list"]["bullets"]) + extra
        blocks["bullet_list"]["bullets"] = (pool * 3)[:args.items]
        outline = [dict(item, count=args.items) if item["role"] == "bullet_list" else item
                   for item in OUTLINE]
    outline = repeat_outline(outline, args.repeat)
    rows = []
    for name, template_id in REAL_TEMPLATES:
        if args.only and args.only != name:
            continue
        rows.append(probe(name, template_id, blocks, args.out, outline,
                          stamp=(args.content, args.items)))

    print(f"\n# Офлайн-прогон, контент: {args.content}\n")
    for r in rows:
        if r.get("skipped"):
            print(f"- {r['template']:10} пропущен: {r['skipped']}")
            continue
        weak_parse = ""
        if r.get("parse"):
            got, total = (int(x) for x in r["parse"].split("/"))
            weak_parse = f"  разбор {r['parse']:>6}" + ("  ← почти всё «other»" if got * 2 < total else "")
        print(f"- {r['template']:10} итог {r['total_100']:6}  из шаблона {r['native']:>4}{weak_parse}"
              + (f"  слабые: {r['weak']}" if r["weak"] else "  слабых нет")
              + (f"  переполнится при рендерном переносе: {r['over_at_wrap']}"
                 if r.get("over_at_wrap") else "")
              + (f"  слотов канвы {r['slots'][0]}/{r['slots'][1]} (слайд {r['slots'][2]})"
                 if r.get("slots") else "")
              + (f"  зазор под титулом {r['title_gap'][0]}% против {r['title_gap'][1]}% у шаблона"
                 if r.get("title_gap") else "")
              + (f"  заголовки {r['title_len'][0]} знаков против {r['title_len'][1]} у шаблона"
                 if r.get("title_len") and r["title_len"][0] > r["title_len"][1] * 1.4 else "")
              + (("  кегли титулов вразнобой: мода {}pt, но ".format(int(r["title_sizes"][1]))
                  + "; ".join(f"слайд {n} {int(ours)}pt"
                              + (f" (у шаблона там {int(theirs)}pt)" if theirs else " (синтез)")
                              for n, ours, theirs in r["title_sizes"][2]))
                 if r.get("title_sizes") and r["title_sizes"][2] else "")
              + (("  БОКС ТИТУЛА РЕЖЕТ ТЕКСТ НА РЕНДЕРЕ: "
                  + "; ".join(f"слайд {n} (запас {gap}in)" for n, gap in r["title_ink"]))
                 if r.get("title_ink") else "")
              + (f"  НЕ ПРОВЕРЕН контраст: {r['unmeasured']}"
                 + (f" ({'; '.join(r['unmeasured_why'])})" if r.get("unmeasured_why") else "")
                 if r["unmeasured"] else ""))
    print("\nПОСМОТРИ ГЛАЗАМИ — числа выше не заменяют взгляд на слайд:")
    for r in rows:
        for path in r.get("renders", []):
            print("    " + path)


if __name__ == "__main__":
    main()
