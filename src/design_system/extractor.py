from common.json_utils import extract_json
from common.model_fallback import TEXT_MODELS, VISION_MODELS, call_with_model_fallback
from design_system.clustering import (
    cluster_slides,
    slide_features,
    slide_fingerprint,
    structural_distance,
)

# Loose-cluster probing (plan 9.5): spread threshold measured on the real
# merged case (T-Ж cover/number/thanks cluster: farthest member at 0.273 from
# the medoid; genuinely homogeneous families sit well under 0.2).
LOOSE_CLUSTER_MIN_SPREAD = 0.2
LOOSE_PROBE_COUNT = 2
LOOSE_PROBE_MIN_MEMBERS = 3

ARCHETYPES = [
    "title", "section_divider", "bullet_list", "stats_kpi",
    "two_column_comparison", "image_caption", "quote", "agenda", "closing", "other",
]

VISION_PROMPT = """Перед тобой рендер слайда презентации.
Определи архетип слайда — строго одно значение из списка:
{archetypes}

Ответь ТОЛЬКО валидным JSON без пояснений и markdown-разметки, по схеме:
{{"archetype": "...", "regions": [{{"role": "...", "position": "..."}}]}}
""".format(archetypes=", ".join(ARCHETYPES))

TEXT_PROMPT_TEMPLATE = """Ниже — структурное описание слайда презентации (типы фигур, их
расположение и текст), без изображения. Определи архетип слайда — строго одно значение из списка:
__ARCHETYPES__

Если структуры недостаточно, чтобы уверенно определить архетип (например, слайд состоит в основном
из графики/картинок, которых в описании не видно), поставь "confidence": "low".

Ответь ТОЛЬКО валидным JSON без пояснений и markdown-разметки, по схеме:
{"archetype": "...", "confidence": "high" | "low"}

Структура слайда:
---
__DESCRIPTION__
---
"""


def _largest_size_pt(shape):
    # A picture's "text" is None, not an empty list.
    sizes = [
        run["size_pt"] for para in (shape.get("text") or []) for run in para
        if run.get("size_pt") and (run.get("text") or "").strip()
    ]
    return max(sizes) if sizes else None


def _describe_slide(slide_struct, slide_size_in=None):
    # Deliberately NO layout name in the description: exporters routinely bind
    # every slide to a layout literally named "TITLE" (seen on a real T-Ж
    # template), and the model anchors on that word over the actual shape
    # structure — a 5-item bullet list was confidently classified "title" with
    # the layout line present and confidently "bullet_list" without it. Same
    # lesson as clustering: layout_name is exporter noise, not signal.
    #
    # Two facts the description used to drop, both decisive and both already in
    # the struct. Measured on a real customer template (survey-31): its title
    # slide is one 8.15x3.2in box reading "Employee Short Survey Results" set at
    # 92pt, plus a small decorative picture — unmistakable to the eye, and
    # classified "other", which makes the deck's most visible slide unusable.
    # What the model was given could not distinguish it from a caption: no font
    # size at all, and a box size with no canvas to measure it against (8.15in
    # is 61% of a 13.33in slide and 100% of a 8in one). Clustering is not to
    # blame — that slide is its own cluster, so the classifier judged it alone.
    lines = []
    if slide_size_in:
        lines.append(f'Slide: {slide_size_in["width"]}x{slide_size_in["height"]} in')
    lines.append(f"Shapes: {len(slide_struct['shapes'])}")
    for shape in slide_struct["shapes"]:
        role = shape.get("placeholder_type") or shape["shape_type"]
        geo = shape["geometry_in"]
        # PARAGRAPH STRUCTURE, not a flattened string. Everything above was
        # measured for survey-31 and its parse is still 1 slide of 31: the
        # reason is here. Its content slides are "title + one text box", and
        # that box holds 4, 5, 7, 8 separate paragraphs — unmistakable bullet
        # lists — which this line joined into a single 80-character run-on.
        # The model was being asked to tell a list from prose with the list-ness
        # removed, so it shrugged, and every deck built on that template got
        # every content slide synthesized.
        snippet = ""
        paragraph_note = ""
        if shape["text"]:
            paragraphs = [
                " ".join(run["text"] for run in para if run["text"]).strip()
                for para in shape["text"]
            ]
            paragraphs = [para for para in paragraphs if para]
            if len(paragraphs) > 1:
                paragraph_note = f", {len(paragraphs)} paragraphs"
                snippet = " | ".join(para[:60] for para in paragraphs[:3])
            else:
                snippet = (paragraphs[0] if paragraphs else "")[:80]
        # Only when the file states it: Google-Slides exports leave most runs
        # unsized, and inventing a number there would be worse than silence
        # (the "title = biggest font" heuristic is blind on those decks — see
        # _pick_title_shape for what that already costs us).
        size_pt = _largest_size_pt(shape)
        font = f", font {size_pt:g}pt" if size_pt else ""
        lines.append(
            f'- {role} at ({geo["left"]}, {geo["top"]}) '
            f'size {geo["width"]}x{geo["height"]}{font}{paragraph_note}: "{snippet}"')
    return "\n".join(lines)


def classify_slide(client, image_path, models=VISION_MODELS):
    """Vision-based classification from a rendered PNG. Expensive (image tokens +
    a vision-capable model tier) — used only as a fallback, see classify_slide_from_structure."""
    def call(model):
        result = client.ask_about_image(image_path, VISION_PROMPT, model=model, max_tokens=500)
        return extract_json(result["choices"][0]["message"]["content"])

    parsed = call_with_model_fallback(call, models)
    if parsed.get("archetype") not in ARCHETYPES:
        parsed["archetype"] = "other"
    return parsed


def classify_slide_from_structure(client, slide_struct, models=TEXT_MODELS, slide_size_in=None):
    """Cheap, image-free classification from the OOXML structure alone (geometry,
    placeholder roles, text). No vision-tier model or image tokens needed."""
    prompt = (
        TEXT_PROMPT_TEMPLATE
        .replace("__ARCHETYPES__", ", ".join(ARCHETYPES))
        .replace("__DESCRIPTION__", _describe_slide(slide_struct, slide_size_in))
    )

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=300)
        return extract_json(result["choices"][0]["message"]["content"])

    parsed = call_with_model_fallback(call, models)
    if parsed.get("archetype") not in ARCHETYPES:
        parsed["archetype"] = "other"
        parsed["confidence"] = "low"
    return parsed


def build_archetype_map_from_images(client, rendered_png_paths, models=VISION_MODELS):
    """Classifies every single rendered slide via vision. Kept for the case where no
    parsed structure is available. Prefer build_archetype_map (structure-first) —
    this is O(n) vision calls, i.e. the expensive path."""
    archetype_map = {}
    for i, path in enumerate(rendered_png_paths):
        try:
            classification = classify_slide(client, path, models=models)
            archetype_map[i] = classification["archetype"]
        except Exception:
            # One slide misbehaving (network hiccup, an unexpected response shape,
            # anything) must not take the rest of the template down with it —
            # "other" is always a safe, already-anticipated fallback here.
            archetype_map[i] = "other"
    return archetype_map


def build_archetype_map(client, template_struct, rendered_png_paths=None,
                         text_models=TEXT_MODELS, vision_models=VISION_MODELS,
                         fingerprint_cache=None, progress_cb=None):
    """Structure-first archetype classification:
    1. Cluster slides by structural similarity + visual veto (free, no LLM) —
       decks repeat the same handful of layout recipes across dozens of slides,
       rarely exactly (see design_system.clustering for why equality fails).
    2. Check the cross-template fingerprint cache — a geometry already
       classified on any earlier template costs nothing here.
    3. On a miss, classify the cluster's medoid via a cheap text-only call (no
       image, no vision-tier model, uses the plentiful base GigaChat quota).
    4. Only if that comes back low-confidence AND a render is available, fall
       back to a single vision call for the medoid.
    5. Propagate the result to every slide in the cluster; confident non-"other"
       results feed the cache for future templates.

    6. LOOSE clusters (plan 9.5) get probed: visually-uniform decks (identical
       card backgrounds everywhere) can merge structurally-similar but
       semantically different slides — a cover, a big-number slide and a
       thank-you all being "header + heading + a couple of boxes". When the
       cluster's internal spread is high, the farthest members are classified
       too; a disagreement dissolves the cluster into per-slide labels. Costs
       a few extra base-tier calls only on the rare loose cluster.

    progress_cb(done, total), when given, fires before each cluster is
    processed — classification is the long, LLM-bound stretch of template
    upload, and the API layer surfaces this as live progress."""
    slides = template_struct["slides"]
    slide_size = template_struct.get("slide_size_in")
    clusters = cluster_slides(slides, slide_size, rendered_png_paths)

    def classify_one(idx):
        """(archetype, confident) for a single slide, with the standard
        text-first / vision-fallback ladder and per-slide fingerprint cache."""
        features = slide_features(slides[idx], slide_size)
        fp = slide_fingerprint(features)
        cached = fingerprint_cache.get(fp) if fingerprint_cache else None
        if cached is not None:
            return cached, True
        try:
            result = classify_slide_from_structure(
                client, slides[idx], models=text_models, slide_size_in=slide_size)
            archetype = result["archetype"]
            confident = result.get("confidence") != "low"
        except Exception:
            # Any single slide's classification failing (of any kind) degrades
            # to "other" rather than aborting the rest of the template.
            archetype, confident = "other", False
        if not confident and rendered_png_paths:
            try:
                archetype = classify_slide(client, rendered_png_paths[idx], models=vision_models)["archetype"]
                confident = True
            except Exception:
                pass  # keep the structure-based guess (likely "other")
        # "other" is a shrug, not knowledge — caching it would freeze a
        # non-answer into every future template with this geometry.
        if fingerprint_cache and confident and archetype != "other":
            fingerprint_cache.put(fp, archetype)
        return archetype, confident

    archetype_map = {}
    for cluster_no, cluster in enumerate(clusters):
        if progress_cb:
            progress_cb(cluster_no, len(clusters))

        medoid = cluster["medoid"]
        archetype, _ = classify_one(medoid)

        members = [i for i in cluster["indices"] if i != medoid]
        if len(members) >= LOOSE_PROBE_MIN_MEMBERS - 1:
            medoid_features = slide_features(slides[medoid], slide_size)
            distances = {
                i: structural_distance(slide_features(slides[i], slide_size), medoid_features)
                for i in members
            }
            if max(distances.values()) >= LOOSE_CLUSTER_MIN_SPREAD:
                probes = sorted(distances, key=distances.get, reverse=True)[:LOOSE_PROBE_COUNT]
                if any(classify_one(p)[0] != archetype for p in probes):
                    # Disagreement: the cluster is a structural coincidence,
                    # not a family — label every member for itself.
                    for idx in cluster["indices"]:
                        archetype_map[idx], _ = classify_one(idx)
                    continue

        for idx in cluster["indices"]:
            archetype_map[idx] = archetype

    return archetype_map
