from common.synthesis import SYNTHESIZE, SYNTHESIZABLE_TYPES

DEFAULT_COUNT = 3

# How many slots past the requested count a fixed-capacity slide may have and
# still beat a flexible free-form box. Stretching by 1-2 items just means phase B
# writes slightly more content; stretching further forces the model to pad with
# filler, at which point an honest N lines in a plain text box reads better.
STRETCH_BUDGET = 2


def _target_colors(spec, n_blocks):
    """Per-position color targets from the template's discovered rotation
    (spec["rotation"], see style_profile) — [] when the template has none,
    which turns every color preference below into a no-op."""
    rotation = spec.get("rotation")
    if not rotation or not rotation.get("order"):
        return []
    order = rotation["order"]
    bookends = rotation.get("bookends", {})
    start = bookends.get("title")
    offset = order.index(start) if start in order else 0
    targets = [order[(offset + i) % len(order)] for i in range(n_blocks)]
    closing = bookends.get("closing")
    if closing in order and n_blocks > 1:
        targets[-1] = closing
        if n_blocks > 2 and targets[-2] == targets[-1]:
            alternatives = [c for c in order if c != targets[-1]]
            if alternatives:
                targets[-2] = alternatives[0]
    return targets


def _pick(candidates, want, target_color):
    """Capacity first, color second, document order last — never sacrifice a
    fitting capacity for a color, but among equally-fitting members prefer the
    one whose measured background continues the deck's rotation."""
    def color_rank(s):
        return 0 if target_color is not None and s.get("bg") == target_color else 1

    flexible = sorted((s for s in candidates if not s["capacity"]), key=color_rank)
    fitting = sorted(
        (s for s in candidates if s["capacity"] and s["capacity"] >= want),
        key=lambda s: (s["capacity"], color_rank(s)),
    )
    close_fit = sorted(
        (s for s in fitting if s["capacity"] - want <= STRETCH_BUDGET),
        key=lambda s: (color_rank(s), s["capacity"]),
    )
    if close_fit:
        return close_fit[0]
    if flexible:
        return flexible[0]
    if fitting:
        return fitting[0]
    return max(candidates, key=lambda s: (s["capacity"], -color_rank(s)))


def plan_from_outline(outline, spec):
    """Slot-based matching for the two-phase flow: assigns each outline item a
    concrete template slide BEFORE any content text exists, so phase B can
    generate text sized to that slide's real capacity — the inversion of the
    legacy flow (generate → match → resize after the fact).

    outline: [{role, theme, count}] from two_phase.generate_outline.
    spec: template_spec.builder.build_spec output.
    Returns (assignments, skipped): assignments = [(outline_item, slide_idx |
    SYNTHESIZE, final_count)] — final_count is the chosen slide's actual
    capacity, which is what phase B must generate exactly."""
    slides_by_role = {}
    for fam in spec["families"]:
        slides_by_role.setdefault(fam["role"], []).extend(fam["slides"])

    # Color rhythm discovered from the template itself (plan 9.2): each output
    # position gets a target background; member choice prefers (never forces)
    # it. On templates without rotation this list is empty and the logic
    # degrades to pure capacity + round-robin.
    color_targets = _target_colors(spec, len(outline))

    used = set()
    reuse_rotation = {}
    assignments = []
    skipped = []
    for position, item in enumerate(outline):
        want = item.get("count") or DEFAULT_COUNT
        target_color = color_targets[position] if color_targets else None
        # Prefer a slide no other block has claimed, but when the family is
        # exhausted, reuse one — the generator clones repeats from the pristine
        # template (see generator._clone_slide), and a cloned native slide
        # keeps the template's icons/typography where a synthesized substitute
        # loses them. SYNTHESIZE remains only for roles with zero native slides.
        # Reuse prefers the rotation's target color, then cycles the family so
        # a deck of clones never reads as a broken record next to the original.
        family = slides_by_role.get(item["role"], [])
        candidates = [s for s in family if s["idx"] not in used]
        # The closing bookend is the one position where color outranks
        # freshness: the deck must close on its brand color (T-Ж opens and
        # closes on yellow), so if no unused member carries it but a used one
        # does, clone the used one rather than break the frame.
        is_closing = color_targets and position == len(outline) - 1
        if is_closing and target_color is not None and not any(
            s.get("bg") == target_color for s in candidates
        ):
            by_color = [s for s in family if s.get("bg") == target_color]
            if by_color:
                candidates = by_color
        if not candidates and family:
            by_color = [s for s in family if target_color is not None and s.get("bg") == target_color]
            if by_color:
                candidates = by_color
            else:
                offset = reuse_rotation.get(item["role"], 0)
                reuse_rotation[item["role"]] = offset + 1
                candidates = [family[offset % len(family)]]
        if candidates:
            best = _pick(candidates, want, target_color)
            used.add(best["idx"])
            final_count = best["capacity"] or want
            assignments.append((item, best["idx"], final_count))
        elif item["role"] in SYNTHESIZABLE_TYPES:
            assignments.append((item, SYNTHESIZE, want))
        else:
            skipped.append(item)
    return assignments, skipped


def match_content_to_slides(content_blocks, archetype_map):
    """archetype_map: {slide_index: archetype}. Returns ordered list of
    (content_block, template_slide_index) — reusing a template slide when an
    archetype has more blocks than slides (the generator clones repeats from
    the pristine template) — or (content_block, SYNTHESIZE) for a synthesizable
    archetype the template has no slide for at all, skipping only blocks whose
    archetype has neither a template slide nor a synthesis path.
    """
    slides_by_archetype = {}
    for idx, archetype in archetype_map.items():
        slides_by_archetype.setdefault(archetype, []).append(idx)

    used = set()
    plan = []
    skipped = []
    for block in sorted(content_blocks, key=lambda b: b["order"]):
        family = slides_by_archetype.get(block["type"], [])
        candidates = [i for i in family if i not in used] or family
        if candidates:
            slide_idx = candidates[0]
            used.add(slide_idx)
            plan.append((block, slide_idx))
        elif block["type"] in SYNTHESIZABLE_TYPES:
            plan.append((block, SYNTHESIZE))
        else:
            skipped.append(block)
    return plan, skipped
