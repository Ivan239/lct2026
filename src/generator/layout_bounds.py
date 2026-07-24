"""Derives a safe content area from the template's own real slides, instead of
guessing fixed percentages — a logo or footer baked into the slide master isn't
visible in slide.shapes, but real content on real slides never overlaps it, so
where real content starts/ends is the best evidence we have of where it's safe
to place new content."""

MIN_SHAPE_WIDTH_IN = 1.0
MIN_SHAPE_HEIGHT_IN = 0.3


def _percentile(sorted_values, p):
    if not sorted_values:
        return None
    idx = min(len(sorted_values) - 1, int(len(sorted_values) * p))
    return sorted_values[idx]


def infer_content_bounds(template_struct):
    """Returns {left_in, top_in, right_in, bottom_in} in inches, or None if the
    template has too few real content shapes to infer anything meaningful from."""
    tops, lefts, rights, bottoms = [], [], [], []

    for slide in template_struct["slides"]:
        for shape in slide["shapes"]:
            if not shape["text"]:
                continue
            geo = shape["geometry_in"]
            width, height = geo["width"] or 0, geo["height"] or 0
            # Skip small decorative/icon-sized shapes — they're not representative
            # of where substantial content blocks are placed.
            if width < MIN_SHAPE_WIDTH_IN or height < MIN_SHAPE_HEIGHT_IN:
                continue
            if geo["left"] is None or geo["top"] is None:
                continue
            tops.append(geo["top"])
            lefts.append(geo["left"])
            rights.append(geo["left"] + width)
            bottoms.append(geo["top"] + height)

    if len(tops) < 3:
        return None

    tops.sort()
    lefts.sort()
    rights.sort()
    bottoms.sort()

    return {
        # 10th/90th percentile rather than strict min/max so a single unusual
        # outlier slide doesn't skew the bounds for everything else.
        "top_in": _percentile(tops, 0.1),
        "left_in": _percentile(lefts, 0.1),
        "right_in": _percentile(rights, 0.9),
        "bottom_in": _percentile(bottoms, 0.9),
    }
