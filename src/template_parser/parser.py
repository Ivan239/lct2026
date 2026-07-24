import zipfile

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
COLOR_SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]


def _clr_value(clr_scheme, slot):
    node = clr_scheme.find(f"a:{slot}", NS)
    if node is None:
        return None
    srgb = node.find("a:srgbClr", NS)
    if srgb is not None:
        return "#" + srgb.get("val")
    sysclr = node.find("a:sysClr", NS)
    if sysclr is not None:
        return "#" + sysclr.get("lastClr")
    return None


def extract_theme(pptx_path, theme_part="ppt/theme/theme1.xml"):
    with zipfile.ZipFile(pptx_path) as z:
        xml_bytes = z.read(theme_part)
    root = etree.fromstring(xml_bytes)
    clr_scheme = root.find(".//a:clrScheme", NS)
    palette = {slot: _clr_value(clr_scheme, slot) for slot in COLOR_SLOTS} if clr_scheme is not None else {}

    font_scheme = root.find(".//a:fontScheme", NS)
    fonts = {}
    if font_scheme is not None:
        for role in ("majorFont", "minorFont"):
            node = font_scheme.find(f"a:{role}/a:latin", NS)
            fonts[role] = node.get("typeface") if node is not None else None

    return {"palette": palette, "fonts": fonts}


def _emu_inches(value):
    return round(Emu(value).inches, 2) if value is not None else None


def _shape_fill(shape):
    try:
        fill = shape.fill
        if fill.type is None:
            return None
        if fill.type == 1:  # MSO_FILL.SOLID
            color = fill.fore_color
            if color.type == 1:  # MSO_THEME_COLOR / RGB check below
                return {"kind": "rgb", "value": str(color.rgb)}
            return {"kind": "theme", "value": str(color.theme_color)}
        return {"kind": str(fill.type)}
    except Exception:
        return None


def _shape_text(shape):
    if not shape.has_text_frame:
        return None
    paragraphs = []
    for p in shape.text_frame.paragraphs:
        run_info = []
        for r in p.runs:
            font = r.font
            color = None
            try:
                if font.color and font.color.type is not None:
                    color = str(font.color.rgb) if font.color.type == 1 else f"theme:{font.color.theme_color}"
            except Exception:
                pass
            run_info.append({
                "text": r.text,
                "font_name": font.name,
                "size_pt": font.size.pt if font.size else None,
                "bold": font.bold,
                "color": color,
            })
        paragraphs.append(run_info)
    return paragraphs


def extract_shape(shape):
    info = {
        "name": shape.name,
        "shape_type": str(shape.shape_type) if shape.shape_type is not None else None,
        "is_placeholder": shape.is_placeholder,
        "geometry_in": {
            "left": _emu_inches(shape.left),
            "top": _emu_inches(shape.top),
            "width": _emu_inches(shape.width),
            "height": _emu_inches(shape.height),
        },
        "fill": _shape_fill(shape),
        "text": _shape_text(shape),
    }
    if shape.is_placeholder:
        info["placeholder_type"] = str(shape.placeholder_format.type)
        info["placeholder_idx"] = shape.placeholder_format.idx
    return info


def extract_slide(slide, index):
    return {
        "index": index,
        "layout_name": slide.slide_layout.name,
        "shapes": [extract_shape(s) for s in slide.shapes],
    }


def extract_template(pptx_path):
    prs = Presentation(pptx_path)
    return {
        "theme": extract_theme(pptx_path),
        "slide_size_in": {"width": _emu_inches(prs.slide_width), "height": _emu_inches(prs.slide_height)},
        "slides": [extract_slide(s, i) for i, s in enumerate(prs.slides)],
    }
