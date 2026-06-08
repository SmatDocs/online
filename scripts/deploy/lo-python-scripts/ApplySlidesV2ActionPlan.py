"""
ApplySlidesV2ActionPlan.py

New V2-only Slides action executor for Collabora/LibreOffice Impress.
This script intentionally does not share handler wiring with the legacy
ApplySlidesActionPlan.py implementation.
"""

import copy
import json
import os
import re
import time
import traceback
import hashlib
import ssl
import sys
import math
import tempfile
from urllib.parse import urlparse
from urllib.request import Request, urlopen


try:
    import uno  # type: ignore
except Exception:  # pragma: no cover - only available in UNO runtime
    uno = None

_BLANK_SLIDE_LAYOUT = 20  # Impress AutoLayout enum: Blank Slide (PPTX layout[6] equivalent)
_SHAPE_TAG_PREFIX = "__smartdocs_shape_tag__:"
_PT_TO_MM100_FACTOR = 35.2777777778
_DEFAULT_CANVAS_WIDTH_PT = 1280.0
_DEFAULT_CANVAS_HEIGHT_PT = 720.0

_CUSTOM_SHAPE_TYPE_MAP = {
    "RECTANGLE": "rectangle",
    "ROUNDED_RECTANGLE": "round-rectangle",
    "OVAL": "ellipse",
    "ELLIPSE": "ellipse",
    "DIAMOND": "diamond",
    "TRIANGLE": "isosceles-triangle",
    "RIGHT_TRIANGLE": "right-triangle",
    "PARALLELOGRAM": "parallelogram",
    "TRAPEZOID": "trapezoid",
    "HEXAGON": "hexagon",
    "OCTAGON": "octagon",
    "PENTAGON": "pentagon",
    "CHEVRON": "chevron",
    "RIGHT_ARROW": "right-arrow",
    "LEFT_ARROW": "left-arrow",
    "UP_ARROW": "up-arrow",
    "DOWN_ARROW": "down-arrow",
    "LEFT_RIGHT_ARROW": "left-right-arrow",
    "UP_DOWN_ARROW": "up-down-arrow",
    "STAR_5_POINT": "star5",
    "FLOWCHART_TERMINATOR": "flowchart-terminator",
    "FLOWCHART_PROCESS": "flowchart-process",
    "FLOWCHART_DECISION": "flowchart-decision",
    "FLOWCHART_DATA": "flowchart-data",
    "FLOWCHART_DOCUMENT": "flowchart-document",
    "FLOWCHART_PREDEFINED_PROCESS": "flowchart-predefined-process",
    "FLOWCHART_OFFPAGE_CONNECTOR": "flowchart-off-page-connector",
    "FLOWCHART_DELAY": "flowchart-delay",
    "FLOWCHART_INTERNAL_STORAGE": "flowchart-internal-storage",
    "FLOWCHART_MANUAL_INPUT": "flowchart-manual-input",
    "FLOWCHART_MANUAL_OPERATION": "flowchart-manual-operation",
    "FLOWCHART_OR": "flowchart-or",
    "FLOWCHART_PREPARATION": "flowchart-preparation",
    "FLOWCHART_CONNECTOR": "flowchart-connector",
    "FLOWCHART_OFF_PAGE_CONNECTOR": "flowchart-off-page-connector",
    "FLOWCHART_MULTIDOCUMENT": "flowchart-multidocument",
    "FLOWCHART_SUMMING_JUNCTION": "flowchart-summing-junction",
    "FLOWCHART_COLLATE": "flowchart-collate",
    "FLOWCHART_SORT": "flowchart-sort",
    "FLOWCHART_EXTRACT": "flowchart-extract",
    "FLOWCHART_MERGE": "flowchart-merge",
    "FLOWCHART_STORED_DATA": "flowchart-stored-data",
    "FLOWCHART_SEQUENTIAL_ACCESS": "flowchart-sequential-access",
    "FLOWCHART_MAGNETIC_DISK": "flowchart-magnetic-disk",
    "FLOWCHART_DIRECT_ACCESS_STORAGE": "flowchart-direct-access-storage",
    "FLOWCHART_DISPLAY": "flowchart-display",
    "FLOWCHART_ALTERNATE_PROCESS": "flowchart-alternate-process",
    "ACTION_BUTTON_INFORMATION": "actionbutton-information",
    "CAN": "can",
    "ROUND_2_DIAG_RECTANGLE": "round-2-diag-rectangle",
    "CROSS": "cross",
    "PLUS": "cross",
    "CUBE": "cube",
    "RING": "ring",
    "DONUT": "ring",
    "PENTAGON_RIGHT": "pentagon-right",
    "HOME_PLATE": "pentagon-right",
    "FORBIDDEN": "forbidden",
    "NO_SYMBOL": "forbidden",
    "BANG": "bang",
    "LIGHTNING": "lightning",
    "LIGHTNING_BOLT": "lightning",
    "HEART": "heart",
    "SMILEY": "smiley",
    "SMILEY_FACE": "smiley",
    "QUAD_ARROW": "quad-arrow",
    "LEFT_ARROW_CALLOUT": "left-arrow-callout",
    "RIGHT_ARROW_CALLOUT": "right-arrow-callout",
    "UP_ARROW_CALLOUT": "up-arrow-callout",
    "DOWN_ARROW_CALLOUT": "down-arrow-callout",
    "LEFT_RIGHT_ARROW_CALLOUT": "left-right-arrow-callout",
    "UP_DOWN_ARROW_CALLOUT": "up-down-arrow-callout",
    "QUAD_ARROW_CALLOUT": "quad-arrow-callout",
    "LEFT_BRACKET": "left-bracket",
    "RIGHT_BRACKET": "right-bracket",
    "LEFT_BRACE": "left-brace",
    "RIGHT_BRACE": "right-brace",
    "BRACKET_PAIR": "bracket-pair",
    "BRACE_PAIR": "brace-pair",
    "RECTANGULAR_CALLOUT": "rectangular-callout",
    "ROUND_RECTANGULAR_CALLOUT": "round-rectangular-callout",
    "ROUNDED_RECTANGULAR_CALLOUT": "round-rectangular-callout",
    "ROUND_CALLOUT": "round-callout",
    "CLOUD_CALLOUT": "cloud-callout",
    "LINE_CALLOUT_1": "line-callout-1",
    "LINE_CALLOUT_2": "line-callout-2",
    "LINE_CALLOUT_3": "line-callout-3",
    "PAPER": "paper",
    "FOLDED_CORNER": "paper",
    "STAR_4_POINT": "star4",
    "STAR4": "star4",
    "STAR5": "star5",
    "STAR_6_POINT": "star6",
    "STAR6": "star6",
    "STAR_8_POINT": "star8",
    "STAR8": "star8",
    "STAR_24_POINT": "star24",
    "STAR24": "star24",
    "STRIPED_RIGHT_ARROW": "striped-right-arrow",
    "NOTCHED_RIGHT_ARROW": "notched-right-arrow",
    "BLOCK_ARC": "block-arc",
    "CIRCULAR_ARROW": "circular-arrow",
    "VERTICAL_SCROLL": "vertical-scroll",
    "HORIZONTAL_SCROLL": "horizontal-scroll",
    "SUN": "sun",
    "MOON": "moon",
    "TEARDROP": "teardrop",
    "SINUSOID": "sinusoid",
    "CLOUD": "cloud",
    "PUZZLE": "puzzle",
    "FLOWER": "flower",
    "FRAME": "frame",
    "OCTAGON_BEVEL": "octagon-bevel",
    "DIAMOND_BEVEL": "diamond-bevel",
    "UP_RIGHT_ARROW": "up-right-arrow",
    "UP_RIGHT_DOWN_ARROW": "up-right-down-arrow",
    "CORNER_RIGHT_ARROW": "corner-right-arrow",
    "SPLIT_ARROW": "split-arrow",
    "SPLIT_ROUND_ARROW": "split-round-arrow",
}

_ARROW_MARKER_MAP = {
    "none": "",
    "arrow": "Arrow",
    "triangle": "Arrow",
    "stealth": "Arrow concave",
    "diamond": "Diamond",
    "oval": "Circle",
    "circle": "Circle",
    "dot": "Circle",
    "square": "Square",
    "open": "Arrow",
}

_IMAGE_FILE_CACHE = {}
_IMAGE_STYLE_CACHE = {}
_DOWNLOAD_HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) SmartDocs/1.0",
}

_UNVERIFIED_SSL_CONTEXT = ssl.create_default_context()
try:
    _UNVERIFIED_SSL_CONTEXT.check_hostname = False
    _UNVERIFIED_SSL_CONTEXT.verify_mode = ssl.CERT_NONE
except Exception:
    pass


def _log(message: str) -> None:
    try:
        print(message, flush=True)
    except Exception:
        pass


def _log_json(prefix: str, value) -> None:
    try:
        _log(f"{prefix}{json.dumps(value, ensure_ascii=False)}")
    except Exception as exc:
        _log(f"{prefix}<json_dump_failed: {exc}>")


def _shape_debug_info(shape):
    info = {"name": None, "canonical_id": None, "service": None}
    if shape is None:
        return info
    try:
        info["name"] = shape.getName() if hasattr(shape, "getName") else None
    except Exception:
        info["name"] = None
    try:
        info["canonical_id"] = _shape_canonical_id(shape)
    except Exception:
        info["canonical_id"] = None
    try:
        info["service"] = shape.getShapeType() if hasattr(shape, "getShapeType") else None
    except Exception:
        info["service"] = None
    return info


def _const(name: str, default=None):
    if uno is None:
        return default
    try:
        return uno.getConstantByName(name)
    except Exception:
        return default


def _pt_to_mm100(value) -> int:
    try:
        return int(float(value) * _PT_TO_MM100_FACTOR)
    except Exception:
        return 0


def _mm100_to_pt(value) -> float:
    try:
        return float(value) / _PT_TO_MM100_FACTOR
    except Exception:
        return 0.0


def _inches_to_mm100(value) -> int:
    try:
        return int(float(value) * 2540.0)
    except Exception:
        return 0


def _hex_to_rgb_int(value):
    if value is None:
        return None
    text = str(value).strip().lstrip("#")
    if len(text) != 6:
        return None
    try:
        return int(text, 16)
    except Exception:
        return None


def _rgb_int_to_tuple(value):
    try:
        rgb = int(value)
    except Exception:
        return None
    return ((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF)


def _rgb_tuple_to_int(rgb_tuple):
    if not isinstance(rgb_tuple, (list, tuple)) or len(rgb_tuple) != 3:
        return None
    try:
        r = max(0, min(255, int(rgb_tuple[0])))
        g = max(0, min(255, int(rgb_tuple[1])))
        b = max(0, min(255, int(rgb_tuple[2])))
    except Exception:
        return None
    return (r << 16) | (g << 8) | b


def _apply_color_brightness(rgb_int, brightness):
    if rgb_int is None or brightness is None:
        return rgb_int
    try:
        brightness = max(-1.0, min(1.0, float(brightness)))
    except Exception:
        return rgb_int
    channels = _rgb_int_to_tuple(rgb_int)
    if channels is None:
        return rgb_int
    if brightness >= 0.0:
        adjusted = tuple(int(round(c + (255 - c) * brightness)) for c in channels)
    else:
        factor = 1.0 + brightness
        adjusted = tuple(int(round(c * factor)) for c in channels)
    return _rgb_tuple_to_int(adjusted)


def _uno_property(name, value):
    prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    prop.Name = str(name)
    prop.Value = value
    return prop


def _set_uno_property_quiet(obj, prop_name: str, value) -> bool:
    if obj is None or not prop_name:
        return False
    try:
        if hasattr(obj, "setPropertyValue"):
            obj.setPropertyValue(prop_name, value)
            return True
    except Exception:
        pass
    try:
        setattr(obj, prop_name, value)
        return True
    except Exception:
        return False


def _set_any_uno_property_quiet(obj, prop_names, value) -> bool:
    if obj is None or not prop_names:
        return False
    for prop_name in prop_names:
        if _set_uno_property_quiet(obj, prop_name, value):
            return True
    return False


def _mark_fill_color_as_direct_rgb(obj) -> bool:
    """Prevent newer Collabora builds from exporting direct RGB fills as theme colors."""
    if obj is None:
        return False
    changed = False
    # Collabora 26.04 gives new Impress shapes an Accent1 FillComplexColor.
    # FillColor alone updates the visible RGB value, but OOXML export still
    # prefers the theme-bound complex color unless the theme index is cleared.
    if _set_any_uno_property_quiet(obj, ("FillColorTheme",), -1):
        changed = True
    _set_any_uno_property_quiet(obj, ("FillColorLumMod",), 10000)
    _set_any_uno_property_quiet(obj, ("FillColorLumOff",), 0)
    return changed


def _non_theme_complex_color_from_shape(shape):
    if shape is None:
        return None
    try:
        _mark_fill_color_as_direct_rgb(shape)
        x_complex = shape.getPropertyValue("FillComplexColor")
        if x_complex is not None and hasattr(x_complex, "getThemeColorType"):
            if int(x_complex.getThemeColorType()) == -1:
                return x_complex
    except Exception:
        pass
    return None


def _mark_line_color_as_direct_rgb(shape) -> bool:
    """Make LineColor win over the default theme LineComplexColor on export."""
    x_complex = _non_theme_complex_color_from_shape(shape)
    if x_complex is None:
        return False
    return _set_any_uno_property_quiet(shape, ("LineComplexColor",), x_complex)


def _safe_int(value, default=0):
    try:
        return int(float(value))
    except Exception:
        return int(default)


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return float(default)


def _shape_type_token(value):
    text = str(value or "").strip().replace("-", "_").replace(" ", "_")
    text = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", text)
    text = re.sub(r"_+", "_", text).strip("_").upper()
    if text == "ELLIPSE":
        return "OVAL"
    if text == "CIRCLE":
        return "OVAL"
    if text == "ROUNDED_RECT":
        return "ROUNDED_RECTANGLE"
    if text == "RECT":
        return "RECTANGLE"
    return text or "RECTANGLE"


def _shape_service_for_payload(payload):
    comp_type = str(payload.get("type") or "").strip().lower()
    if comp_type == "textbox":
        return "com.sun.star.drawing.TextShape"
    if comp_type == "image":
        return "com.sun.star.drawing.GraphicObjectShape"
    if comp_type == "connector":
        return "com.sun.star.drawing.ConnectorShape"
    if comp_type == "shape":
        shape_token = _shape_type_token(payload.get("shape_type"))
        if shape_token in {"OVAL"}:
            return "com.sun.star.drawing.EllipseShape"
        if shape_token in {"LINE"}:
            return "com.sun.star.drawing.LineShape"
        return "com.sun.star.drawing.CustomShape"
    return "com.sun.star.drawing.RectangleShape"


def _custom_shape_type_candidates(shape_type):
    token = _shape_type_token(shape_type)
    if not token:
        return []

    # Prefer explicit aliases first, then progressively looser fallbacks.
    candidates = []
    mapped = _CUSTOM_SHAPE_TYPE_MAP.get(token)
    if mapped:
        candidates.append(mapped)

    token_hyphen = token.lower().replace("_", "-")
    token_compact = token.lower().replace("_", "")
    candidates.extend((token_hyphen, token_compact))

    if token.startswith("FLOWCHART_"):
        tail = token.split("_", 1)[1].lower().replace("_", "-")
        candidates.extend((f"flowchart-{tail}", tail))
    if token.startswith("ACTION_BUTTON_"):
        tail = token.split("_", 2)[2].lower().replace("_", "-")
        candidates.extend((f"actionbutton-{tail}", f"action-button-{tail}", tail))
    if token == "ROUND_2_DIAG_RECTANGLE":
        candidates.extend(("round-2-same-rectangle", "round2-diag-rectangle", "round-2diag-rectangle"))

    # Deduplicate while preserving order.
    deduped = []
    seen = set()
    for candidate in candidates:
        text = str(candidate or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        deduped.append(text)
    return deduped


def _set_custom_shape_type(shape, shape_type):
    token = _shape_type_token(shape_type)
    for custom_type in _custom_shape_type_candidates(shape_type):
        try:
            geometry = (_uno_property("Type", custom_type),)
            shape.CustomShapeGeometry = geometry
            return True
        except Exception:
            continue
    _log(f"V2/set_custom_shape_type: no match token={token} raw={shape_type}")
    return False


def _path_or_url_to_file_url(raw_path_or_url):
    text = str(raw_path_or_url or "").strip()
    if not text:
        return None
    parsed = urlparse(text)
    if parsed.scheme in {"http", "https", "file", "private"}:
        return text
    try:
        return uno.systemPathToFileUrl(text)
    except Exception:
        return text


def _build_image_request(url):
    return Request(str(url or "").strip(), headers=dict(_DOWNLOAD_HEADERS))


def _urlopen_with_headers(url, timeout=20):
    return urlopen(
        _build_image_request(url),
        timeout=timeout,
        context=_UNVERIFIED_SSL_CONTEXT,
    )


def _guess_image_extension_from_url(image_url):
    parsed = urlparse(str(image_url or ""))
    path = parsed.path or ""
    _, ext = os.path.splitext(path)
    ext = ext.lower()
    if ext and re.match(r"^\.[a-z0-9]{1,10}$", ext):
        return ext
    return ".img"


def _is_vector_image_ref(image_ref):
    parsed = urlparse(str(image_ref or ""))
    path = parsed.path or ""
    _, ext = os.path.splitext(path)
    return ext.lower() == ".svg"


def _try_import_slide_image_normalizer():
    try:
        from services.slides.image_normalizer import prepare_image_source_for_slide  # type: ignore

        return prepare_image_source_for_slide
    except Exception:
        pass

    try:
        here = os.path.abspath(__file__)
    except Exception:
        here = None
    current = os.path.dirname(here) if here else None
    repo_root = None
    while current:
        candidate = os.path.join(
            current, "backend", "services", "slides", "image_normalizer.py"
        )
        if os.path.exists(candidate):
            repo_root = current
            break
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent

    if repo_root:
        backend_path = os.path.join(repo_root, "backend")
        if backend_path not in sys.path:
            sys.path.insert(0, backend_path)
        try:
            from services.slides.image_normalizer import prepare_image_source_for_slide  # type: ignore

            return prepare_image_source_for_slide
        except Exception:
            pass
    return None


def _normalize_downloaded_image_content(image_url, content):
    ext = _guess_image_extension_from_url(image_url)
    if not content:
        return content, ext
    if _is_vector_image_ref(image_url):
        return content, ext

    prepare_image_source_for_slide = _try_import_slide_image_normalizer()
    if prepare_image_source_for_slide is None:
        return content, ext

    try:
        prepared = prepare_image_source_for_slide(content)
        payload = getattr(prepared, "payload", None)
        prepared_ext = str(getattr(prepared, "extension", "") or "").strip().lower()
        if hasattr(payload, "getvalue"):
            content = payload.getvalue()
        elif isinstance(payload, bytes):
            content = payload
        if prepared_ext:
            ext = prepared_ext
    except Exception as exc:
        _log(f"V2/normalize_downloaded_image_content: normalization skipped: {exc}")
    return content, ext


def _download_image_to_cache(image_url):
    raw = str(image_url or "").strip()
    if not raw:
        return None
    parsed = urlparse(raw)
    if parsed.scheme not in {"http", "https"}:
        return None

    if raw in _IMAGE_FILE_CACHE:
        cached_path = _IMAGE_FILE_CACHE.get(raw)
        if cached_path and os.path.exists(cached_path):
            return cached_path

    cache_dir = os.path.join(tempfile.gettempdir(), "smartdocs_lo_image_cache")
    try:
        os.makedirs(cache_dir, exist_ok=True)
    except Exception:
        pass

    digest = hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest()
    ext = _guess_image_extension_from_url(raw)
    local_path = os.path.join(cache_dir, f"{digest}{ext}")
    if not os.path.exists(local_path):
        try:
            with _urlopen_with_headers(raw, timeout=20) as response:  # noqa: S310
                content = response.read()
            content, normalized_ext = _normalize_downloaded_image_content(raw, content)
            if normalized_ext and normalized_ext != ext:
                local_path = os.path.join(cache_dir, f"{digest}{normalized_ext}")
            with open(local_path, "wb") as handle:
                handle.write(content)
        except Exception as exc:
            _log(f"V2/download_image_to_cache: failed url={raw}: {exc}")
            return None

    _IMAGE_FILE_CACHE[raw] = local_path
    return local_path


def _system_path_to_file_url(system_path):
    if not system_path:
        return None
    try:
        return uno.systemPathToFileUrl(system_path)
    except Exception:
        return system_path


def _image_ref_to_local_path(image_ref):
    raw = str(image_ref or "").strip()
    if not raw:
        return None
    parsed = urlparse(raw)
    if parsed.scheme in {"http", "https"}:
        return _download_image_to_cache(raw)
    if parsed.scheme == "file":
        try:
            return uno.fileUrlToSystemPath(raw)
        except Exception:
            return raw
    if parsed.scheme:
        return None
    return raw if os.path.exists(raw) else None


def _load_image_bytes(image_ref):
    local_path = _image_ref_to_local_path(image_ref)
    if local_path and os.path.exists(local_path):
        try:
            with open(local_path, "rb") as handle:
                return handle.read()
        except Exception:
            return None

    raw = str(image_ref or "").strip()
    parsed = urlparse(raw)
    if parsed.scheme in {"http", "https"}:
        try:
            with _urlopen_with_headers(raw, timeout=20) as response:  # noqa: S310
                return response.read()
        except Exception as exc:
            _log(f"V2/load_image_bytes: failed url={raw}: {exc}")
    return None


def _try_import_image_processor():
    try:
        from services.slides.image_utils import process_image  # type: ignore

        return process_image
    except Exception:
        pass

    # Macro runtime may not have repo backend on sys.path; bootstrap best-effort.
    try:
        here = os.path.abspath(__file__)
    except Exception:
        here = None
    current = os.path.dirname(here) if here else None
    repo_root = None
    while current:
        candidate = os.path.join(current, "backend", "services", "slides", "image_utils.py")
        if os.path.exists(candidate):
            repo_root = current
            break
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    if repo_root:
        backend_path = os.path.join(repo_root, "backend")
        if backend_path not in sys.path:
            sys.path.insert(0, backend_path)
        try:
            from services.slides.image_utils import process_image  # type: ignore

            return process_image
        except Exception:
            pass
    return None


def _shape_target_pixels(shape, dpi_multiplier=3.0):
    if shape is None:
        return 0, 0
    try:
        size = shape.getSize()
        width_pt = _mm100_to_pt(size.Width)
        height_pt = _mm100_to_pt(size.Height)
        width_px = max(1, _safe_int(width_pt * float(dpi_multiplier), 1))
        height_px = max(1, _safe_int(height_pt * float(dpi_multiplier), 1))
        return width_px, height_px
    except Exception:
        return 0, 0


def _image_style_signature(style):
    if not isinstance(style, dict):
        return ""
    keep = {
        "object_fit": style.get("object_fit"),
        "focal_point": style.get("focal_point"),
        "corner_radius": style.get("corner_radius"),
        "is_circle": style.get("is_circle"),
        "shape": style.get("shape"),
        "transparency": style.get("transparency"),
    }
    try:
        return json.dumps(keep, sort_keys=True, ensure_ascii=True)
    except Exception:
        return str(keep)


def _styled_image_cache_file(image_ref, style, width_px, height_px):
    cache_key = f"{image_ref}|{_image_style_signature(style)}|{width_px}x{height_px}"
    digest = hashlib.sha256(cache_key.encode("utf-8", errors="ignore")).hexdigest()
    cache_dir = os.path.join(tempfile.gettempdir(), "smartdocs_lo_image_cache")
    try:
        os.makedirs(cache_dir, exist_ok=True)
    except Exception:
        pass
    return os.path.join(cache_dir, f"{digest}.png")


def _resolve_image_graphic_url(image_ref, image_style=None, shape=None):
    style = image_style if isinstance(image_style, dict) else {}
    object_fit = str(style.get("object_fit") or "contain").strip().lower()
    focal_point = str(style.get("focal_point") or "center").strip().lower()
    is_circle = bool(style.get("is_circle")) or str(style.get("shape") or "").strip().lower() == "circle"
    corner_radius = _safe_int(style.get("corner_radius"), 0)
    transparency = style.get("transparency")
    should_process = (
        object_fit != "contain"
        or is_circle
        or corner_radius > 0
        or (transparency is not None and float(transparency) > 0.0)
    )
    if should_process and _is_vector_image_ref(image_ref):
        should_process = False

    if should_process and shape is not None:
        process_image = _try_import_image_processor()
        if process_image is not None:
            width_px, height_px = _shape_target_pixels(shape)
            source_bytes = _load_image_bytes(image_ref)
            if width_px > 0 and height_px > 0 and source_bytes:
                pil_opacity = 1.0 - max(0.0, min(1.0, _safe_float(transparency, 0.0)))
                styled_path = _styled_image_cache_file(image_ref, style, width_px, height_px)
                if not os.path.exists(styled_path):
                    try:
                        processed_bytes = process_image(
                            source_bytes,
                            target_width=width_px,
                            target_height=height_px,
                            object_fit=object_fit,
                            focal_point=focal_point,
                            corner_radius=max(0, corner_radius),
                            opacity=pil_opacity,
                            is_circle=is_circle,
                        )
                        with open(styled_path, "wb") as handle:
                            handle.write(processed_bytes)
                    except Exception as exc:
                        _log(f"V2/resolve_image_graphic_url: style processing failed: {exc}")
                        styled_path = None
                if styled_path and os.path.exists(styled_path):
                    _IMAGE_STYLE_CACHE[image_ref] = styled_path
                    return _system_path_to_file_url(styled_path)

    local_path = _image_ref_to_local_path(image_ref)
    if local_path and os.path.exists(local_path):
        return _system_path_to_file_url(local_path)

    return _path_or_url_to_file_url(image_ref)


def _query_graphic_from_url(graphic_url):
    if uno is None or not graphic_url:
        return None
    try:
        context = XSCRIPTCONTEXT.getComponentContext()  # type: ignore # NOQA
    except Exception:
        return None
    try:
        provider = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider",
            context,
        )
        return provider.queryGraphic((_uno_property("URL", str(graphic_url)),))
    except Exception as exc:
        _log(f"V2/query_graphic_from_url: failed url={graphic_url}: {exc}")
        return None


def _get_draw_pages(document):
    return document.getDrawPages()


def _extract_canvas_from_blueprint(blueprint):
    if not isinstance(blueprint, dict):
        return None
    root = blueprint.get("root")
    if not isinstance(root, dict):
        return None
    geometry = root.get("geometry")
    if not isinstance(geometry, dict):
        return None
    width = geometry.get("width")
    height = geometry.get("height")
    try:
        width = float(width) if width is not None else None
    except Exception:
        width = None
    try:
        height = float(height) if height is not None else None
    except Exception:
        height = None
    if width and height and width > 0 and height > 0:
        return (width, height)
    return None


def _infer_canvas_size_from_actions(actions):
    if not isinstance(actions, list):
        return (_DEFAULT_CANVAS_WIDTH_PT, _DEFAULT_CANVAS_HEIGHT_PT)
    for action in actions:
        if not isinstance(action, dict):
            continue
        blueprint = action.get("blueprint")
        dims = _extract_canvas_from_blueprint(blueprint)
        if dims:
            return dims
        plan = action.get("plan")
        if isinstance(plan, dict):
            slides = plan.get("slides")
            if isinstance(slides, list):
                for slide_plan in slides:
                    if not isinstance(slide_plan, dict):
                        continue
                    candidate = slide_plan.get("blueprint")
                    dims = _extract_canvas_from_blueprint(candidate)
                    if dims:
                        return dims
    return (_DEFAULT_CANVAS_WIDTH_PT, _DEFAULT_CANVAS_HEIGHT_PT)


def _ensure_document_canvas_size(document, draw_pages, width_pt, height_pt):
    width_mm100 = max(1, _pt_to_mm100(width_pt))
    height_mm100 = max(1, _pt_to_mm100(height_pt))

    applied_draw_pages = 0
    if draw_pages is not None:
        try:
            for i in range(draw_pages.getCount()):
                page = draw_pages.getByIndex(i)
                width_ok = _set_any_uno_property_quiet(page, ("Width",), int(width_mm100))
                height_ok = _set_any_uno_property_quiet(page, ("Height",), int(height_mm100))
                if width_ok or height_ok:
                    applied_draw_pages += 1
        except Exception as exc:
            _log(f"V2/ensure_canvas: draw page sizing failed: {exc}")

    applied_master_pages = 0
    try:
        master_pages = document.getMasterPages() if document is not None else None
    except Exception:
        master_pages = None
    if master_pages is not None:
        try:
            for i in range(master_pages.getCount()):
                page = master_pages.getByIndex(i)
                width_ok = _set_any_uno_property_quiet(page, ("Width",), int(width_mm100))
                height_ok = _set_any_uno_property_quiet(page, ("Height",), int(height_mm100))
                if width_ok or height_ok:
                    applied_master_pages += 1
        except Exception as exc:
            _log(f"V2/ensure_canvas: master page sizing failed: {exc}")

    return {
        "width_pt": float(width_pt),
        "height_pt": float(height_pt),
        "width_mm100": int(width_mm100),
        "height_mm100": int(height_mm100),
        "draw_pages_updated": int(applied_draw_pages),
        "master_pages_updated": int(applied_master_pages),
    }


def _get_page_canvas_size(page):
    if page is None:
        return None
    width_mm100 = _property_value(page, "Width")
    height_mm100 = _property_value(page, "Height")
    try:
        width_mm100 = int(width_mm100) if width_mm100 is not None else None
    except Exception:
        width_mm100 = None
    try:
        height_mm100 = int(height_mm100) if height_mm100 is not None else None
    except Exception:
        height_mm100 = None
    if not width_mm100 or not height_mm100 or width_mm100 <= 0 or height_mm100 <= 0:
        return None
    return {
        "width_pt": float(_mm100_to_pt(width_mm100)),
        "height_pt": float(_mm100_to_pt(height_mm100)),
        "width_mm100": int(width_mm100),
        "height_mm100": int(height_mm100),
    }


def _get_document_canvas_size(document, draw_pages, controller=None):
    candidates = []
    if controller is not None:
        try:
            candidates.append(("current_page", controller.getCurrentPage()))
        except Exception:
            pass
    if draw_pages is not None:
        try:
            if draw_pages.getCount() > 0:
                candidates.append(("first_draw_page", draw_pages.getByIndex(0)))
        except Exception:
            pass

    for source, page in candidates:
        dims = _get_page_canvas_size(page)
        if dims:
            dims["source"] = source
            return dims

    return {
        "width_pt": float(_DEFAULT_CANVAS_WIDTH_PT),
        "height_pt": float(_DEFAULT_CANVAS_HEIGHT_PT),
        "width_mm100": int(_pt_to_mm100(_DEFAULT_CANVAS_WIDTH_PT)),
        "height_mm100": int(_pt_to_mm100(_DEFAULT_CANVAS_HEIGHT_PT)),
        "source": "default_canvas",
    }


def _scale_value(value, factor):
    try:
        return float(value) * float(factor)
    except Exception:
        return value


def _scale_mapping_values(mapping, scale_by_key):
    if not isinstance(mapping, dict):
        return
    for key, factor in scale_by_key.items():
        if key in mapping and mapping.get(key) is not None:
            mapping[key] = _scale_value(mapping.get(key), factor)


def _scale_component_style(style, scale_x, scale_y, uniform_scale):
    if not isinstance(style, dict):
        return

    _scale_mapping_values(
        style,
        {
            "font_size": uniform_scale,
            "header_font_size": uniform_scale,
            "body_font_size": uniform_scale,
            "letter_spacing": scale_x,
            "character_spacing": scale_x,
            "corner_radius": uniform_scale,
            "cell_margin": uniform_scale,
            "space_before": scale_y,
            "space_after": scale_y,
        },
    )

    padding = style.get("padding")
    if isinstance(padding, dict):
        _scale_mapping_values(
            padding,
            {
                "left": scale_x,
                "right": scale_x,
                "top": scale_y,
                "bottom": scale_y,
            },
        )

    text_style = style.get("text")
    if isinstance(text_style, dict):
        _scale_component_style(text_style, scale_x, scale_y, uniform_scale)

    stroke = style.get("stroke")
    if isinstance(stroke, dict):
        _scale_mapping_values(
            stroke,
            {
                "stroke_width": uniform_scale,
                "width": uniform_scale,
            },
        )

    effects = style.get("effects")
    if isinstance(effects, dict):
        shadow = effects.get("shadow")
        if isinstance(shadow, dict):
            _scale_mapping_values(
                shadow,
                {
                    "distance": uniform_scale,
                    "blur": uniform_scale,
                },
            )
        glow = effects.get("glow")
        if isinstance(glow, dict):
            _scale_mapping_values(glow, {"radius": uniform_scale})
        reflection = effects.get("reflection")
        if isinstance(reflection, dict):
            _scale_mapping_values(
                reflection,
                {
                    "distance": uniform_scale,
                    "blur": uniform_scale,
                },
            )
        soft_edge = effects.get("soft_edge")
        if isinstance(soft_edge, dict):
            _scale_mapping_values(soft_edge, {"radius": uniform_scale})

    column_styles = style.get("column_styles")
    if isinstance(column_styles, list):
        for column_style in column_styles:
            if isinstance(column_style, dict):
                _scale_component_style(column_style, scale_x, scale_y, uniform_scale)


def _scale_table_data(table_data, scale_x, scale_y):
    if not isinstance(table_data, dict):
        return
    col_widths = table_data.get("col_widths")
    if isinstance(col_widths, list):
        table_data["col_widths"] = [_scale_value(value, scale_x) for value in col_widths]
    row_heights = table_data.get("row_heights")
    if isinstance(row_heights, list):
        table_data["row_heights"] = [_scale_value(value, scale_y) for value in row_heights]


def _normalize_table_axis_sizes_to_points(raw_sizes, total_points):
    """
    Normalize table width/height lists to point units.

    We need to support:
    - normalized fractions summing to ~1.0
    - slide/canvas point units
    - legacy inch values
    """
    if not isinstance(raw_sizes, list):
        return []

    values = []
    for raw_value in raw_sizes:
        try:
            value = float(raw_value)
        except Exception:
            continue
        if value > 0:
            values.append(value)

    if not values:
        return []

    try:
        total_points = float(total_points) if total_points is not None else 0.0
    except Exception:
        total_points = 0.0

    total_value = sum(values)
    max_value = max(values)

    if total_points > 0:
        if max_value <= 1.0 + 1e-6 and 0.95 <= total_value <= 1.05:
            return [total_points * (value / total_value) for value in values]

        ratio = total_value / total_points if total_points else 0.0
        if 0.5 <= ratio <= 1.5:
            return values

    return [value * 72.0 for value in values]


def _fit_table_axis_sizes_to_points(raw_sizes, total_points, total_count, body_count=None, header_count=0):
    """
    Normalize table dimensions and make the list match the rendered row/column count.

    Row heights from agents may include all rows or only body rows. If a table
    has a header and the supplied length matches body rows, prepend a header row
    height while keeping the total bounded to the table geometry.
    """
    try:
        total_count = int(total_count)
    except Exception:
        total_count = 0
    if total_count <= 0:
        return []

    sizes = _normalize_table_axis_sizes_to_points(raw_sizes, total_points)
    if not sizes:
        return []

    if len(sizes) == total_count:
        fitted = sizes
    elif header_count and body_count is not None and len(sizes) == int(body_count):
        header_size = sizes[0] if sizes else _safe_float(total_points, 0.0) / total_count
        fitted = [header_size] + sizes
    elif len(sizes) < total_count:
        total_points_float = _safe_float(total_points, 0.0)
        remainder = total_points_float - sum(sizes)
        missing = total_count - len(sizes)
        fallback = (
            max(1.0, remainder / missing)
            if remainder > 0 and missing > 0
            else sizes[-1]
        )
        fitted = [*sizes, *([fallback] * missing)]
    else:
        fitted = sizes[:total_count]

    total_points_float = _safe_float(total_points, 0.0)
    fitted_sum = sum(fitted)
    if total_points_float > 0 and fitted_sum > total_points_float and fitted_sum > 0:
        scale = total_points_float / fitted_sum
        fitted = [value * scale for value in fitted]
    return fitted


def _resolve_table_column_style(column_styles, col_idx):
    """
    Return style for a zero-based column index.

    Public table styles document col_index as 1-based. Payloads containing index
    0 are treated as legacy zero-based input for backwards compatibility.
    """
    if not isinstance(column_styles, list):
        return None

    parsed = []
    has_zero_based_index = False
    for col_style in column_styles:
        if not isinstance(col_style, dict):
            continue
        try:
            raw_idx = int(float(col_style.get("col_index")))
        except Exception:
            continue
        if raw_idx == 0:
            has_zero_based_index = True
        parsed.append((raw_idx, col_style))

    for raw_idx, col_style in parsed:
        normalized_idx = raw_idx if has_zero_based_index else raw_idx - 1
        if normalized_idx == col_idx:
            return col_style
    return None


def _scale_component_payload(component, scale_x, scale_y, uniform_scale):
    if not isinstance(component, dict):
        return

    for section_name in ("geometry", "position", "size"):
        section = component.get(section_name)
        if not isinstance(section, dict):
            continue
        _scale_mapping_values(
            section,
            {
                "x": scale_x,
                "y": scale_y,
                "width": scale_x,
                "height": scale_y,
            },
        )

    _scale_component_style(component.get("style"), scale_x, scale_y, uniform_scale)
    _scale_table_data(component.get("table_data"), scale_x, scale_y)

    text_frame = component.get("text_frame")
    if isinstance(text_frame, dict):
        paragraphs = text_frame.get("paragraphs")
        if isinstance(paragraphs, list):
            for paragraph in paragraphs:
                if not isinstance(paragraph, dict):
                    continue
                _scale_component_style(paragraph, scale_x, scale_y, uniform_scale)
                runs = paragraph.get("runs")
                if isinstance(runs, list):
                    for run in runs:
                        if isinstance(run, dict):
                            _scale_component_style(run, scale_x, scale_y, uniform_scale)

    paragraphs = component.get("paragraphs")
    if isinstance(paragraphs, list):
        for paragraph in paragraphs:
            if not isinstance(paragraph, dict):
                continue
            _scale_component_style(paragraph, scale_x, scale_y, uniform_scale)
            runs = paragraph.get("runs")
            if isinstance(runs, list):
                for run in runs:
                    if isinstance(run, dict):
                        _scale_component_style(run, scale_x, scale_y, uniform_scale)

    for child_key in ("children", "items"):
        children = component.get(child_key)
        if isinstance(children, list):
            for child in children:
                if isinstance(child, dict):
                    _scale_component_payload(child, scale_x, scale_y, uniform_scale)


def _scale_blueprint_to_canvas(blueprint, target_width_pt, target_height_pt):
    if not isinstance(blueprint, dict):
        return blueprint

    source_dims = _extract_canvas_from_blueprint(blueprint)
    source_width_pt = (
        float(source_dims[0])
        if source_dims and source_dims[0] and source_dims[0] > 0
        else float(_DEFAULT_CANVAS_WIDTH_PT)
    )
    source_height_pt = (
        float(source_dims[1])
        if source_dims and source_dims[1] and source_dims[1] > 0
        else float(_DEFAULT_CANVAS_HEIGHT_PT)
    )

    try:
        target_width_pt = float(target_width_pt)
        target_height_pt = float(target_height_pt)
    except Exception:
        return copy.deepcopy(blueprint)

    if (
        source_width_pt <= 0
        or source_height_pt <= 0
        or target_width_pt <= 0
        or target_height_pt <= 0
    ):
        return copy.deepcopy(blueprint)

    scale_x = target_width_pt / source_width_pt
    scale_y = target_height_pt / source_height_pt
    uniform_scale = min(scale_x, scale_y)

    scaled = copy.deepcopy(blueprint)
    root = scaled.get("root")
    if isinstance(root, dict):
        _scale_component_payload(root, scale_x, scale_y, uniform_scale)

    _log(
        "V2/scale_blueprint: "
        f"source=({source_width_pt:.2f},{source_height_pt:.2f}) "
        f"target=({target_width_pt:.2f},{target_height_pt:.2f}) "
        f"scale_x={scale_x:.4f} scale_y={scale_y:.4f} uniform={uniform_scale:.4f}"
    )
    return scaled


def _get_page(draw_pages, slide_index_zero_based: int):
    if slide_index_zero_based < 0 or slide_index_zero_based >= draw_pages.getCount():
        return None
    return draw_pages.getByIndex(slide_index_zero_based)


def _resolve_slide_index(raw_target, draw_pages=None):
    """
    Returns 1-based slide index if available, else None.
    Accepts:
      - {"slide_index": N}
      - {"slide": {"slide_index": N}}
      - {"slide_id": "..."} / {"slide": {"slide_id": "..."}}
    """
    if not isinstance(raw_target, dict):
        return None
    if isinstance(raw_target.get("slide"), dict):
        slide_ref = raw_target.get("slide", {})
        value = slide_ref.get("slide_index")
        slide_id = slide_ref.get("slide_id")
    else:
        value = raw_target.get("slide_index")
        slide_id = raw_target.get("slide_id")
    try:
        if value is not None:
            return int(value)
    except Exception:
        pass

    if draw_pages is not None and slide_id:
        slide_id_norm = _normalize_shape_name(slide_id)
        try:
            for page_index in range(draw_pages.getCount()):
                page = draw_pages.getByIndex(page_index)
                page_name = ""
                try:
                    page_name = page.getName() if hasattr(page, "getName") else ""
                except Exception:
                    page_name = ""
                if not page_name:
                    continue
                if page_name == slide_id or _normalize_shape_name(page_name) == slide_id_norm:
                    return page_index + 1
        except Exception:
            pass
    return None


def _find_page_index(draw_pages, page):
    if draw_pages is None or page is None:
        return None
    try:
        for i in range(draw_pages.getCount()):
            if draw_pages.getByIndex(i) == page:
                return i
    except Exception:
        return None
    return None


def _dispatch_uno(frame, command, props=()):
    try:
        service_manager = XSCRIPTCONTEXT.getComponentContext().ServiceManager  # type: ignore # NOQA
        dispatcher = service_manager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper",
            XSCRIPTCONTEXT.getComponentContext(),  # type: ignore # NOQA
        )
        dispatcher.executeDispatch(frame, command, "", 0, tuple(props) if props else ())
        return True
    except Exception:
        return False


def _force_ui_refresh(controller):
    """Force Impress UI/tile refresh after mutations.

    NOTE: Neither .uno:Refresh (Calc-only) nor .uno:UpdateAllLinks work
    reliably in Impress LOK — both throw dispatch exceptions.  The effective
    tile invalidation is handled by setModified(True) + shape nudge + slide
    selection bounce in the post-unlock sequence.  This function is kept as
    a no-op hook in case a working Impress UNO refresh command is found.
    """
    _log("ApplySlidesV2ActionPlan: _force_ui_refresh (no-op, relying on setModified+nudge)")


def _nudge_shape_invalidation(shape):
    """
    Force object/tile invalidation by tiny geometry nudge and restore.
    This avoids .uno:Repaint while still signaling a visible object mutation.
    """
    if shape is None:
        return False
    # Preferred: position nudge by 1 (1/100 mm), then restore.
    try:
        pos = shape.getPosition()
        p1 = uno.createUnoStruct("com.sun.star.awt.Point")
        p1.X = int(pos.X) + 1
        p1.Y = int(pos.Y)
        shape.setPosition(p1)
        p2 = uno.createUnoStruct("com.sun.star.awt.Point")
        p2.X = int(pos.X)
        p2.Y = int(pos.Y)
        shape.setPosition(p2)
        _log("V2/nudge_shape_invalidation: position nudge applied")
        return True
    except Exception:
        pass
    # Fallback: size nudge by 1 (1/100 mm), then restore.
    try:
        size = shape.getSize()
        s1 = uno.createUnoStruct("com.sun.star.awt.Size")
        s1.Width = int(size.Width) + 1
        s1.Height = int(size.Height)
        shape.setSize(s1)
        s2 = uno.createUnoStruct("com.sun.star.awt.Size")
        s2.Width = int(size.Width)
        s2.Height = int(size.Height)
        shape.setSize(s2)
        _log("V2/nudge_shape_invalidation: size nudge applied")
        return True
    except Exception as exc:
        _log(f"V2/nudge_shape_invalidation: failed: {exc}")
        return False


def _nudge_slide_selection(controller, draw_pages, target_page):
    """
    Force tile invalidation by bouncing selection to a neighbor slide and back.
    This is safer in LOK than relying on .uno:Repaint for Impress.
    """
    if controller is None or draw_pages is None or target_page is None:
        return
    try:
        count = draw_pages.getCount()
    except Exception:
        count = 0
    if count <= 0:
        return
    try:
        target_idx = _find_page_index(draw_pages, target_page)
        current_page = controller.getCurrentPage()
        current_idx = _find_page_index(draw_pages, current_page)
        if target_idx is None:
            controller.setCurrentPage(target_page)
            return
        if count > 1 and current_idx == target_idx:
            alt_idx = target_idx - 1 if target_idx > 0 else target_idx + 1
            if 0 <= alt_idx < count:
                controller.setCurrentPage(draw_pages.getByIndex(alt_idx))
                # Let LOK publish a page-change event before returning.
                time.sleep(0.06)
        controller.setCurrentPage(target_page)
        # Small delay helps Collabora flush tile invalidation on same-slide updates.
        time.sleep(0.03)
    except Exception as exc:
        _log(f"ApplySlidesV2ActionPlan: slide nudge failed: {exc}")


def _focus_shape(document, shape):
    """Best-effort: focus the edited shape to force a visible UI event."""
    if document is None or shape is None:
        return False
    try:
        controller = document.getCurrentController()
    except Exception:
        controller = None
    if controller is None:
        return False
    try:
        controller.select(shape)
        _log("V2/focus_shape: selected edited shape")
        return True
    except Exception as exc:
        _log(f"V2/focus_shape: select failed: {exc}")
        return False


def _move_page_to_index(document, controller, draw_pages, page, target_idx_zero_based):
    current_idx = _find_page_index(draw_pages, page)
    if current_idx is None:
        return None

    target_idx = max(0, min(int(target_idx_zero_based), draw_pages.getCount() - 1))
    if current_idx == target_idx:
        return current_idx

    frame = controller.getFrame() if controller is not None else None
    if frame is None:
        return current_idx

    max_steps = max(0, draw_pages.getCount() * 2)
    steps = 0
    while current_idx != target_idx and steps < max_steps:
        try:
            controller.setCurrentPage(page)
        except Exception:
            break

        command = ".uno:MovePageUp" if current_idx > target_idx else ".uno:MovePageDown"
        if not _dispatch_uno(frame, command):
            break

        next_idx = _find_page_index(draw_pages, page)
        if next_idx is None or next_idx == current_idx:
            break
        current_idx = next_idx
        steps += 1

    return current_idx


def _set_slide_layout_value(slide, layout_value):
    if slide is None:
        return False
    try:
        slide.setPropertyValue("Layout", int(layout_value))
        return True
    except Exception:
        pass
    try:
        slide.Layout = int(layout_value)
        return True
    except Exception:
        return False


def _clear_slide_shapes(slide):
    if slide is None:
        return 0
    try:
        count = int(slide.getCount())
    except Exception:
        return 0

    removed = 0
    for idx in range(count - 1, -1, -1):
        try:
            shape = slide.getByIndex(idx)
            slide.remove(shape)
            removed += 1
        except Exception:
            continue
    return removed


def _prepare_blank_slide_surface(document, slide):
    """
    Mirror PPTX slide_layouts[6] behavior:
    1) assign Impress blank autolayout (20) when possible
    2) clear any existing placeholders/shapes so rendering starts from an empty canvas
    """
    if slide is None:
        return {"layout_assigned": False, "removed_shape_count": 0}

    layout_assigned = _set_slide_layout_value(slide, _BLANK_SLIDE_LAYOUT)

    removed = _clear_slide_shapes(slide)
    _log(
        "V2/prepare_blank_slide_surface: "
        f"layout_assigned={layout_assigned} removed_shape_count={removed}"
    )
    return {"layout_assigned": bool(layout_assigned), "removed_shape_count": int(removed)}


def _normalize_shape_name(value):
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", text)
    text = re.sub(r"([A-Za-z])([0-9])", r"\1_\2", text)
    text = re.sub(r"([0-9])([A-Za-z])", r"\1_\2", text)
    text = re.sub(r"[^A-Za-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_").lower()
    return text


def _canonical_id_from_target(element_id):
    if element_id is None:
        return None
    text = _normalize_shape_name(element_id)
    if re.match(r"^[a-z0-9]+(?:_[a-z0-9]+)*$", text):
        return text
    return None


def _shape_tag_payload_from_value(value):
    text = str(value or "").strip()
    if not text:
        return None
    if text.startswith(_SHAPE_TAG_PREFIX):
        text = text[len(_SHAPE_TAG_PREFIX) :].strip()
    if not (text.startswith("{") and text.endswith("}")):
        return None
    try:
        payload = json.loads(text)
    except Exception:
        return None
    return payload if isinstance(payload, dict) else None


def _canonical_shape_id_from_metadata(value):
    payload = _shape_tag_payload_from_value(value)
    if isinstance(payload, dict):
        tagged_id = payload.get("id")
        canonical = _canonical_id_from_target(tagged_id)
        if canonical:
            return canonical
    return _canonical_id_from_target(value)


def _ancestor_ids_from_metadata(value):
    payload = _shape_tag_payload_from_value(value)
    if not isinstance(payload, dict):
        return []
    raw_anc = payload.get("anc")
    if not isinstance(raw_anc, list):
        return []
    ancestors = []
    for item in raw_anc:
        canonical = _canonical_id_from_target(item)
        if canonical:
            ancestors.append(canonical)
    return ancestors


def _parse_int_from_text(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            iv = int(value)
            return iv if iv > 0 else None
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    matches = re.findall(r"\d+", text)
    if not matches:
        return None
    try:
        iv = int(matches[-1])
        return iv if iv > 0 else None
    except Exception:
        return None


def _property_value(shape, prop_name):
    if shape is None or not prop_name:
        return None
    # XPropertySet getter.
    try:
        return shape.getPropertyValue(prop_name)
    except Exception:
        pass
    # Attribute getter fallback.
    try:
        return getattr(shape, prop_name)
    except Exception:
        return None


def _extract_shape_id_from_grab_bag(grab_bag):
    if grab_bag is None:
        return None
    # Typical UNO grab-bags are sequences of PropertyValue(Name, Value).
    items = grab_bag if isinstance(grab_bag, (list, tuple)) else [grab_bag]
    for item in items:
        name = None
        value = None
        try:
            name = str(getattr(item, "Name", "") or "").lower()
            value = getattr(item, "Value", None)
        except Exception:
            # Some grab-bags may be plain dict-like structures.
            if isinstance(item, dict):
                for key in ("name", "Name"):
                    if key in item:
                        try:
                            name = str(item.get(key) or "").lower()
                        except Exception:
                            name = None
                        break
                for key in ("value", "Value"):
                    if key in item:
                        value = item.get(key)
                        break
        if not name:
            continue
        if any(token in name for token in ("spid", "shapeid", "shape_id", "oox")):
            parsed = _parse_int_from_text(value)
            if parsed is not None:
                return parsed
        # Nested grab-bag fallback.
        if isinstance(value, (list, tuple)):
            parsed = _extract_shape_id_from_grab_bag(value)
            if parsed is not None:
                return parsed
    return None


def _shape_canonical_id(shape):
    if shape is None:
        return None

    # Prefer metadata-aware IDs so JSON-tagged names like
    # {"id":"main_title","anc":[...]} resolve to "main_title".
    try:
        name = str(shape.getName() or "") if hasattr(shape, "getName") else ""
    except Exception:
        name = ""
    canonical_from_name = _canonical_shape_id_from_metadata(name)
    if canonical_from_name:
        return canonical_from_name

    for prop_name in ("Description", "AlternativeText", "Title"):
        canonical_from_prop = _canonical_shape_id_from_metadata(
            _property_value(shape, prop_name)
        )
        if canonical_from_prop:
            return canonical_from_prop

    # Direct integer-ish properties.
    for prop_name in ("ShapeId", "OOXMLShapeId", "spid"):
        parsed = _parse_int_from_text(_property_value(shape, prop_name))
        if parsed is not None:
            return f"id_{parsed}"

    # Interop grab-bags (OOXML metadata).
    for bag_name in ("InteropGrabBag", "ShapeInteropGrabBag"):
        parsed = _extract_shape_id_from_grab_bag(_property_value(shape, bag_name))
        if parsed is not None:
            return f"id_{parsed}"

    return None


def _shape_ancestors(shape):
    if shape is None:
        return []
    try:
        name = str(shape.getName() or "") if hasattr(shape, "getName") else ""
    except Exception:
        name = ""
    ancestors = _ancestor_ids_from_metadata(name)
    if ancestors:
        return ancestors
    for prop_name in ("Description", "AlternativeText", "Title"):
        ancestors = _ancestor_ids_from_metadata(_property_value(shape, prop_name))
        if ancestors:
            return ancestors
    return []


def _set_shape_identity(shape, element_id, ancestors=None):
    canonical = _canonical_id_from_target(element_id)
    if not canonical:
        return False
    changed = False
    try:
        shape.setName(canonical)
        changed = True
    except Exception:
        pass

    normalized_ancestors = []
    if isinstance(ancestors, list):
        for anc in ancestors:
            anc_canonical = _canonical_id_from_target(anc)
            if anc_canonical:
                normalized_ancestors.append(anc_canonical)
    tag = _SHAPE_TAG_PREFIX + json.dumps({"id": canonical, "anc": normalized_ancestors}, ensure_ascii=False)
    for prop_name in ("Description", "AlternativeText"):
        try:
            if hasattr(shape, "setPropertyValue"):
                shape.setPropertyValue(prop_name, tag)
            else:
                setattr(shape, prop_name, tag)
            changed = True
        except Exception:
            continue
    return changed


def _shape_matches_target(shape, canonical_target, include_descendants):
    canonical_shape_id = _shape_canonical_id(shape)
    if canonical_shape_id == canonical_target:
        return True
    if not include_descendants:
        return False
    return canonical_target in set(_shape_ancestors(shape))


def _collect_shapes_for_target(slide, element_id, include_descendants=True):
    canonical_target = _canonical_id_from_target(element_id)
    if slide is None or not canonical_target:
        return []
    matched = []
    try:
        count = slide.getCount()
    except Exception:
        return []
    for idx in range(count):
        try:
            shape = slide.getByIndex(idx)
        except Exception:
            continue
        try:
            if _shape_matches_target(shape, canonical_target, include_descendants):
                matched.append(shape)
        except Exception:
            continue
    return matched


def _remove_shapes(slide, shapes):
    removed = 0
    for shape in list(shapes or []):
        try:
            slide.remove(shape)
            removed += 1
        except Exception:
            continue
    return removed


def _find_shape_by_canonical_id(slide, element_id: str):
    _log(f"V2/find_shape: requested element_id={element_id}")
    if not slide or not element_id:
        _log("V2/find_shape: missing slide or element_id")
        return None, -1
    canonical_target = _canonical_id_from_target(element_id)
    if not canonical_target:
        _log(f"V2/find_shape: invalid canonical target from element_id={element_id}")
        return None, -1
    try:
        count = slide.getCount()
    except Exception:
        _log("V2/find_shape: unable to read slide count")
        return None, -1
    _log(f"V2/find_shape: scanning {count} shapes for target={canonical_target}")

    for i in range(count):
        try:
            shape = slide.getByIndex(i)
            canonical_shape_id = _shape_canonical_id(shape)
            info = _shape_debug_info(shape)
            _log(
                "V2/find_shape: "
                f"idx={i} name={info.get('name')} canonical={canonical_shape_id} service={info.get('service')}"
            )
            if canonical_shape_id == canonical_target:
                # Normalize name to canonical id for future deterministic lookup.
                _set_shape_identity(shape, canonical_target)
                _log(f"V2/find_shape: matched idx={i} target={canonical_target}")
                return shape, i
        except Exception as exc:
            _log(f"V2/find_shape: exception on idx={i}: {exc}")
        except Exception:
            continue
    _log(f"V2/find_shape: no match for target={canonical_target}")
    return None, -1


def _create_point(x_mm100: int, y_mm100: int):
    point = uno.createUnoStruct("com.sun.star.awt.Point")
    point.X = int(x_mm100)
    point.Y = int(y_mm100)
    return point


def _create_size(width_mm100: int, height_mm100: int):
    size = uno.createUnoStruct("com.sun.star.awt.Size")
    size.Width = int(width_mm100)
    size.Height = int(height_mm100)
    return size


def _apply_geometry(shape, payload):
    geometry = payload.get("geometry") if isinstance(payload.get("geometry"), dict) else {}
    position = payload.get("position") if isinstance(payload.get("position"), dict) else {}
    size = payload.get("size") if isinstance(payload.get("size"), dict) else {}

    x = geometry.get("x", position.get("x"))
    y = geometry.get("y", position.get("y"))
    width = geometry.get("width", size.get("width"))
    height = geometry.get("height", size.get("height"))

    changed = False
    try:
        if x is not None or y is not None:
            old = shape.getPosition()
            px = _pt_to_mm100(x) if x is not None else old.X
            py = _pt_to_mm100(y) if y is not None else old.Y
            shape.setPosition(_create_point(px, py))
            changed = True
            _log(f"V2/apply_geometry: position -> X={px} Y={py}")
    except Exception as exc:
        _log(f"V2/apply_geometry: position failed: {exc}")

    try:
        if width is not None or height is not None:
            old = shape.getSize()
            w = _pt_to_mm100(width) if width is not None else old.Width
            h = _pt_to_mm100(height) if height is not None else old.Height
            shape.setSize(_create_size(w, h))
            changed = True
            _log(f"V2/apply_geometry: size -> W={w} H={h}")
    except Exception as exc:
        _log(f"V2/apply_geometry: size failed: {exc}")
    _log(f"V2/apply_geometry: changed={changed}")
    return changed


def _extract_text_from_payload(payload):
    if payload.get("text") is not None:
        return str(payload.get("text"))
    if payload.get("text_content") is not None:
        return str(payload.get("text_content"))

    paragraphs = _extract_payload_paragraphs(payload)
    if paragraphs:
        lines = []
        for para in paragraphs:
            runs = para.get("runs")
            if isinstance(runs, list) and runs:
                pieces = []
                has_explicit_run_text = False
                for run in runs:
                    if not isinstance(run, dict):
                        continue
                    if "text" in run and run.get("text") is not None:
                        pieces.append(str(run.get("text")))
                        has_explicit_run_text = True
                if has_explicit_run_text:
                    lines.append("".join(pieces))
                    continue
            if "text" in para and para.get("text") is not None:
                lines.append(str(para.get("text")))
        if lines:
            return "\n".join(lines)
    return None


def _apply_text(shape, payload):
    if not isinstance(payload, dict):
        return False
    has_explicit_text = (
        payload.get("text") is not None
        or payload.get("text_content") is not None
    )
    if not has_explicit_text:
        paragraphs = _extract_payload_paragraphs(payload)
        has_explicit_text = _paragraphs_have_explicit_text(paragraphs)
    if not has_explicit_text:
        _log("V2/apply_text: no explicit text update in payload")
        return False

    text_value = _extract_text_from_payload(payload)
    if text_value is None:
        _log("V2/apply_text: no text in payload")
        return False
    _log(f"V2/apply_text: attempting text_len={len(text_value)}")
    try:
        if hasattr(shape, "getText"):
            shape.getText().setString(text_value)
            _log("V2/apply_text: applied via getText().setString()")
            return True
    except Exception as exc:
        _log(f"V2/apply_text: getText().setString() failed: {exc}")
    try:
        if hasattr(shape, "setString"):
            shape.setString(text_value)
            _log("V2/apply_text: applied via setString()")
            return True
    except Exception as exc:
        _log(f"V2/apply_text: setString() failed: {exc}")
    _log("V2/apply_text: no text-capable method succeeded")
    return False


def _resolved_text_style(payload):
    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    text_style = {}
    if isinstance(style.get("text"), dict):
        text_style.update(style.get("text"))
    text_style.update(style)
    return text_style


def _extract_payload_paragraphs(payload):
    if not isinstance(payload, dict):
        return []
    raw = payload.get("paragraphs")
    if isinstance(raw, list):
        return [item for item in raw if isinstance(item, dict)]
    text_frame = payload.get("text_frame")
    if isinstance(text_frame, dict) and isinstance(text_frame.get("paragraphs"), list):
        return [item for item in text_frame.get("paragraphs") if isinstance(item, dict)]
    return []


def _paragraphs_have_explicit_text(paragraphs):
    if not isinstance(paragraphs, list):
        return False
    for para in paragraphs:
        if not isinstance(para, dict):
            continue
        if "text" in para and para.get("text") is not None:
            return True
        runs = para.get("runs")
        if isinstance(runs, list):
            for run in runs:
                if isinstance(run, dict) and "text" in run and run.get("text") is not None:
                    return True
    return False


def _resolved_first_paragraph(payload):
    paragraphs = _extract_payload_paragraphs(payload)
    if not paragraphs:
        return {}
    first = paragraphs[0]
    return first if isinstance(first, dict) else {}


def _resolved_first_run(paragraph):
    if not isinstance(paragraph, dict):
        return {}
    runs = paragraph.get("runs")
    if not isinstance(runs, list) or not runs:
        return {}
    first = runs[0]
    return first if isinstance(first, dict) else {}


def _resolve_para_adjust(value):
    token = str(value or "").strip().lower()
    mapping = {
        "left": _const("com.sun.star.style.ParagraphAdjust.LEFT", 0),
        "start": _const("com.sun.star.style.ParagraphAdjust.LEFT", 0),
        "center": _const("com.sun.star.style.ParagraphAdjust.CENTER", 3),
        "middle": _const("com.sun.star.style.ParagraphAdjust.CENTER", 3),
        "right": _const("com.sun.star.style.ParagraphAdjust.RIGHT", 1),
        "end": _const("com.sun.star.style.ParagraphAdjust.RIGHT", 1),
        "justify": _const("com.sun.star.style.ParagraphAdjust.BLOCK", 2),
    }
    return mapping.get(token)


def _resolve_vertical_adjust(value):
    token = str(value or "").strip().lower()
    mapping = {
        "top": _const("com.sun.star.drawing.TextVerticalAdjust.TOP", 0),
        "middle": _const("com.sun.star.drawing.TextVerticalAdjust.CENTER", 1),
        "center": _const("com.sun.star.drawing.TextVerticalAdjust.CENTER", 1),
        "bottom": _const("com.sun.star.drawing.TextVerticalAdjust.BOTTOM", 2),
    }
    return mapping.get(token)


def _apply_paragraph_cursor_style(cursor, style, paragraph):
    changed = False
    alignment = _resolve_para_adjust(paragraph.get("alignment") or style.get("text_anchor"))
    if alignment is not None:
        try:
            cursor.ParaAdjust = alignment
            changed = True
        except Exception as exc:
            _log(f"V2/apply_text_style: paragraph alignment failed: {exc}")

    line_spacing = paragraph.get("line_spacing", style.get("line_spacing"))
    if line_spacing is not None:
        try:
            line = uno.createUnoStruct("com.sun.star.style.LineSpacing")
            line.Mode = _const("com.sun.star.style.LineSpacingMode.PROP", 0)
            line.Height = max(1, _safe_int(float(line_spacing) * 100, 120))
            cursor.ParaLineSpacing = line
            changed = True
        except Exception as exc:
            _log(f"V2/apply_text_style: line_spacing failed: {exc}")

    for key, prop in (("space_before", "ParaTopMargin"), ("space_after", "ParaBottomMargin")):
        val = paragraph.get(key)
        if val is None:
            continue
        try:
            setattr(cursor, prop, _pt_to_mm100(val))
            changed = True
        except Exception as exc:
            _log(f"V2/apply_text_style: {key} failed: {exc}")
    return changed


def _apply_run_cursor_style(cursor, run_style):
    changed = False

    try:
        font_name = run_style.get("font_name")
        if font_name:
            cursor.CharFontName = str(font_name)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: font_name failed: {exc}")

    try:
        font_size = run_style.get("font_size")
        if font_size is not None:
            cursor.CharHeight = float(font_size)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: font_size failed: {exc}")

    try:
        color = _hex_to_rgb_int(run_style.get("font_color") or run_style.get("color"))
        if color is not None:
            cursor.CharColor = color
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: color failed: {exc}")

    try:
        letter_spacing = run_style.get("letter_spacing") or run_style.get("character_spacing")
        if letter_spacing is not None:
            cursor.CharKerning = _safe_int(float(letter_spacing) * 100, 0)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: letter_spacing failed: {exc}")

    try:
        if run_style.get("bold") is not None:
            cursor.CharWeight = (
                _const("com.sun.star.awt.FontWeight.BOLD", 150.0)
                if bool(run_style.get("bold"))
                else _const("com.sun.star.awt.FontWeight.NORMAL", 100.0)
            )
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: bold failed: {exc}")

    try:
        if run_style.get("italic") is not None:
            cursor.CharPosture = (
                _const("com.sun.star.awt.FontSlant.ITALIC", 2)
                if bool(run_style.get("italic"))
                else _const("com.sun.star.awt.FontSlant.NONE", 0)
            )
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: italic failed: {exc}")

    try:
        if run_style.get("underline") is not None:
            cursor.CharUnderline = (
                _const("com.sun.star.awt.FontUnderline.SINGLE", 1)
                if bool(run_style.get("underline"))
                else _const("com.sun.star.awt.FontUnderline.NONE", 0)
            )
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: underline failed: {exc}")

    try:
        if "superscript" in run_style or "subscript" in run_style:
            if run_style.get("superscript"):
                cursor.CharEscapement = 33
                cursor.CharEscapementHeight = 58
                changed = True
            elif run_style.get("subscript"):
                cursor.CharEscapement = -33
                cursor.CharEscapementHeight = 58
                changed = True
            else:
                cursor.CharEscapement = 0
                cursor.CharEscapementHeight = 100
                changed = True
    except Exception:
        pass

    try:
        has_hyperlink_key = "hyperlink_url" in run_style or "hyperlink" in run_style
        if has_hyperlink_key:
            link = run_style.get("hyperlink_url")
            if link is None and isinstance(run_style.get("hyperlink"), dict):
                link = run_style.get("hyperlink", {}).get("target")
            if link is not None and str(link) != "":
                cursor.HyperLinkURL = str(link)
            elif hasattr(cursor, "HyperLinkURL"):
                cursor.HyperLinkURL = ""
            changed = True
    except Exception:
        pass

    return changed


def _apply_rich_text_payload(text_obj, style, paragraphs):
    if text_obj is None or not paragraphs:
        return False

    try:
        text_obj.setString("")
        cursor = text_obj.createTextCursor()
    except Exception as exc:
        _log(f"V2/apply_text_style: rich text init failed: {exc}")
        return False

    changed = False
    paragraph_break = _const("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK", 0)

    for para_index, paragraph in enumerate(paragraphs):
        para = paragraph if isinstance(paragraph, dict) else {}
        runs = para.get("runs") if isinstance(para.get("runs"), list) else []
        if not runs:
            runs = [{"text": para.get("text", "")}]

        for run in runs:
            run_data = run if isinstance(run, dict) else {"text": str(run)}
            run_text = str(run_data.get("text", ""))
            run_style = dict(style)
            for key, value in para.items():
                if key != "runs":
                    run_style[key] = value
            run_style.update(run_data)

            if _apply_paragraph_cursor_style(cursor, style, para):
                changed = True
            if _apply_run_cursor_style(cursor, run_style):
                changed = True
            try:
                text_obj.insertString(cursor, run_text, False)
                if run_text:
                    changed = True
            except Exception as exc:
                _log(f"V2/apply_text_style: insert run text failed: {exc}")

        if para_index < len(paragraphs) - 1:
            try:
                text_obj.insertControlCharacter(cursor, paragraph_break, False)
                changed = True
            except Exception as exc:
                _log(f"V2/apply_text_style: paragraph break failed: {exc}")

    return changed


def _apply_text_style(shape, payload):
    style = _resolved_text_style(payload)
    paragraph = _resolved_first_paragraph(payload)
    first_run = _resolved_first_run(paragraph)
    run_style = dict(style)
    run_style.update(first_run)
    text_obj = None
    try:
        if hasattr(shape, "getText"):
            text_obj = shape.getText()
    except Exception as exc:
        _log(f"V2/apply_text_style: shape.getText() failed: {exc}")
        text_obj = None
    if text_obj is None or not hasattr(text_obj, "createTextCursor"):
        _log("V2/apply_text_style: no text cursor support on target shape")
        return False

    changed = False
    try:
        cursor = text_obj.createTextCursor()
    except Exception as exc:
        _log(f"V2/apply_text_style: createTextCursor failed: {exc}")
        return False

    try:
        v_align = _resolve_vertical_adjust(style.get("vertical_anchor"))
        if v_align is not None:
            shape.TextVerticalAdjust = v_align
            changed = True
    except Exception as exc:
        _log(f"V2/apply_text_style: vertical_anchor failed: {exc}")

    try:
        text_fit = style.get("auto_fit")
        if text_fit is not None:
            shape.TextAutoGrowHeight = bool(text_fit)
            changed = True
    except Exception:
        pass

    try:
        padding = style.get("padding") if isinstance(style.get("padding"), dict) else {}
        if padding:
            if padding.get("left") is not None:
                shape.TextLeftDistance = max(0, _pt_to_mm100(padding.get("left")))
                changed = True
            if padding.get("right") is not None:
                shape.TextRightDistance = max(0, _pt_to_mm100(padding.get("right")))
                changed = True
            if padding.get("top") is not None:
                shape.TextUpperDistance = max(0, _pt_to_mm100(padding.get("top")))
                changed = True
            if padding.get("bottom") is not None:
                shape.TextLowerDistance = max(0, _pt_to_mm100(padding.get("bottom")))
                changed = True
    except Exception:
        pass

    try:
        wrap_value = style.get("word_wrap")
        if wrap_value is not None:
            shape.TextWordWrap = bool(wrap_value)
            changed = True
    except Exception:
        pass

    paragraphs = _extract_payload_paragraphs(payload)
    if paragraphs and _paragraphs_have_explicit_text(paragraphs):
        if _apply_rich_text_payload(text_obj, style, paragraphs):
            changed = True
        _log(f"V2/apply_text_style: changed={changed}")
        return changed

    if not style:
        _log("V2/apply_text_style: no style payload")
        return changed

    if _apply_run_cursor_style(cursor, run_style):
        changed = True
    if _apply_paragraph_cursor_style(cursor, style, paragraph):
        changed = True

    _log(f"V2/apply_text_style: changed={changed}")
    return changed


def _apply_style(shape, payload):
    comp_type = str(payload.get("type") or "").strip().lower()
    is_text_component = comp_type in {"textbox", "text", "title", "subtitle"}

    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    fill = style.get("fill") if isinstance(style.get("fill"), dict) else {}
    stroke = style.get("stroke") if isinstance(style.get("stroke"), dict) else {}
    effects = style.get("effects") if isinstance(style.get("effects"), dict) else {}

    # update_style payloads may provide nested fragments at top-level
    if isinstance(payload.get("fill"), dict):
        merged_fill = dict(fill)
        merged_fill.update(payload.get("fill"))
        fill = merged_fill
    if isinstance(payload.get("stroke"), dict):
        merged_stroke = dict(stroke)
        merged_stroke.update(payload.get("stroke"))
        stroke = merged_stroke
    if isinstance(payload.get("effects"), dict):
        merged_effects = dict(effects)
        merged_effects.update(payload.get("effects"))
        effects = merged_effects

    # Flat fallback keys (update_style / partial replace payloads)
    fill = dict(fill)
    stroke = dict(stroke)
    effects = dict(effects)
    for source in (style, payload):
        if not isinstance(source, dict):
            continue
        for key in (
            "fill_type",
            "fill_color",
            "background_color",
            "fill_transparency",
            "fill_brightness",
            "gradient",
            "transparency",
        ):
            if key in source and fill.get(key) is None:
                fill[key] = source.get(key)
        # Generic `style.color` is commonly text color for textbox components.
        # Only map it to shape fill for non-text components.
        if (
            not is_text_component
            and "color" in source
            and fill.get("fill_color") is None
            and fill.get("color") is None
        ):
            fill["color"] = source.get("color")
        for key in (
            "stroke_color",
            "stroke_brightness",
            "stroke_width",
            "width",
            "no_line",
            "line_style",
            "arrow_start",
            "arrow_end",
        ):
            if key in source and stroke.get(key) is None:
                stroke[key] = source.get(key)
        if "effects" in source and not effects and isinstance(source.get("effects"), dict):
            effects = dict(source.get("effects"))

    changed = False

    try:
        fill_type = str(fill.get("fill_type") or fill.get("type") or "").strip().lower()
        if fill_type in {"none", "no_fill"}:
            shape.FillStyle = _const("com.sun.star.drawing.FillStyle.NONE", 0)
            changed = True
        elif fill_type == "gradient":
            gradient = fill.get("gradient") if isinstance(fill.get("gradient"), dict) else {}
            stop_1 = _hex_to_rgb_int(gradient.get("stop_1_color") or fill.get("fill_color") or fill.get("color"))
            stop_2 = _hex_to_rgb_int(gradient.get("stop_2_color") or fill.get("fill_color") or fill.get("color"))
            brightness = fill.get("fill_brightness")
            stop_1 = _apply_color_brightness(stop_1, brightness)
            stop_2 = _apply_color_brightness(stop_2, brightness)
            stop_1_pos = max(0.0, min(1.0, _safe_float(gradient.get("stop_1_position"), 0.0)))
            stop_2_pos = max(0.0, min(1.0, _safe_float(gradient.get("stop_2_position"), 1.0)))
            grad = uno.createUnoStruct("com.sun.star.awt.Gradient")
            grad.Style = _const("com.sun.star.awt.GradientStyle.LINEAR", 0)
            grad.StartColor = stop_1 if stop_1 is not None else 0xFFFFFF
            grad.EndColor = stop_2 if stop_2 is not None else grad.StartColor
            grad.Angle = _safe_int(float(gradient.get("angle", 90.0)) * 10.0, 900)
            # UNO does not support arbitrary gradient stops; approximate stop positions.
            grad.Border = max(0, min(100, _safe_int(stop_1_pos * 100.0, 0)))
            grad.XOffset = max(0, min(100, _safe_int(stop_2_pos * 100.0, 100)))
            grad.YOffset = 50
            grad.StartIntensity = 100
            grad.EndIntensity = 100
            grad.StepCount = 0
            shape.FillStyle = _const("com.sun.star.drawing.FillStyle.GRADIENT", 2)
            shape.FillGradient = grad
            changed = True
        elif fill_type and any(fill.get(key) is not None for key in ("fill_color", "background_color", "color")):
            shape.FillStyle = _const("com.sun.star.drawing.FillStyle.SOLID", 1)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: fill_type failed: {exc}")

    try:
        fill_color = _hex_to_rgb_int(
            fill.get("fill_color") or fill.get("background_color") or fill.get("color")
        )
        fill_color = _apply_color_brightness(fill_color, fill.get("fill_brightness"))
        if fill_color is None and fill.get("fill_brightness") is not None:
            fill_color = _apply_color_brightness(_property_value(shape, "FillColor"), fill.get("fill_brightness"))
        if fill_color is not None:
            shape.FillStyle = _const("com.sun.star.drawing.FillStyle.SOLID", 1)
            shape.FillColor = fill_color
            _mark_fill_color_as_direct_rgb(shape)
            shape.FillColor = fill_color
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: fill_color failed: {exc}")

    try:
        fill_transparency = fill.get("fill_transparency")
        if fill_transparency is None:
            fill_transparency = fill.get("transparency")
        if fill_transparency is not None:
            # model uses 0.0=opaque, 1.0=fully transparent
            pct = max(0, min(100, _safe_int(float(fill_transparency) * 100.0, 0)))
            shape.FillTransparence = pct
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: fill_transparency failed: {exc}")

    try:
        if stroke.get("no_line") is True:
            shape.LineStyle = _const("com.sun.star.drawing.LineStyle.NONE", 0)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: no_line failed: {exc}")

    try:
        stroke_color = _hex_to_rgb_int(stroke.get("stroke_color") or stroke.get("color"))
        stroke_color = _apply_color_brightness(stroke_color, stroke.get("stroke_brightness"))
        if stroke_color is None and stroke.get("stroke_brightness") is not None:
            stroke_color = _apply_color_brightness(_property_value(shape, "LineColor"), stroke.get("stroke_brightness"))
        if stroke_color is not None:
            shape.LineStyle = _const("com.sun.star.drawing.LineStyle.SOLID", 1)
            shape.LineColor = stroke_color
            _mark_line_color_as_direct_rgb(shape)
            shape.LineColor = stroke_color
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: stroke_color failed: {exc}")

    try:
        stroke_width = stroke.get("stroke_width") or stroke.get("width")
        if stroke_width is not None:
            # Best-effort: map points to 1/100 mm
            shape.LineWidth = max(1, _pt_to_mm100(stroke_width))
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: stroke_width failed: {exc}")

    try:
        line_style = str(stroke.get("line_style") or "").strip().lower()
        if line_style in {"dash", "dash_dot"}:
            shape.LineStyle = _const("com.sun.star.drawing.LineStyle.DASH", 2)
            dash = uno.createUnoStruct("com.sun.star.drawing.LineDash")
            dash.Style = _const("com.sun.star.drawing.DashStyle.RECT", 1)
            dash.Dots = 1
            dash.DotLen = 120
            dash.Dashes = 1 if line_style == "dash" else 2
            dash.DashLen = 280
            dash.Distance = 180
            shape.LineDash = dash
            changed = True
        elif line_style == "solid":
            shape.LineStyle = _const("com.sun.star.drawing.LineStyle.SOLID", 1)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: line_style failed: {exc}")

    try:
        marker_start = _ARROW_MARKER_MAP.get(str(stroke.get("arrow_start", "")).strip().lower())
        marker_end = _ARROW_MARKER_MAP.get(str(stroke.get("arrow_end", "")).strip().lower())
        if marker_start is not None:
            shape.LineStartName = marker_start
            if marker_start:
                shape.LineStartCenter = True
                shape.LineStartWidth = max(100, shape.LineWidth if hasattr(shape, "LineWidth") else 100)
            changed = True
        if marker_end is not None:
            shape.LineEndName = marker_end
            if marker_end:
                shape.LineEndCenter = True
                shape.LineEndWidth = max(100, shape.LineWidth if hasattr(shape, "LineWidth") else 100)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: arrowheads failed: {exc}")

    try:
        rotation = payload.get("rotation")
        if rotation is None and isinstance(style, dict):
            rotation = style.get("rotation")
        if rotation is not None:
            # UNO RotateAngle is in 1/100 degrees
            shape.RotateAngle = _safe_int(float(rotation) * 100.0, 0)
            changed = True
    except Exception as exc:
        _log(f"V2/apply_style: rotation failed: {exc}")

    try:
        corner_radius = style.get("corner_radius")
        if corner_radius is None:
            corner_radius = payload.get("corner_radius")
        if corner_radius is not None:
            shape.CornerRadius = max(0, _pt_to_mm100(corner_radius))
            changed = True
    except Exception:
        pass

    try:
        shadow = effects.get("shadow") if isinstance(effects.get("shadow"), dict) else {}
        if shadow:
            shadow_color = _hex_to_rgb_int(shadow.get("color") or "000000")
            shadow_changed = False
            if shadow_color is not None:
                if _set_any_uno_property_quiet(shape, ("ShadowColor",), shadow_color):
                    shadow_changed = True
            if _set_any_uno_property_quiet(shape, ("Shadow", "ShadowOn"), True):
                shadow_changed = True
            if shadow.get("distance") is not None:
                distance_mm100 = _pt_to_mm100(shadow.get("distance"))
                angle_deg = _safe_float(shadow.get("angle"), 315.0)
                radians = math.radians(angle_deg)
                dx = int(round(math.cos(radians) * float(distance_mm100)))
                dy = int(round(-math.sin(radians) * float(distance_mm100)))
                if distance_mm100 > 0 and dx == 0 and dy == 0:
                    dx = distance_mm100
                if _set_any_uno_property_quiet(shape, ("ShadowXDistance",), dx):
                    shadow_changed = True
                if _set_any_uno_property_quiet(shape, ("ShadowYDistance",), dy):
                    shadow_changed = True
            if shadow.get("blur") is not None:
                blur_mm100 = max(0, _pt_to_mm100(shadow.get("blur")))
                if _set_any_uno_property_quiet(shape, ("ShadowBlur", "ShadowBlurRadius"), blur_mm100):
                    shadow_changed = True
            if shadow.get("transparency") is not None:
                shadow_pct = max(0, min(100, _safe_int(float(shadow.get("transparency")) * 100.0, 0)))
                if _set_any_uno_property_quiet(shape, ("ShadowTransparence", "ShadowTransparency"), shadow_pct):
                    shadow_changed = True
            if shadow_changed:
                changed = True
    except Exception as exc:
        _log(f"V2/apply_style: effects.shadow failed: {exc}")

    try:
        glow = effects.get("glow") if isinstance(effects.get("glow"), dict) else {}
        if glow:
            glow_changed = False
            if _set_any_uno_property_quiet(shape, ("GlowEffect", "GlowOn"), True):
                glow_changed = True
            glow_color = _hex_to_rgb_int(glow.get("color") or "FFFFFF")
            if glow_color is not None:
                if _set_any_uno_property_quiet(shape, ("GlowEffectColor", "GlowColor"), glow_color):
                    glow_changed = True
            glow_radius = glow.get("radius")
            if glow_radius is not None:
                radius_mm100 = max(0, _pt_to_mm100(glow_radius))
                if _set_any_uno_property_quiet(shape, ("GlowEffectRadius", "GlowRadius"), radius_mm100):
                    glow_changed = True
            if glow.get("transparency") is not None:
                glow_pct = max(0, min(100, _safe_int(float(glow.get("transparency")) * 100.0, 0)))
                if _set_any_uno_property_quiet(shape, ("GlowEffectTransparency", "GlowTransparency"), glow_pct):
                    glow_changed = True
            if glow_changed:
                changed = True
    except Exception as exc:
        _log(f"V2/apply_style: effects.glow failed: {exc}")

    try:
        reflection = effects.get("reflection") if isinstance(effects.get("reflection"), dict) else {}
        if reflection:
            reflection_changed = False
            if _set_any_uno_property_quiet(shape, ("Reflection", "ReflectionOn"), True):
                reflection_changed = True
            if reflection.get("size") is not None:
                if _set_any_uno_property_quiet(shape, ("ReflectionSize",), _safe_int(reflection.get("size"), 100)):
                    reflection_changed = True
            if reflection.get("distance") is not None:
                if _set_any_uno_property_quiet(
                    shape,
                    ("ReflectionDistance",),
                    max(0, _pt_to_mm100(reflection.get("distance"))),
                ):
                    reflection_changed = True
            if reflection.get("blur") is not None:
                if _set_any_uno_property_quiet(
                    shape,
                    ("ReflectionBlur", "ReflectionBlurRadius"),
                    max(0, _pt_to_mm100(reflection.get("blur"))),
                ):
                    reflection_changed = True
            if reflection.get("transparency") is not None:
                reflection_pct = max(
                    0,
                    min(100, _safe_int(float(reflection.get("transparency")) * 100.0, 0)),
                )
                if _set_any_uno_property_quiet(
                    shape,
                    ("ReflectionTransparency", "ReflectionTransparence"),
                    reflection_pct,
                ):
                    reflection_changed = True
            if reflection_changed:
                changed = True
    except Exception as exc:
        _log(f"V2/apply_style: effects.reflection failed: {exc}")

    try:
        soft_edge = effects.get("soft_edge") if isinstance(effects.get("soft_edge"), dict) else {}
        if soft_edge:
            soft_radius = soft_edge.get("radius")
            if soft_radius is not None:
                radius_mm100 = max(0, _pt_to_mm100(soft_radius))
                if _set_any_uno_property_quiet(shape, ("SoftEdgeRadius", "EdgeRadius"), radius_mm100):
                    changed = True
    except Exception as exc:
        _log(f"V2/apply_style: effects.soft_edge failed: {exc}")

    _log(f"V2/apply_style: changed={changed}")
    return changed


def _create_shape_for_payload(document, payload):
    service_name = _shape_service_for_payload(payload)
    try:
        shape = document.createInstance(service_name)
    except Exception as exc:
        _log(f"V2/create_shape_for_payload: service create failed service={service_name}: {exc}")
        shape = document.createInstance("com.sun.star.drawing.RectangleShape")

    if service_name == "com.sun.star.drawing.CustomShape":
        _set_custom_shape_type(shape, payload.get("shape_type"))
    _initialize_new_shape_defaults(shape)
    return shape


def _initialize_new_shape_defaults(shape):
    try:
        shape.FillStyle = _const("com.sun.star.drawing.FillStyle.NONE", 0)
    except Exception:
        pass
    try:
        shape.LineStyle = _const("com.sun.star.drawing.LineStyle.NONE", 0)
    except Exception:
        pass


def _style_has_visual_surface(style):
    if not isinstance(style, dict) or not style:
        return False

    fill = style.get("fill") if isinstance(style.get("fill"), dict) else style
    fill_type = str(fill.get("fill_type") or fill.get("type") or "").strip().lower()
    if fill_type and fill_type not in {"none", "no_fill"}:
        return True
    if any(fill.get(key) is not None for key in ("fill_color", "color", "background_color")):
        return True
    if isinstance(fill.get("gradient"), dict) and fill.get("gradient"):
        return True

    stroke = style.get("stroke") if isinstance(style.get("stroke"), dict) else {}
    if not stroke:
        stroke = style
    if isinstance(stroke, dict) and stroke.get("no_line") is not True:
        stroke_width = stroke.get("stroke_width", stroke.get("width"))
        if stroke.get("stroke_color") is not None or stroke.get("line_style") is not None:
            return True
        if stroke_width is not None:
            try:
                if float(stroke_width) > 0:
                    return True
            except Exception:
                return True

    effects = style.get("effects")
    return isinstance(effects, dict) and bool(effects)


def _payload_has_explicit_visual_fill(payload):
    comp_type = str(payload.get("type") or "").strip().lower()
    is_text_component = comp_type in {"textbox", "text", "title", "subtitle"}

    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    style_fill = style.get("fill") if isinstance(style.get("fill"), dict) else {}
    payload_fill = payload.get("fill") if isinstance(payload.get("fill"), dict) else {}
    fill_sources = [style_fill, payload_fill]
    flat_sources = [style, payload]

    for source in fill_sources:
        if not isinstance(source, dict):
            continue
        fill_type = str(source.get("fill_type") or source.get("type") or "").strip().lower()
        if fill_type in {"none", "no_fill"}:
            return False
        if fill_type == "gradient":
            return True
        gradient = source.get("gradient")
        if isinstance(gradient, dict) and gradient:
            return True
        for key in ("fill_color", "background_color"):
            if _hex_to_rgb_int(source.get(key)) is not None:
                return True
        if _hex_to_rgb_int(source.get("color")) is not None:
            return True

    for source in flat_sources:
        if not isinstance(source, dict):
            continue
        fill_type = str(source.get("fill_type") or source.get("type") or "").strip().lower()
        if fill_type in {"none", "no_fill"}:
            return False
        if fill_type == "gradient":
            return True
        gradient = source.get("gradient")
        if isinstance(gradient, dict) and gradient:
            return True
        for key in ("fill_color", "background_color"):
            if _hex_to_rgb_int(source.get(key)) is not None:
                return True
        if not is_text_component and _hex_to_rgb_int(source.get("color")) is not None:
            return True
    return False


def _payload_has_explicit_visual_stroke(payload):
    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    style_stroke = style.get("stroke") if isinstance(style.get("stroke"), dict) else {}
    payload_stroke = payload.get("stroke") if isinstance(payload.get("stroke"), dict) else {}
    stroke_sources = [style_stroke, payload_stroke]
    flat_sources = [style, payload]

    for source in stroke_sources:
        if not isinstance(source, dict):
            continue
        if source.get("no_line") is True:
            return False
        if _hex_to_rgb_int(source.get("stroke_color") or source.get("color")) is not None:
            return True

    for source in flat_sources:
        if not isinstance(source, dict):
            continue
        if source.get("no_line") is True:
            return False
        if _hex_to_rgb_int(source.get("stroke_color")) is not None:
            return True
    return False


def _payload_has_explicit_visual_stroke_color(payload):
    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    style_stroke = style.get("stroke") if isinstance(style.get("stroke"), dict) else {}
    payload_stroke = payload.get("stroke") if isinstance(payload.get("stroke"), dict) else {}
    sources = [style_stroke, payload_stroke, style, payload]

    for source in sources:
        if not isinstance(source, dict):
            continue
        if source.get("no_line") is True:
            return False
        if _hex_to_rgb_int(source.get("stroke_color") or source.get("color")) is not None:
            return True
    return False


_OFFICE_DEFAULT_BLUE_COLORS = {
    0x4F81BD,  # Office 2007 accent1
    0x4472C4,  # Office 2013+ accent1
    0x3465A4,  # LibreOffice classic accent1
}


def _payload_identifier_text(payload, element_id=None):
    parts = []
    if element_id:
        parts.append(str(element_id))
    for key in ("id", "element_id", "name", "component_id"):
        value = payload.get(key) if isinstance(payload, dict) else None
        if value:
            parts.append(str(value))
    return " ".join(parts).lower().replace("-", "_")


def _payload_shape_type_text(payload):
    return str((payload or {}).get("shape_type") or "").strip().lower().replace("-", "_")


def _payload_text_content(payload):
    if not isinstance(payload, dict):
        return ""
    parts = []
    for key in ("text", "content", "label"):
        value = payload.get(key)
        if isinstance(value, str):
            parts.append(value)
    text_frame = payload.get("text_frame")
    if isinstance(text_frame, dict):
        parts.append(str(text_frame))
    return " ".join(parts).strip()


def _payload_color_values(payload, include_fill=True, include_stroke=True):
    if not isinstance(payload, dict):
        return []
    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    sources = []
    if include_fill:
        if isinstance(style.get("fill"), dict):
            sources.append(style.get("fill"))
        if isinstance(payload.get("fill"), dict):
            sources.append(payload.get("fill"))
    if include_stroke:
        if isinstance(style.get("stroke"), dict):
            sources.append(style.get("stroke"))
        if isinstance(payload.get("stroke"), dict):
            sources.append(payload.get("stroke"))
    sources.extend([style, payload])

    values = []
    for source in sources:
        if not isinstance(source, dict):
            continue
        for key in ("fill_color", "background_color", "stroke_color", "color"):
            if key in source:
                values.append(source.get(key))
        gradient = source.get("gradient")
        if isinstance(gradient, dict):
            values.append(gradient.get("stop_1_color"))
            values.append(gradient.get("stop_2_color"))
    return values


def _payload_uses_office_default_blue(payload, *, fill=False, stroke=False):
    for value in _payload_color_values(payload, include_fill=fill, include_stroke=stroke):
        color = _hex_to_rgb_int(value)
        if color in _OFFICE_DEFAULT_BLUE_COLORS:
            return True
    return False


def _is_decorative_helper_payload(payload, element_id=None):
    identifier = _payload_identifier_text(payload, element_id)
    decorative_name_tokens = (
        "background",
        "bg_",
        "blob",
        "contour",
        "halo",
        "orb",
        "ring",
        "smoke",
        "thermal",
        "wave",
    )
    if any(token in identifier for token in decorative_name_tokens):
        return True

    comp_type = str((payload or {}).get("type") or "").strip().lower()
    if comp_type not in {"shape", ""}:
        return False
    if _payload_text_content(payload):
        return False

    shape_type = _payload_shape_type_text(payload)
    decorative_shape_types = {
        "arc",
        "block_arc",
        "blockarc",
        "circle",
        "cloud",
        "donut",
        "ellipse",
        "oval",
        "sinusoid",
        "ring",
    }
    return shape_type in decorative_shape_types


def _uno_color_is_office_default_blue(value):
    try:
        color = int(value)
    except Exception:
        return False
    return color in _OFFICE_DEFAULT_BLUE_COLORS


def _reassert_implicit_new_shape_defaults(shape, payload, element_id=None):
    decorative_helper = _is_decorative_helper_payload(payload, element_id)
    suppress_default_blue_fill = (
        decorative_helper and _payload_uses_office_default_blue(payload, fill=True, stroke=False)
    )
    suppress_default_blue_stroke = (
        decorative_helper
        and (
            not _payload_has_explicit_visual_stroke_color(payload)
            or _payload_uses_office_default_blue(payload, fill=False, stroke=True)
        )
    )

    if not _payload_has_explicit_visual_fill(payload) or suppress_default_blue_fill:
        try:
            shape.FillStyle = _const("com.sun.star.drawing.FillStyle.NONE", 0)
        except Exception:
            pass
    if not _payload_has_explicit_visual_stroke(payload) or suppress_default_blue_stroke:
        try:
            shape.LineStyle = _const("com.sun.star.drawing.LineStyle.NONE", 0)
        except Exception:
            pass

    # Some decorative shapes become Office accent-blue after applying
    # transparency-only fills or width-only lines. Inspect the actual UNO state
    # after all payload styling and clear only default-blue decorative surfaces.
    if decorative_helper:
        try:
            if _uno_color_is_office_default_blue(_property_value(shape, "FillColor")):
                shape.FillStyle = _const("com.sun.star.drawing.FillStyle.NONE", 0)
        except Exception:
            pass
        try:
            if _uno_color_is_office_default_blue(_property_value(shape, "LineColor")):
                shape.LineStyle = _const("com.sun.star.drawing.LineStyle.NONE", 0)
        except Exception:
            pass


def _apply_payload_to_shape(shape, payload):
    _log_json("V2/apply_payload_to_shape: payload=", payload)
    changed = False
    try:
        if str(payload.get("type") or "").strip().lower() == "shape":
            if hasattr(shape, "CustomShapeGeometry") and _set_custom_shape_type(shape, payload.get("shape_type")):
                changed = True
    except Exception:
        pass
    geometry_changed = _apply_geometry(shape, payload)
    if geometry_changed:
        changed = True
    text_changed = _apply_text(shape, payload)
    if text_changed:
        changed = True
    text_style_changed = _apply_text_style(shape, payload)
    if text_style_changed:
        changed = True
    style_changed = _apply_style(shape, payload)
    if style_changed:
        changed = True

    # image_url support for image components
    if str(payload.get("type") or "").lower() == "image":
        image_style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
        image_url = payload.get("image_url") or payload.get("url")
        if image_url:
            try:
                graphic_url = _resolve_image_graphic_url(image_url, image_style=image_style, shape=shape)
                graphic_loaded = False
                if graphic_url and str(graphic_url).startswith(("file:", "private:")):
                    graphic = _query_graphic_from_url(graphic_url)
                    if graphic is not None and _set_any_uno_property_quiet(shape, ("Graphic",), graphic):
                        graphic_loaded = True
                        changed = True
                        _log(f"V2/apply_payload_to_shape: image Graphic embedded url={graphic_url}")
                if not graphic_loaded and graphic_url:
                    shape.GraphicURL = str(graphic_url)
                    changed = True
                    _log(f"V2/apply_payload_to_shape: image GraphicURL set url={graphic_url}")
            except Exception as exc:
                _log(f"V2/apply_payload_to_shape: image GraphicURL failed: {exc}")

        transparency = image_style.get("transparency")
        if transparency is not None:
            try:
                shape.GraphicTransparency = max(0, min(100, _safe_int(float(transparency) * 100.0, 0)))
                changed = True
            except Exception:
                pass
        try:
            image_corner_radius = image_style.get("corner_radius")
            if image_corner_radius is not None and hasattr(shape, "CornerRadius"):
                shape.CornerRadius = max(0, _pt_to_mm100(image_corner_radius))
                changed = True
            elif image_style.get("is_circle") and hasattr(shape, "CornerRadius"):
                size = shape.getSize()
                radius = min(int(size.Width), int(size.Height)) // 2
                shape.CornerRadius = max(0, radius)
                changed = True
        except Exception:
            pass
    _log(
        "V2/apply_payload_to_shape: "
        f"geometry_changed={geometry_changed} text_changed={text_changed} "
        f"text_style_changed={text_style_changed} style_changed={style_changed} changed={changed}"
    )
    return changed


def _extract_component_geometry(component):
    if not isinstance(component, dict):
        return 0.0, 0.0, 0.0, 0.0
    geometry = component.get("geometry") if isinstance(component.get("geometry"), dict) else {}
    position = component.get("position") if isinstance(component.get("position"), dict) else {}
    size = component.get("size") if isinstance(component.get("size"), dict) else {}
    x = geometry.get("x", position.get("x", 0.0))
    y = geometry.get("y", position.get("y", 0.0))
    width = geometry.get("width", size.get("width", 0.0))
    height = geometry.get("height", size.get("height", 0.0))
    return _safe_float(x), _safe_float(y), _safe_float(width), _safe_float(height)


def _table_dimensions(table_data):
    if not isinstance(table_data, dict):
        return 0, 0
    raw_data = table_data.get("data")
    data = raw_data if isinstance(raw_data, list) else []
    if not data and isinstance(table_data.get("rows"), list):
        data = table_data.get("rows")
    headers = table_data.get("headers") if isinstance(table_data.get("headers"), list) else []
    body_rows = table_data.get("row_count")
    if body_rows is None:
        body_rows = table_data.get("rows")
    if isinstance(body_rows, list):
        body_rows = len(body_rows)
    if body_rows is None:
        body_rows = len(data)

    cols = table_data.get("col_count")
    if cols is None:
        cols = table_data.get("cols")
    if cols is None:
        if headers:
            cols = len(headers)
        else:
            max_cells = 0
            for row_data in data:
                if isinstance(row_data, dict) and isinstance(row_data.get("cells"), list):
                    max_cells = max(max_cells, len(row_data.get("cells")))
                elif isinstance(row_data, list):
                    max_cells = max(max_cells, len(row_data))
            cols = max_cells if max_cells > 0 else 1

    try:
        body_rows = int(body_rows)
    except Exception:
        body_rows = len(data)
    rows = max(0, body_rows) + (1 if headers else 0)
    try:
        cols = int(cols)
    except Exception:
        cols = 1
    return max(1, rows), max(1, cols)


def _table_model_from_shape(shape_or_table):
    if shape_or_table is None:
        return None
    try:
        if hasattr(shape_or_table, "getTable"):
            table = shape_or_table.getTable()
            if table is not None:
                return table
    except Exception:
        pass

    for attr_name in ("Model", "Table", "TableModel"):
        try:
            table = getattr(shape_or_table, attr_name, None)
            if table is not None:
                return table
        except Exception:
            continue
    for prop_name in ("Model", "Table", "TableModel"):
        try:
            table = _property_value(shape_or_table, prop_name)
            if table is not None:
                return table
        except Exception:
            continue
    return shape_or_table


def _resize_uno_table_grid(shape_or_table, target_rows, target_cols):
    table = _table_model_from_shape(shape_or_table)
    if table is None:
        return False

    changed = False
    try:
        rows_obj = table.getRows() if hasattr(table, "getRows") else None
        if rows_obj is not None and hasattr(rows_obj, "getCount"):
            current_rows = int(rows_obj.getCount())
            if current_rows < target_rows and hasattr(rows_obj, "insertByIndex"):
                rows_obj.insertByIndex(current_rows, int(target_rows - current_rows))
                changed = True
            elif current_rows > target_rows and hasattr(rows_obj, "removeByIndex"):
                rows_obj.removeByIndex(int(target_rows), int(current_rows - target_rows))
                changed = True
    except Exception:
        pass

    try:
        cols_obj = table.getColumns() if hasattr(table, "getColumns") else None
        if cols_obj is not None and hasattr(cols_obj, "getCount"):
            current_cols = int(cols_obj.getCount())
            if current_cols < target_cols and hasattr(cols_obj, "insertByIndex"):
                cols_obj.insertByIndex(current_cols, int(target_cols - current_cols))
                changed = True
            elif current_cols > target_cols and hasattr(cols_obj, "removeByIndex"):
                cols_obj.removeByIndex(int(target_cols), int(current_cols - target_cols))
                changed = True

            # Distribute column widths equally after resize.
            # Without this, column 0 takes most space and added columns get minimal width.
            try:
                total_cols = int(cols_obj.getCount())
                if total_cols > 0:
                    shape_width = 0
                    try:
                        shape_width = int(shape_or_table.getSize().Width)
                    except Exception:
                        pass
                    if shape_width > 0:
                        equal_width = max(1, shape_width // total_cols)
                        for ci in range(total_cols):
                            try:
                                col = cols_obj.getByIndex(ci)
                                _set_uno_property(col, "Width", equal_width)
                            except Exception:
                                continue
                        changed = True
            except Exception:
                pass
    except Exception:
        pass

    return changed


def _create_table_shape(document, slide, payload, element_id, ancestors=None):
    table_data = payload.get("table_data") if isinstance(payload.get("table_data"), dict) else {}
    rows, cols = _table_dimensions(table_data)
    shape = None
    last_exc = None
    for service_name in ("com.sun.star.drawing.TableShape", "com.sun.star.presentation.TableShape"):
        try:
            shape = document.createInstance(service_name)
            break
        except Exception as exc:
            last_exc = exc
            continue
    if shape is None:
        _log(f"V2/create_table_shape: table service unavailable, fallback rectangle: {last_exc}")
        shape = document.createInstance("com.sun.star.drawing.RectangleShape")
        _set_shape_identity(shape, element_id, ancestors)
        slide.add(shape)
        _apply_geometry(shape, payload)
    else:
        # Set geometry and add to slide BEFORE initialize —
        # Collabora requires the shape to be on the draw page first.
        _set_shape_identity(shape, element_id, ancestors)
        _apply_geometry(shape, payload)
        slide.add(shape)

        initialized = False
        if hasattr(shape, "initialize"):
            try:
                shape.initialize(rows, cols)
                initialized = True
                _log(f"V2/create_table_shape: initialized rows={rows} cols={cols}")
            except Exception as exc:
                _log(f"V2/create_table_shape: initialize failed rows={rows} cols={cols}: {exc}")
        if not initialized:
            resized = _resize_uno_table_grid(shape, rows, cols)
            _log(f"V2/create_table_shape: initialize={initialized} resize_grid={resized}")

    if not isinstance(table_data, dict):
        return shape

    # Best-effort table population.
    table = _table_model_from_shape(shape)
    if table is None:
        _log("V2/create_table_shape: no table model on shape, returning fallback geometry only")
        return shape
    if not hasattr(table, "getCellByPosition"):
        _log(
            "V2/create_table_shape: table model missing getCellByPosition "
            f"type={type(table)} has_getRows={hasattr(table, 'getRows')} has_getColumns={hasattr(table, 'getColumns')}"
        )
        return shape

    headers = table_data.get("headers") if isinstance(table_data.get("headers"), list) else []
    data_rows = table_data.get("data") if isinstance(table_data.get("data"), list) else []
    if not data_rows and isinstance(table_data.get("rows"), list):
        data_rows = table_data.get("rows")
    header_count = 1 if headers else 0
    body_row_count = max(0, rows - header_count)
    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    header_fill = _hex_to_rgb_int(style.get("header_fill") or "203864")
    header_font_color = _hex_to_rgb_int(style.get("header_font_color") or "FFFFFF")
    body_fill = _hex_to_rgb_int(style.get("body_fill") or "FFFFFF")
    body_font_color = _hex_to_rgb_int(style.get("body_font_color") or "111827")
    banded_rows = bool(style.get("banded_rows", True))
    banded_fill = _hex_to_rgb_int(style.get("banded_fill") or "F3F4F6")
    header_font_size = _safe_float(style.get("header_font_size"), 14.0)
    body_font_size = _safe_float(style.get("body_font_size"), 12.0)
    cell_margin_mm100 = _pt_to_mm100(style.get("cell_margin", 5))
    column_styles = style.get("column_styles") if isinstance(style.get("column_styles"), list) else []
    specified_table_width_mm100 = None
    specified_table_height_mm100 = None

    # Best-effort column widths / row heights parity with PPTX table renderer.
    try:
        col_widths = table_data.get("col_widths") if isinstance(table_data.get("col_widths"), list) else []
        if col_widths:
            raw_shape_size = shape.getSize()
            table_widths_pt = _normalize_table_axis_sizes_to_points(
                col_widths,
                _mm100_to_pt(getattr(raw_shape_size, "Width", 0)),
            )
            specified_table_width_mm100 = sum(
                max(1, _pt_to_mm100(width_pt))
                for width_pt in table_widths_pt[:cols]
            )
            columns_obj = table.getColumns() if hasattr(table, "getColumns") else None
            if columns_obj is not None:
                for col_idx, width_pt in enumerate(table_widths_pt):
                    if col_idx >= cols:
                        break
                    try:
                        column = columns_obj.getByIndex(col_idx)
                        _set_uno_property(column, "Width", max(1, _pt_to_mm100(width_pt)))
                    except Exception:
                        continue
    except Exception:
        pass

    try:
        row_heights = table_data.get("row_heights") if isinstance(table_data.get("row_heights"), list) else []
        if row_heights:
            raw_shape_size = shape.getSize()
            table_heights_pt = _fit_table_axis_sizes_to_points(
                row_heights,
                _mm100_to_pt(getattr(raw_shape_size, "Height", 0)),
                rows,
                body_count=body_row_count,
                header_count=header_count,
            )
            specified_table_height_mm100 = sum(
                max(1, _pt_to_mm100(height_pt))
                for height_pt in table_heights_pt[:rows]
            )
            rows_obj = table.getRows() if hasattr(table, "getRows") else None
            if rows_obj is not None:
                for row_idx, height_pt in enumerate(table_heights_pt):
                    if row_idx >= rows:
                        break
                    try:
                        row = rows_obj.getByIndex(row_idx)
                        _set_uno_property(row, "Height", max(1, _pt_to_mm100(height_pt)))
                    except Exception:
                        continue
    except Exception:
        pass

    # Sync table shape extents to effective column/row dimensions, matching python-pptx behavior.
    try:
        effective_width = 0
        columns_obj = table.getColumns() if hasattr(table, "getColumns") else None
        if columns_obj is not None and hasattr(columns_obj, "getCount"):
            count = min(cols, int(columns_obj.getCount()))
            for col_idx in range(count):
                try:
                    column = columns_obj.getByIndex(col_idx)
                    width = _property_value(column, "Width")
                    if width is not None:
                        effective_width += max(0, int(width))
                except Exception:
                    continue

        effective_height = 0
        rows_obj = table.getRows() if hasattr(table, "getRows") else None
        if rows_obj is not None and hasattr(rows_obj, "getCount"):
            count = min(rows, int(rows_obj.getCount()))
            for row_idx in range(count):
                try:
                    row = rows_obj.getByIndex(row_idx)
                    height = _property_value(row, "Height")
                    if height is not None:
                        effective_height += max(0, int(height))
                except Exception:
                    continue

        if effective_width > 0 or effective_height > 0:
            old_size = shape.getSize()
            target_width = (
                int(specified_table_width_mm100)
                if specified_table_width_mm100 is not None and specified_table_width_mm100 > 0
                else effective_width
            )
            target_height = (
                int(specified_table_height_mm100)
                if specified_table_height_mm100 is not None and specified_table_height_mm100 > 0
                else effective_height
            )
            shape.setSize(
                _create_size(
                    target_width if target_width > 0 else int(old_size.Width),
                    target_height if target_height > 0 else int(old_size.Height),
                )
            )
    except Exception as exc:
        _log(f"V2/create_table_shape: failed to sync table extents: {exc}")

    for col_idx, header in enumerate(headers):
        if col_idx >= cols:
            break
        try:
            cell = table.getCellByPosition(col_idx, 0)
            cell.setString(str(header))
            if header_fill is not None:
                cell.FillColor = header_fill
            for prop in ("TextLeftDistance", "TextRightDistance", "TextUpperDistance", "TextLowerDistance"):
                _set_uno_property(cell, prop, cell_margin_mm100)
            cursor = cell.getText().createTextCursor()
            if header_font_color is not None:
                cursor.CharColor = header_font_color
            cursor.CharHeight = float(header_font_size)
            cursor.CharWeight = _const("com.sun.star.awt.FontWeight.BOLD", 150.0)
            cursor.ParaAdjust = _const("com.sun.star.style.ParagraphAdjust.CENTER", 3)
        except Exception:
            continue

    row_offset = 1 if headers else 0
    for row_idx, row_data in enumerate(data_rows):
        table_row = row_idx + row_offset
        if table_row >= rows:
            break
        if isinstance(row_data, dict):
            cells = row_data.get("cells") if isinstance(row_data.get("cells"), list) else []
        elif isinstance(row_data, list):
            cells = row_data
        else:
            cells = []
        for col_idx, value in enumerate(cells):
            if col_idx >= cols:
                break
            try:
                cell = table.getCellByPosition(col_idx, table_row)
                cell.setString(str(value))
                fill = body_fill
                if banded_rows and row_idx % 2 == 1 and banded_fill is not None:
                    fill = banded_fill
                if fill is not None:
                    cell.FillColor = fill
                for prop in ("TextLeftDistance", "TextRightDistance", "TextUpperDistance", "TextLowerDistance"):
                    _set_uno_property(cell, prop, cell_margin_mm100)
                cursor = cell.getText().createTextCursor()
                if body_font_color is not None:
                    cursor.CharColor = body_font_color
                cursor.CharHeight = float(body_font_size)
                alignment = _const("com.sun.star.style.ParagraphAdjust.LEFT", 0)
                col_style = _resolve_table_column_style(column_styles, col_idx)
                if col_style:
                    align_str = str(col_style.get("alignment") or "left").strip().lower()
                    if align_str in {"center", "middle"}:
                        alignment = _const("com.sun.star.style.ParagraphAdjust.CENTER", 3)
                    elif align_str in {"right", "end"}:
                        alignment = _const("com.sun.star.style.ParagraphAdjust.RIGHT", 1)
                    else:
                        alignment = _const("com.sun.star.style.ParagraphAdjust.LEFT", 0)
                cursor.ParaAdjust = alignment
            except Exception:
                continue

    # Apply table-level effects (shadow) on the containing shape.
    try:
        effects = style.get("effects") if isinstance(style.get("effects"), dict) else {}
        if effects.get("shadow"):
            _apply_style(shape, {"style": {"effects": {"shadow": effects.get("shadow")}}})
    except Exception:
        pass
    return shape


def _get_chart_document(shape):
    chart_doc = None
    for getter in ("getEmbeddedObject",):
        try:
            chart_doc = getattr(shape, getter)()
        except Exception:
            chart_doc = None
        if chart_doc is not None:
            return chart_doc
    for attr in ("Model", "ChartModel"):
        try:
            chart_doc = getattr(shape, attr, None)
        except Exception:
            chart_doc = None
        if chart_doc is not None:
            return chart_doc
    return None


def _initialize_chart_document(chart_document):
    if chart_document is None:
        return False
    changed = False

    try:
        creator = getattr(chart_document, "createDefaultChart", None)
        if callable(creator):
            creator()
            changed = True
    except Exception:
        pass

    try:
        has_internal = getattr(chart_document, "hasInternalDataProvider", None)
        already_internal = bool(has_internal()) if callable(has_internal) else False
    except Exception:
        already_internal = False

    if not already_internal:
        try:
            creator = getattr(chart_document, "createInternalDataProvider", None)
            if callable(creator):
                creator(False)
                changed = True
        except Exception:
            pass

    return changed


def _normalize_chart_type(chart_type):
    token = str(chart_type or "").strip().lower().replace("-", "_")
    # PPTX-like names used by backend renderer styles.
    mapping = {
        "column_clustered": "column",
        "column_stacked": "column",
        "column_stacked_100": "column",
        "bar_clustered": "bar",
        "bar_stacked": "bar",
        "bar_stacked_100": "bar",
        "line": "line",
        "line_markers": "line",
        "line_markers_stacked": "line",
        "pie": "pie",
        "doughnut": "doughnut",
        "area": "area",
        "area_stacked": "area",
        "area_stacked_100": "area",
        "xy_scatter": "scatter",
        "xy_scatter_lines": "scatter",
        "xy_scatter_lines_no_markers": "scatter",
        "xy_scatter_smooth": "scatter",
        "xy_scatter_smooth_no_markers": "scatter",
    }
    if token in mapping:
        return mapping[token]
    if token.endswith("_3d"):
        token = token[:-3]
    if token.startswith("column"):
        return "column"
    if token.startswith("bar"):
        return "bar"
    if token.startswith("line"):
        return "line"
    if token.startswith("area"):
        return "area"
    return token or "column"


def _chart_type_to_uno(chart_type):
    normalized = _normalize_chart_type(chart_type)
    mapping = {
        "bar": "com.sun.star.chart2.BarDiagram",
        "column": "com.sun.star.chart2.ColumnDiagram",
        "line": "com.sun.star.chart2.LineDiagram",
        "pie": "com.sun.star.chart2.PieDiagram",
        "doughnut": "com.sun.star.chart2.DoughnutDiagram",
        "area": "com.sun.star.chart2.AreaDiagram",
        "radar": "com.sun.star.chart2.NetDiagram",
        "scatter": "com.sun.star.chart2.XYDiagram",
        "surface": "com.sun.star.chart2.SurfaceDiagram",
    }
    return mapping.get(normalized)


def _apply_chart_type(chart_document, chart_type):
    diagram_service = _chart_type_to_uno(chart_type)
    if not chart_document or not diagram_service:
        return False
    changed = False
    try:
        factory = getattr(chart_document, "createInstance", None)
        if callable(factory):
            diagram = factory(diagram_service)
            if diagram is not None:
                chart_document.setDiagram(diagram)
                return True
    except Exception:
        pass
    try:
        fallback_factory = getattr(chart_document, "getFactory", None)
        if callable(fallback_factory):
            service = fallback_factory()
            if service is not None:
                diagram = service.createInstance(diagram_service)
                if diagram is not None:
                    chart_document.setDiagram(diagram)
                    changed = True
    except Exception:
        pass
    return changed


def _extract_chart_number_format(chart_data):
    if not isinstance(chart_data, dict):
        return None
    series = chart_data.get("series") if isinstance(chart_data.get("series"), list) else []
    for serie in series:
        if not isinstance(serie, dict):
            continue
        number_format = serie.get("number_format")
        if isinstance(number_format, str) and number_format.strip():
            return number_format.strip()
    return None


def _default_uno_locale():
    if uno is None:
        return None
    try:
        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        locale.Language = "en"
        locale.Country = "US"
        locale.Variant = ""
        return locale
    except Exception:
        return None


def _chart_number_formats(chart_document):
    if chart_document is None:
        return None

    getter = getattr(chart_document, "getNumberFormats", None)
    if callable(getter):
        try:
            formats = getter()
            if formats is not None:
                return formats
        except Exception:
            pass

    for prop_name in ("NumberFormats", "NumberFormatsSupplier"):
        try:
            candidate = _property_value(chart_document, prop_name)
        except Exception:
            candidate = None
        if candidate is None:
            continue
        if hasattr(candidate, "queryKey") and hasattr(candidate, "addNew"):
            return candidate
        supplier_getter = getattr(candidate, "getNumberFormats", None)
        if callable(supplier_getter):
            try:
                formats = supplier_getter()
                if formats is not None:
                    return formats
            except Exception:
                pass
    return None


def _resolve_number_format_key(number_formats, format_code):
    if number_formats is None:
        return None
    text = str(format_code or "").strip()
    if not text:
        return None

    locale = _default_uno_locale()
    if locale is None:
        return None

    key = None
    try:
        key = number_formats.queryKey(text, locale, True)
    except Exception:
        key = None

    try:
        if key is not None and int(key) >= 0:
            return int(key)
    except Exception:
        pass

    try:
        key = number_formats.addNew(text, locale)
    except Exception:
        key = None

    try:
        key = int(key)
        return key if key >= 0 else None
    except Exception:
        return None


def _set_number_format_on_target(target, format_key):
    if target is None or format_key is None:
        return False
    changed = False
    if _set_any_uno_property_quiet(target, ("NumberFormat", "NumberFormatValue"), int(format_key)):
        changed = True
    # Prefer explicit unlinking from source-linked defaults when available.
    if _set_any_uno_property_quiet(
        target,
        ("LinkNumberFormatToSource", "NumberFormatIsLinked", "SourceLinked"),
        False,
    ):
        changed = True
    return changed


def _iter_chart_number_format_targets(chart_document):
    targets = []
    if chart_document is None:
        return targets

    targets.append(chart_document)

    diagram = None
    try:
        diagram = chart_document.getDiagram() if hasattr(chart_document, "getDiagram") else None
    except Exception:
        diagram = None
    if diagram is None:
        return targets

    targets.append(diagram)

    for axis_getter in ("getXAxis", "getYAxis", "getZAxis"):
        getter = getattr(diagram, axis_getter, None)
        if not callable(getter):
            continue
        try:
            axis = getter()
            if axis is not None:
                targets.append(axis)
        except Exception:
            continue

    coordinate_systems = None
    try:
        coordinate_systems = diagram.getCoordinateSystems() if hasattr(diagram, "getCoordinateSystems") else None
    except Exception:
        coordinate_systems = None
    if not coordinate_systems:
        return targets

    for coordinate_system in coordinate_systems:
        if coordinate_system is None:
            continue
        for dimension in (0, 1, 2):
            try:
                axis = coordinate_system.getAxisByDimension(dimension, 0)
                if axis is not None:
                    targets.append(axis)
            except Exception:
                continue
        try:
            chart_types = coordinate_system.getChartTypes()
        except Exception:
            chart_types = None
        if not chart_types:
            continue
        for chart_type_obj in chart_types:
            if chart_type_obj is None:
                continue
            targets.append(chart_type_obj)
            try:
                data_series = chart_type_obj.getDataSeries()
            except Exception:
                data_series = None
            if not data_series:
                continue
            for series in data_series:
                if series is not None:
                    targets.append(series)

    return targets


def _apply_chart_number_format(chart_document, chart_data):
    number_format_code = _extract_chart_number_format(chart_data)
    if not number_format_code:
        return False

    number_formats = _chart_number_formats(chart_document)
    if number_formats is None:
        _log("V2/apply_chart_number_format: no NumberFormats supplier available")
        return False

    format_key = _resolve_number_format_key(number_formats, number_format_code)
    if format_key is None:
        _log(f"V2/apply_chart_number_format: unable to resolve format_code={number_format_code}")
        return False

    changed = False
    for target in _iter_chart_number_format_targets(chart_document):
        if _set_number_format_on_target(target, format_key):
            changed = True

    if changed:
        _log(
            "V2/apply_chart_number_format: "
            f"applied format_code={number_format_code} format_key={format_key}"
        )
    else:
        _log(
            "V2/apply_chart_number_format: "
            f"no writable number format target for format_code={number_format_code}"
        )
    return changed


def _set_chart_receiver_arguments(chart_document, has_series_labels, has_categories):
    if chart_document is None or uno is None:
        return False

    setter = getattr(chart_document, "setArguments", None)
    if not callable(setter):
        return False

    try:
        setter(
            (
                _uno_property("CellRangeRepresentation", "all"),
                _uno_property(
                    "DataRowSource",
                    _const("com.sun.star.chart.ChartDataRowSource.COLUMNS", 1),
                ),
                _uno_property("FirstCellAsLabel", bool(has_series_labels)),
                _uno_property("HasCategories", bool(has_categories)),
            )
        )
        return True
    except Exception as exc:
        _log(f"V2/set_chart_receiver_arguments: failed: {exc}")
        return False


def _set_chart_data(chart_document, chart_data):
    if not chart_document or not isinstance(chart_data, dict):
        _log("V2/set_chart_data: skipped invalid chart_document or chart_data")
        return False
    categories = chart_data.get("categories") if isinstance(chart_data.get("categories"), list) else []
    series = chart_data.get("series") if isinstance(chart_data.get("series"), list) else []

    series_matrix = []
    full_table_matrix = []
    series_labels = []
    max_length = len(categories)
    for serie in series:
        if not isinstance(serie, dict):
            continue
        values = serie.get("values") if isinstance(serie.get("values"), list) else []
        normalized_values = []
        for value in values:
            try:
                normalized_values.append(float(value))
            except Exception:
                normalized_values.append(0.0)
        max_length = max(max_length, len(normalized_values))
        series_labels.append(str(serie.get("name") or f"Series {len(series_labels) + 1}"))
        series_matrix.append(tuple(normalized_values))

    padded_series = []
    if series_matrix:
        for row in series_matrix:
            if len(row) < max_length:
                row = row + (0.0,) * (max_length - len(row))
            padded_series.append(row)

    padded_rows = []
    for row_idx in range(max_length):
        row_values = []
        for series_values in padded_series:
            row_values.append(series_values[row_idx] if row_idx < len(series_values) else 0.0)
        padded_rows.append(tuple(row_values))

    padded_tuple = tuple(padded_rows)
    row_labels_tuple = tuple(
        str(categories[row_idx]) if row_idx < len(categories) else f"Category {row_idx + 1}"
        for row_idx in range(max_length)
    )
    column_labels_tuple = tuple(series_labels)

    if series_labels or row_labels_tuple or padded_rows:
        full_table_matrix.append([""] + list(column_labels_tuple))
        for row_idx, row_values in enumerate(padded_rows):
            row = [row_labels_tuple[row_idx]]
            row.extend(row_values)
            full_table_matrix.append(row)

    changed = False

    data = None
    getter = getattr(chart_document, "getData", None)
    if callable(getter):
        try:
            data = getter()
        except Exception:
            data = None

    if data is not None:
        if full_table_matrix and hasattr(data, "setData"):
            try:
                data.setData(full_table_matrix)
                changed = True
            except Exception as exc:
                _log(f"V2/set_chart_data: full table setData failed: {exc}")
        if padded_tuple:
            try:
                if not changed and hasattr(data, "setDataArray"):
                    data.setDataArray(padded_tuple)
                    changed = True
                elif not changed and hasattr(data, "setData"):
                    data.setData(padded_tuple)
                    changed = True
            except Exception as exc:
                _log(f"V2/set_chart_data: direct data matrix write failed: {exc}")
        if row_labels_tuple:
            for method_name in ("setRowDescriptions", "setAnyRowDescriptions"):
                if not hasattr(data, method_name):
                    continue
                try:
                    getattr(data, method_name)(row_labels_tuple)
                    changed = True
                    break
                except Exception:
                    continue
        if column_labels_tuple and max_length >= len(row_labels_tuple):
            for method_name in ("setColumnDescriptions", "setAnyColumnDescriptions"):
                if not hasattr(data, method_name):
                    continue
                try:
                    getattr(data, method_name)(column_labels_tuple)
                    changed = True
                    break
                except Exception:
                    continue

    provider = None
    provider_getter = getattr(chart_document, "getDataProvider", None)
    if callable(provider_getter):
        try:
            provider = provider_getter()
        except Exception:
            provider = None

    if provider is not None:
        if padded_tuple:
            for method_name in ("setData", "setDataArray"):
                if not hasattr(provider, method_name):
                    continue
                try:
                    getattr(provider, method_name)(padded_tuple)
                    changed = True
                    break
                except Exception:
                    continue
        if row_labels_tuple:
            for method_name in ("setRowDescriptions", "setAnyRowDescriptions"):
                if not hasattr(provider, method_name):
                    continue
                try:
                    getattr(provider, method_name)(row_labels_tuple)
                    changed = True
                    break
                except Exception:
                    continue
        if column_labels_tuple and max_length >= len(row_labels_tuple):
            for method_name in ("setColumnDescriptions", "setAnyColumnDescriptions"):
                if not hasattr(provider, method_name):
                    continue
                try:
                    getattr(provider, method_name)(column_labels_tuple)
                    changed = True
                    break
                except Exception:
                    continue
        try:
            receiver_attach = getattr(chart_document, "attachDataProvider", None)
            if callable(receiver_attach):
                receiver_attach(provider)
                changed = True
        except Exception:
            pass
        if _set_chart_receiver_arguments(
            chart_document,
            has_series_labels=bool(column_labels_tuple),
            has_categories=bool(row_labels_tuple),
        ):
            changed = True

    if changed:
        _log(
            "V2/set_chart_data: "
            f"series_count={len(series_labels)} category_count={len(categories)} max_len={max_length}"
        )
    else:
        _log(
            "V2/set_chart_data: no writable chart data interface found "
            f"has_data={data is not None} has_provider={provider is not None}"
        )
    return changed


def _style_chart_document(chart_document, payload):
    if not chart_document or not isinstance(payload, dict):
        return False
    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    changed = False

    chart_type = style.get("chart_type")
    if chart_type:
        if _apply_chart_type(chart_document, chart_type):
            changed = True

    has_title = style.get("has_title")
    if has_title is False:
        for prop in ("HasMainTitle", "HasTitle"):
            try:
                if _set_uno_property(chart_document, prop, False):
                    changed = True
            except Exception:
                continue

    title_text = payload.get("title") or style.get("title")
    if title_text and has_title is not False:
        try:
            chart_document.HasMainTitle = True  # type: ignore[attr-defined]
            chart_document.Title.String = str(title_text)  # type: ignore[attr-defined]
            changed = True
        except Exception:
            try:
                title = chart_document.getTitle()
                if title is not None:
                    title.String = str(title_text)
                    changed = True
            except Exception:
                pass

    legend_pref = style.get("has_legend")
    if legend_pref is not None:
        try:
            chart_document.HasMainLegend = bool(legend_pref)  # type: ignore[attr-defined]
            changed = True
        except Exception:
            try:
                if _set_uno_property(chart_document, "HasLegend", bool(legend_pref)):
                    changed = True
            except Exception:
                pass
    legend_position = str(style.get("legend_position") or "").strip().lower()
    if legend_position:
        legend_pos_value = {
            "right": _const("com.sun.star.chart.ChartLegendPosition.RIGHT", 3),
            "left": _const("com.sun.star.chart.ChartLegendPosition.LEFT", 2),
            "top": _const("com.sun.star.chart.ChartLegendPosition.TOP", 0),
            "bottom": _const("com.sun.star.chart.ChartLegendPosition.BOTTOM", 1),
            "corner": _const("com.sun.star.chart.ChartLegendPosition.RIGHT", 3),
        }.get(legend_position)
        if legend_pos_value is not None:
            try:
                if _set_uno_property(chart_document, "LegendPosition", legend_pos_value):
                    changed = True
            except Exception:
                pass
            try:
                legend = chart_document.getLegend() if hasattr(chart_document, "getLegend") else None
                if legend is not None:
                    for prop_name in ("AnchorPosition", "Alignment", "LegendPosition"):
                        if _set_uno_property(legend, prop_name, legend_pos_value):
                            changed = True
                            break
            except Exception:
                pass

    labels_pref = style.get("has_data_labels")
    if labels_pref:
        try:
            diagram = chart_document.getDiagram()
            if diagram:
                coordinate_systems = diagram.getCoordinateSystems()
                if coordinate_systems:
                    chart_types = coordinate_systems[0].getChartTypes()
                    if chart_types:
                        chart_type_obj = chart_types[0]
                        chart_type_obj.DataPointProperties = {"DisplayLabel": True}  # type: ignore[attr-defined]
                        changed = True
        except Exception:
            pass
    return changed


def _apply_chart_payload(shape, payload):
    if not isinstance(payload, dict):
        return False
    changed = False

    chart_doc = _get_chart_document(shape)
    if chart_doc is None:
        _log("V2/apply_chart_payload: no chart document found")
        return False
    _log(
        "V2/apply_chart_payload: chart_document_ready "
        f"has_getData={hasattr(chart_doc, 'getData')} has_setDiagram={hasattr(chart_doc, 'setDiagram')} "
        f"class={getattr(chart_doc, 'ImplementationName', None) or type(chart_doc)}"
    )

    if _initialize_chart_document(chart_doc):
        changed = True

    style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
    chart_type = style.get("chart_type")
    if chart_type and _apply_chart_type(chart_doc, chart_type):
        changed = True

    chart_data = payload.get("chart_data") if isinstance(payload.get("chart_data"), dict) else {}
    if chart_data and _set_chart_data(chart_doc, chart_data):
        changed = True
    if chart_data and _apply_chart_number_format(chart_doc, chart_data):
        changed = True

    style_payload = payload
    if chart_type:
        style_without_type = dict(style)
        style_without_type.pop("chart_type", None)
        style_payload = dict(payload)
        style_payload["style"] = style_without_type

    if _style_chart_document(chart_doc, style_payload):
        changed = True
    return changed


def _create_chart_shape(document, slide, payload, element_id, ancestors=None):
    shape = None
    for service in (
        "com.sun.star.drawing.ChartShape",
        "com.sun.star.drawing.OLE2Shape",
    ):
        try:
            shape = document.createInstance(service)
            break
        except Exception:
            continue
    if shape is None:
        shape = document.createInstance("com.sun.star.drawing.RectangleShape")
    try:
        if hasattr(shape, "CLSID"):
            shape.CLSID = "12dcae26-281f-416f-a234-c3086127382e"
    except Exception:
        pass
    _set_shape_identity(shape, element_id, ancestors)
    slide.add(shape)
    _apply_geometry(shape, payload)
    _apply_style(shape, payload)
    _apply_chart_payload(shape, payload)
    return shape


def _compute_connector_endpoints(source_ref, target_ref):
    if not isinstance(source_ref, dict) or not isinstance(target_ref, dict):
        return None

    try:
        ax = float(source_ref["x"])
        ay = float(source_ref["y"])
        aw = float(source_ref["width"])
        ah = float(source_ref["height"])
        bx = float(target_ref["x"])
        by = float(target_ref["y"])
        bw = float(target_ref["width"])
        bh = float(target_ref["height"])
    except Exception:
        return None

    ax_c = ax + (aw / 2.0)
    ay_c = ay + (ah / 2.0)
    bx_c = bx + (bw / 2.0)
    by_c = by + (bh / 2.0)
    dx = bx_c - ax_c
    dy = by_c - ay_c

    start_x, start_y = ax_c, ay_c
    end_x, end_y = bx_c, by_c
    if abs(dx) > abs(dy):
        if dx > 0:
            start_x = ax + aw
            start_y = ay_c
            end_x = bx
            end_y = by_c
        else:
            start_x = ax
            start_y = ay_c
            end_x = bx + bw
            end_y = by_c
    else:
        if dy > 0:
            start_x = ax_c
            start_y = ay + ah
            end_x = bx_c
            end_y = by
        else:
            start_x = ax_c
            start_y = ay
            end_x = bx_c
            end_y = by + bh

    return {
        "start_x": start_x,
        "start_y": start_y,
        "end_x": end_x,
        "end_y": end_y,
        "dx": dx,
        "dy": dy,
    }


def _connector_connection_names(dx, dy):
    try:
        dx = float(dx)
    except Exception:
        dx = 0.0
    try:
        dy = float(dy)
    except Exception:
        dy = 0.0

    if abs(dx) > abs(dy):
        if dx >= 0.0:
            return ("RIGHT", "LEFT")
        return ("LEFT", "RIGHT")

    if dy >= 0.0:
        return ("BOTTOM", "TOP")
    return ("TOP", "BOTTOM")


def _compute_connector_route_points(endpoints):
    if not isinstance(endpoints, dict):
        return []

    try:
        start_x = float(endpoints["start_x"])
        start_y = float(endpoints["start_y"])
        end_x = float(endpoints["end_x"])
        end_y = float(endpoints["end_y"])
        dx = float(endpoints["dx"])
        dy = float(endpoints["dy"])
    except Exception:
        return []

    if abs(start_x - end_x) < 0.5 or abs(start_y - end_y) < 0.5:
        return [
            (start_x, start_y),
            (end_x, end_y),
        ]

    if abs(dx) > abs(dy):
        mid_x = (start_x + end_x) / 2.0
        raw_points = [
            (start_x, start_y),
            (mid_x, start_y),
            (mid_x, end_y),
            (end_x, end_y),
        ]
    else:
        mid_y = (start_y + end_y) / 2.0
        raw_points = [
            (start_x, start_y),
            (start_x, mid_y),
            (end_x, mid_y),
            (end_x, end_y),
        ]

    points = []
    for point in raw_points:
        if not points:
            points.append(point)
            continue
        prev_x, prev_y = points[-1]
        if abs(prev_x - point[0]) < 0.5 and abs(prev_y - point[1]) < 0.5:
            continue
        points.append(point)
    return points


def _set_polyline_points(shape, route_points):
    if shape is None or not route_points:
        return

    left = min(point[0] for point in route_points)
    top = min(point[1] for point in route_points)
    width = max(max(point[0] for point in route_points) - left, 1.0)
    height = max(max(point[1] for point in route_points) - top, 1.0)

    shape.setPosition(_create_point(_pt_to_mm100(left), _pt_to_mm100(top)))
    shape.setSize(_create_size(_pt_to_mm100(width), _pt_to_mm100(height)))

    local_points = []
    for x, y in route_points:
        local_points.append(
            _create_point(
                _pt_to_mm100(x - left),
                _pt_to_mm100(y - top),
            )
        )

    try:
        shape.PolyPolygon = (tuple(local_points),)
    except Exception:
        _set_any_uno_property_quiet(shape, ("PolyPolygon", "Geometry"), (tuple(local_points),))


def _connector_type_from_component(connector_component):
    style = connector_component.get("style") if isinstance(connector_component.get("style"), dict) else {}
    raw = (
        connector_component.get("connector_type")
        or connector_component.get("route")
        or connector_component.get("kind")
        or style.get("connector_type")
        or style.get("route")
        or style.get("kind")
        or "standard"
    )
    token = str(raw or "").strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "straight": "line",
        "direct": "line",
        "linear": "line",
        "elbow": "standard",
        "orthogonal": "standard",
        "bent": "standard",
        "polyline": "lines",
        "multi_line": "lines",
        "multi_lines": "lines",
        "curved": "curve",
    }
    return aliases.get(token, token or "standard")


def _connector_edge_kind(connector_type):
    token = _connector_type_from_component({"connector_type": connector_type})
    mapping = {
        "standard": "STANDARD",
        "curve": "CURVE",
        "line": "LINE",
        "lines": "LINES",
    }
    enum_name = mapping.get(token, "STANDARD")
    return _const(f"com.sun.star.drawing.ConnectorType.{enum_name}", enum_name)


def _create_connector_shape(document, connector_type):
    try:
        connector = document.createInstance("com.sun.star.drawing.ConnectorShape")
    except Exception:
        return None
    try:
        connector.EdgeKind = _connector_edge_kind(connector_type)
    except Exception:
        _set_any_uno_property_quiet(connector, ("EdgeKind",), _connector_edge_kind(connector_type))
    return connector


def _position_connector_shape(connector, start_x, start_y, end_x, end_y, route_points=None):
    left = min(start_x, end_x)
    top = min(start_y, end_y)
    width = max(abs(end_x - start_x), 1.0)
    height = max(abs(end_y - start_y), 1.0)

    if route_points:
        left = min(point[0] for point in route_points)
        top = min(point[1] for point in route_points)
        width = max(max(point[0] for point in route_points) - left, 1.0)
        height = max(max(point[1] for point in route_points) - top, 1.0)

    try:
        connector.setPosition(_create_point(_pt_to_mm100(left), _pt_to_mm100(top)))
        connector.setSize(_create_size(_pt_to_mm100(width), _pt_to_mm100(height)))
    except Exception:
        pass

    start_point = _create_point(_pt_to_mm100(start_x), _pt_to_mm100(start_y))
    end_point = _create_point(_pt_to_mm100(end_x), _pt_to_mm100(end_y))
    _set_any_uno_property_quiet(connector, ("StartPosition", "PositionStart"), start_point)
    _set_any_uno_property_quiet(connector, ("EndPosition", "PositionEnd"), end_point)


def _animation_node_type_for_trigger(trigger):
    token = str(trigger or "").strip().lower().replace("-", "_").replace(" ", "_")
    if token in {"click", "on_click", "onclick"}:
        return _const("com.sun.star.presentation.EffectNodeType.ON_CLICK", 1)
    if token in {"with_previous", "with", "same_time"}:
        return _const("com.sun.star.presentation.EffectNodeType.WITH_PREVIOUS", 2)
    return _const("com.sun.star.presentation.EffectNodeType.AFTER_PREVIOUS", 3)


def _set_animation_user_data(node, trigger):
    if uno is None or node is None:
        return
    try:
        named_value = uno.createUnoStruct("com.sun.star.beans.NamedValue")
        named_value.Name = "node-type"
        named_value.Value = _animation_node_type_for_trigger(trigger)
        node.UserData = (named_value,)
    except Exception:
        pass


def _append_animation_child(parent, child) -> bool:
    if parent is None or child is None:
        return False
    try:
        parent.appendChild(child)
        return True
    except Exception:
        return False


def _create_animation_service(document, service_name):
    try:
        return document.createInstance(service_name)
    except Exception:
        return None


def _apply_component_animation(document, slide, shape, animation):
    if uno is None or document is None or slide is None or shape is None or not isinstance(animation, dict):
        return False
    try:
        root = slide.getAnimationNode() if hasattr(slide, "getAnimationNode") else None
    except Exception:
        root = None
    if root is None:
        return False

    effect = str(animation.get("effect") or "fade").strip().lower().replace("-", "_").replace(" ", "_")
    duration = max(0.1, min(5.0, _safe_float(animation.get("duration"), 0.5)))
    delay = max(0.0, min(10.0, _safe_float(animation.get("delay"), 0.0)))
    trigger = animation.get("trigger") or animation.get("start") or "after_previous"

    group = _create_animation_service(document, "com.sun.star.animations.ParallelTimeContainer")
    if group is None:
        return False
    _set_any_uno_property_quiet(group, ("Begin",), delay)
    _set_any_uno_property_quiet(group, ("Duration",), duration)
    _set_any_uno_property_quiet(group, ("Fill",), _const("com.sun.star.animations.AnimationFill.HOLD", 3))
    _set_animation_user_data(group, trigger)

    made_child = False
    if effect in {"fade", "fade_in", "float_in", "fly_in", "zoom", "wipe"}:
        fade = _create_animation_service(document, "com.sun.star.animations.Animate")
        if fade is not None:
            _set_any_uno_property_quiet(fade, ("Target",), shape)
            _set_any_uno_property_quiet(fade, ("AttributeName",), "opacity")
            _set_any_uno_property_quiet(fade, ("Values",), (0.0, 1.0))
            _set_any_uno_property_quiet(fade, ("KeyTimes",), (0.0, 1.0))
            _set_any_uno_property_quiet(fade, ("Duration",), duration)
            _set_any_uno_property_quiet(fade, ("Fill",), _const("com.sun.star.animations.AnimationFill.HOLD", 3))
            made_child = _append_animation_child(group, fade) or made_child

    appear = _create_animation_service(document, "com.sun.star.animations.AnimateSet")
    if appear is not None:
        _set_any_uno_property_quiet(appear, ("Target",), shape)
        _set_any_uno_property_quiet(appear, ("AttributeName",), "Visibility")
        _set_any_uno_property_quiet(appear, ("To",), True)
        _set_any_uno_property_quiet(appear, ("Duration",), duration)
        _set_any_uno_property_quiet(appear, ("Fill",), _const("com.sun.star.animations.AnimationFill.HOLD", 3))
        made_child = _append_animation_child(group, appear) or made_child

    if not made_child:
        return False
    return _append_animation_child(root, group)


def _render_connector(document, slide, connector_component, id_map):
    source_id = connector_component.get("source_id")
    target_id = connector_component.get("target_id")
    if not source_id or not target_id:
        return None
    source_ref = id_map.get(str(source_id)) or id_map.get(_canonical_id_from_target(source_id))
    target_ref = id_map.get(str(target_id)) or id_map.get(_canonical_id_from_target(target_id))
    if not source_ref or not target_ref:
        return None

    endpoints = _compute_connector_endpoints(source_ref, target_ref)
    if not endpoints:
        return None

    start_x = float(endpoints["start_x"])
    start_y = float(endpoints["start_y"])
    end_x = float(endpoints["end_x"])
    end_y = float(endpoints["end_y"])
    dx = float(endpoints["dx"])
    dy = float(endpoints["dy"])
    route_points = _compute_connector_route_points(endpoints)
    if not route_points:
        return None

    connector_type = _connector_type_from_component(connector_component)
    use_uno_connector = connector_type in {"standard", "curve", "line", "lines"}
    is_bent = len(route_points) > 2

    try:
        if use_uno_connector:
            connector = _create_connector_shape(document, connector_type)
            if connector is None:
                raise RuntimeError("ConnectorShape unavailable")
        elif is_bent:
            connector = document.createInstance("com.sun.star.drawing.PolyLineShape")
        else:
            connector = document.createInstance("com.sun.star.drawing.LineShape")
    except Exception:
        try:
            connector = document.createInstance("com.sun.star.drawing.LineShape")
        except Exception:
            return None

    connector_id = connector_component.get("id")
    _set_shape_identity(connector, connector_id, connector_component.get("_ancestors"))
    slide.add(connector)

    try:
        if use_uno_connector:
            _position_connector_shape(connector, start_x, start_y, end_x, end_y, route_points)
        elif is_bent:
            _set_polyline_points(connector, route_points)
        else:
            left = min(start_x, end_x)
            top = min(start_y, end_y)
            width = max(abs(end_x - start_x), 1.0)
            height = max(abs(end_y - start_y), 1.0)
            connector.setPosition(_create_point(_pt_to_mm100(left), _pt_to_mm100(top)))
            connector.setSize(_create_size(_pt_to_mm100(width), _pt_to_mm100(height)))

            start_point = _create_point(_pt_to_mm100(start_x), _pt_to_mm100(start_y))
            end_point = _create_point(_pt_to_mm100(end_x), _pt_to_mm100(end_y))
            _set_any_uno_property_quiet(connector, ("StartPosition", "PositionStart"), start_point)
            _set_any_uno_property_quiet(connector, ("EndPosition", "PositionEnd"), end_point)
    except Exception:
        pass

    payload = {
        "style": connector_component.get("style") if isinstance(connector_component.get("style"), dict) else {},
    }
    _apply_style(connector, payload)
    _apply_component_animation(document, slide, connector, connector_component.get("animation"))

    # Optional connector label at midpoint (same behavior as PPTX renderer).
    label = connector_component.get("label")
    if label is None and isinstance(payload["style"], dict):
        label = payload["style"].get("label")
    if label:
        try:
            mid_x = (start_x + end_x) / 2.0
            mid_y = (start_y + end_y) / 2.0
            if abs(dx) > abs(dy):
                mid_y -= 12.0
            else:
                mid_x += 12.0

            label_text = str(label)
            label_font_size = _safe_float(payload["style"].get("label_font_size"), 9.0)
            avg_char_w = max(4.0, float(label_font_size) * 0.55)
            horizontal_padding = 10.0
            min_label_w = 60.0
            max_label_w = max(min_label_w, _safe_float(payload["style"].get("label_max_width"), 260.0))
            estimated_w = (len(label_text) * avg_char_w) + (horizontal_padding * 2.0)
            label_w = max(min_label_w, min(max_label_w, estimated_w))

            chars_per_line = max(
                1,
                int((label_w - (horizontal_padding * 2.0)) / avg_char_w),
            )
            line_count = max(1, (len(label_text) + chars_per_line - 1) // chars_per_line)
            line_height = float(label_font_size) * 1.35
            label_h = max(20.0, min(96.0, (line_count * line_height) + 8.0))

            label_shape = document.createInstance("com.sun.star.drawing.TextShape")
            label_payload = {
                "type": "textbox",
                "text": label_text,
                "geometry": {
                    "x": mid_x - (label_w / 2.0),
                    "y": mid_y - (label_h / 2.0),
                    "width": label_w,
                    "height": label_h,
                },
                "style": {
                    "font_name": payload["style"].get("font_name", "Arial"),
                    "font_size": label_font_size,
                    "color": payload["style"].get("label_color", payload["style"].get("stroke_color", "333333")),
                    "text_anchor": "middle",
                    "vertical_anchor": "middle",
                    "fill_type": "none",
                    "no_line": True,
                },
            }
            label_id = None
            if connector_id:
                label_id = f"{connector_id}_label"
            _set_shape_identity(label_shape, label_id, connector_component.get("_ancestors"))
            slide.add(label_shape)
            _apply_payload_to_shape(label_shape, label_payload)
        except Exception as exc:
            _log(f"V2/render_connector: label render failed: {exc}")
    return connector


def _normalized_component_type(value):
    token = str(value or "").strip().lower()
    mapping = {
        "v_stack": "vstack",
        "h_stack": "hstack",
        "z_stack": "zstack",
    }
    return mapping.get(token, token)


def _existing_shape_ids(slide):
    used = set()
    if slide is None or not hasattr(slide, "getCount"):
        return used
    try:
        count = int(slide.getCount())
    except Exception:
        return used
    for idx in range(count):
        try:
            shape = slide.getByIndex(idx)
        except Exception:
            continue
        canonical = _shape_canonical_id(shape)
        if canonical:
            used.add(canonical)
            continue
        try:
            raw_name = shape.getName() if hasattr(shape, "getName") else ""
        except Exception:
            raw_name = ""
        normalized = _canonical_id_from_target(raw_name)
        if normalized:
            used.add(normalized)
    return used


def _slide_hint(slide):
    try:
        name = slide.getName() if hasattr(slide, "getName") else ""
    except Exception:
        name = ""
    token = _normalize_shape_name(name)
    if token:
        return token
    return "0"


def _fallback_component_id(slide, component, comp_type, path_tokens):

    if comp_type in {"textbox", "text", "title", "subtitle"}:
        base = "text_box"
    elif comp_type == "image":
        base = "picture"
    elif comp_type == "table":
        base = "table"
    elif comp_type == "chart":
        base = "chart"
    elif comp_type == "connector":
        base = "connector"
    elif comp_type == "container":
        base = "rectangle"
    elif comp_type == "shape":
        raw_shape_type = _normalize_shape_name(component.get("shape_type") or "shape")
        base = raw_shape_type or "shape"
    else:
        base = _normalize_shape_name(comp_type) or "shape"
    source = json.dumps(
        {
            "slide": _slide_hint(slide),
            "path": list(path_tokens),
            "type": comp_type,
            "shape_type": component.get("shape_type"),
        },
        sort_keys=True,
        ensure_ascii=False,
        default=str,
    )
    used = _existing_shape_ids(slide)
    salt = 0
    while True:
        digest = hashlib.sha1(f"{source}|{salt}".encode("utf-8")).hexdigest()[:8]
        candidate = f"{base}_{digest}"
        if candidate not in used:
            return candidate
        salt += 1


def _render_blueprint_component(document, slide, component, offset_x, offset_y, id_map, connectors, ancestors=None, path_tokens=None):
    if not isinstance(component, dict):
        return
    ancestors = list(ancestors or [])
    path_tokens = list(path_tokens or [0])

    comp_type = _normalized_component_type(component.get("type"))
    rel_x, rel_y, width, height = _extract_component_geometry(component)
    abs_x = offset_x + rel_x
    abs_y = offset_y + rel_y
    comp_id = component.get("id")
    comp_key = str(comp_id) if comp_id is not None else None
    canonical_comp_key = _canonical_id_from_target(comp_key)
    effective_element_id = canonical_comp_key or comp_key
    if comp_id:
        ref = {"x": abs_x, "y": abs_y, "width": width, "height": height, "shape": None}
        id_map[comp_key] = ref
        if canonical_comp_key and canonical_comp_key != comp_key:
            id_map[canonical_comp_key] = ref

    if comp_type == "connector":
        connector_payload = dict(component)
        connector_payload["_ancestors"] = ancestors
        connectors.append(connector_payload)
    elif comp_type != "spacer":
        payload = dict(component)
        payload.pop("children", None)
        payload.pop("items", None)
        payload["geometry"] = {
            "x": abs_x,
            "y": abs_y,
            "width": width,
            "height": height,
        }
        payload["position"] = {"x": abs_x, "y": abs_y}
        payload["size"] = {"width": width, "height": height}

        style = payload.get("style") if isinstance(payload.get("style"), dict) else {}
        structural = comp_type in {"container", "vstack", "hstack", "grid", "zstack", "list"}
        should_create_shape = (not structural) or _style_has_visual_surface(style)

        shape = None
        if not effective_element_id and should_create_shape:
            effective_element_id = _fallback_component_id(slide, component, comp_type, path_tokens)

        if effective_element_id:
            ref = {"x": abs_x, "y": abs_y, "width": width, "height": height, "shape": None}
            id_map[effective_element_id] = ref

        if should_create_shape:
            if comp_type == "table":
                shape = _create_table_shape(document, slide, payload, effective_element_id, ancestors=ancestors)
            elif comp_type == "chart":
                shape = _create_chart_shape(document, slide, payload, effective_element_id, ancestors=ancestors)
            else:
                shape = _create_shape_for_payload(document, payload)
                _set_shape_identity(shape, effective_element_id, ancestors)
                slide.add(shape)
                # Some Collabora/LibreOffice builds re-apply theme defaults when
                # a shape is inserted into a slide. Reassert no-fill/no-line before
                # applying payload style so reloads do not expose accent-colored
                # structural containers or unstyled helper shapes.
                _initialize_new_shape_defaults(shape)
                _apply_payload_to_shape(shape, payload)
                _reassert_implicit_new_shape_defaults(shape, payload, effective_element_id)
            _apply_component_animation(document, slide, shape, payload.get("animation"))

        if comp_key and comp_key in id_map:
            id_map[comp_key]["shape"] = shape
        if canonical_comp_key and canonical_comp_key in id_map:
            id_map[canonical_comp_key]["shape"] = shape
        if effective_element_id and effective_element_id in id_map:
            id_map[effective_element_id]["shape"] = shape

    children = []
    raw_children = component.get("children") if isinstance(component.get("children"), list) else []
    if raw_children:
        if comp_type in {"vstack", "hstack", "grid", "zstack"}:
            type_priority = {
                "shape": 0,
                "image": 1,
                "chart": 2,
                "table": 2,
                "textbox": 3,
                "vstack": 1,
                "hstack": 1,
                "grid": 1,
                "zstack": 1,
            }
            sortable = []
            for child_idx, child in enumerate(raw_children):
                if not isinstance(child, dict):
                    continue
                child_type = _normalized_component_type(child.get("type"))
                z_idx = _safe_int(child.get("z_index"), 0)
                priority = type_priority.get(child_type, 1)
                sortable.append((z_idx, priority, child_idx, child))
            sortable.sort(key=lambda item: (item[0], item[1], item[2]))
            children.extend(item[3] for item in sortable)
        else:
            children.extend(child for child in raw_children if isinstance(child, dict))

    raw_items = component.get("items") if isinstance(component.get("items"), list) else []
    if raw_items:
        if comp_type == "list":
            shared_style = component.get("shared_style") if isinstance(component.get("shared_style"), dict) else {}
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                merged_item = dict(item)
                item_style = merged_item.get("style") if isinstance(merged_item.get("style"), dict) else {}
                merged_style = dict(item_style)
                for key, value in shared_style.items():
                    merged_style[key] = value
                if merged_style:
                    merged_item["style"] = merged_style
                children.append(merged_item)
        else:
            children.extend(item for item in raw_items if isinstance(item, dict))
    child_ancestors = list(ancestors)
    if effective_element_id:
        child_ancestors.append(effective_element_id)

    for child_idx, child in enumerate(children):
        _render_blueprint_component(
            document,
            slide,
            child,
            abs_x,
            abs_y,
            id_map,
            connectors,
            child_ancestors,
            [*path_tokens, child_idx],
        )


def _render_blueprint_to_slide(document, slide, blueprint):
    if not isinstance(blueprint, dict):
        return False
    root = blueprint.get("root")
    if not isinstance(root, dict):
        return False

    id_map = {}
    connectors = []
    _render_blueprint_component(document, slide, root, 0.0, 0.0, id_map, connectors, path_tokens=[0])
    for connector_component in connectors:
        _render_connector(document, slide, connector_component, id_map)
    _set_slide_notes(slide, blueprint.get("speaker_notes"))
    return True


def _ensure_shape_name(shape, element_id):
    if not element_id:
        return
    _set_shape_identity(shape, element_id)


def _set_slide_notes(slide, notes_value):
    if notes_value is None:
        return False
    try:
        notes_page = slide.getNotesPage()
    except Exception:
        return False
    if notes_page is None:
        return False

    # Best-effort: first text-capable shape on notes page.
    try:
        for i in range(notes_page.getCount()):
            shp = notes_page.getByIndex(i)
            try:
                if hasattr(shp, "getText"):
                    shp.getText().setString(str(notes_value))
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _set_uno_property(obj, prop_name: str, value) -> bool:
    if obj is None or not prop_name:
        return False
    # Try XPropertySet path first.
    try:
        if hasattr(obj, "setPropertyValue"):
            obj.setPropertyValue(prop_name, value)
            return True
    except Exception as exc:
        _log(f"V2/set_uno_property: setPropertyValue({prop_name}) failed: {exc}")

    # Fallback to plain attribute assignment.
    try:
        setattr(obj, prop_name, value)
        return True
    except Exception as exc:
        _log(f"V2/set_uno_property: setattr({prop_name}) failed: {exc}")
        return False


def _set_slide_background(slide, background, document=None):
    if not isinstance(background, dict):
        return False
    gradient = background.get("gradient") if isinstance(background.get("gradient"), dict) else {}
    color = background.get("color")
    if color is None:
        color = background.get("background_color")
    if color is None and gradient:
        color = gradient.get("stop_1_color")
    rgb = _hex_to_rgb_int(color)
    if rgb is None and not gradient:
        _log(f"V2/set_slide_background: invalid color={color} gradient={gradient}")
        return False
    _log(f"V2/set_slide_background: target_rgb={rgb} hex={color} gradient={bool(gradient)}")

    solid = _const("com.sun.star.drawing.FillStyle.SOLID", 1)
    gradient_fill = _const("com.sun.star.drawing.FillStyle.GRADIENT", 2)

    def _apply_background_fill(target):
        local_changed = False
        if gradient:
            try:
                stop_1 = _hex_to_rgb_int(gradient.get("stop_1_color")) if gradient.get("stop_1_color") else rgb
                stop_2 = _hex_to_rgb_int(gradient.get("stop_2_color")) if gradient.get("stop_2_color") else stop_1
                grad = uno.createUnoStruct("com.sun.star.awt.Gradient")
                grad.Style = _const("com.sun.star.awt.GradientStyle.LINEAR", 0)
                grad.StartColor = stop_1 if stop_1 is not None else 0xFFFFFF
                grad.EndColor = stop_2 if stop_2 is not None else grad.StartColor
                grad.Angle = _safe_int(float(gradient.get("angle", 90.0)) * 10.0, 900)
                stop_1_pos = max(0.0, min(1.0, _safe_float(gradient.get("stop_1_position"), 0.0)))
                stop_2_pos = max(0.0, min(1.0, _safe_float(gradient.get("stop_2_position"), 1.0)))
                grad.Border = max(0, min(100, _safe_int(stop_1_pos * 100.0, 0)))
                grad.XOffset = max(0, min(100, _safe_int(stop_2_pos * 100.0, 100)))
                grad.YOffset = 50
                grad.StartIntensity = 100
                grad.EndIntensity = 100
                grad.StepCount = 0
                if _set_uno_property(target, "FillStyle", gradient_fill):
                    local_changed = True
                if _set_uno_property(target, "FillGradient", grad):
                    local_changed = True
                if _set_uno_property(target, "FillTransparence", 0):
                    local_changed = True
            except Exception as exc:
                _log(f"V2/set_slide_background: gradient apply failed: {exc}")
        elif rgb is not None:
            if _set_uno_property(target, "FillStyle", solid):
                local_changed = True
            if _set_uno_property(target, "FillColor", rgb):
                _mark_fill_color_as_direct_rgb(target)
                _set_uno_property(target, "FillColor", rgb)
                local_changed = True
            if _set_uno_property(target, "FillTransparence", 0):
                local_changed = True
        return local_changed

    changed = False

    # ── Primary strategy: Background sub-object (the reliable Impress path) ──
    # In Impress, the DrawPage itself does NOT expose FillStyle/FillColor/BackColor.
    # The correct approach is to configure a Background object and assign it.
    bg = None
    try:
        bg = _property_value(slide, "Background")
    except Exception:
        pass
    if bg is None:
        try:
            bg = getattr(slide, "Background", None)
        except Exception as exc:
            _log(f"V2/set_slide_background: reading slide.Background failed: {exc}")

    if bg is None and document is not None:
        try:
            bg = document.createInstance("com.sun.star.drawing.Background")
            _log("V2/set_slide_background: created com.sun.star.drawing.Background instance")
        except Exception as exc:
            _log(f"V2/set_slide_background: create Background service failed: {exc}")
            bg = None

    if bg is not None:
        bg_changed = _apply_background_fill(bg)
        bg_attach_ok = _set_uno_property(slide, "Background", bg)
        if bg_changed or bg_attach_ok:
            _log("V2/set_slide_background: applied on slide.Background")
            changed = True

    # ── Quiet fallback: direct DrawPage properties (works in Writer/Calc, not Impress) ──
    if not changed:
        try:
            changed = _apply_background_fill(slide)
            if changed:
                _log("V2/set_slide_background: applied direct fill properties on slide")
        except Exception:
            pass

    if not changed:
        try:
            if rgb is not None and hasattr(slide, "setPropertyValue"):
                slide.setPropertyValue("BackTransparent", False)
                slide.setPropertyValue("BackColor", rgb)
                changed = True
                _log("V2/set_slide_background: applied BackTransparent+BackColor on slide")
        except Exception:
            pass

    # ── Quiet: full-page background coverage where supported ──
    try:
        if hasattr(slide, "setPropertyValue"):
            slide.setPropertyValue("BackgroundFullSize", True)
    except Exception:
        pass

    if changed:
        return True

    # Final debug breadcrumb when all strategies failed.
    try:
        props = slide.getPropertySetInfo().getProperties()
        prop_names = [p.Name for p in props] if props is not None else []
        _log_json("V2/set_slide_background: available slide properties=", prop_names)
    except Exception as exc:
        _log(f"V2/set_slide_background: unable to enumerate slide properties: {exc}")

    _log("V2/set_slide_background: failed to apply background color on all strategies")
    return False


def _set_slide_transition(slide, transition):
    if not isinstance(transition, dict):
        return False
    changed = False
    try:
        speed = transition.get("duration")
        if speed is not None and hasattr(slide, "TransitionDuration"):
            slide.TransitionDuration = float(speed)
            changed = True
    except Exception:
        pass
    for src_key, prop_name in (
        ("type", "TransitionType"),
        ("subtype", "TransitionSubtype"),
        ("auto_advance", "TransitionAutoAdvance"),
    ):
        value = transition.get(src_key)
        if value is None:
            continue
        try:
            if _set_uno_property(slide, prop_name, value):
                changed = True
        except Exception:
            continue
    return changed


def _apply_slide_plan_to_slide(document, slide, plan):
    """
    Minimal rendering for slide-level operations in live mode.
    This intentionally keeps scope small and deterministic:
    - one title textbox
    - one body textbox from content_for_artist
    - optional notes from speaker_notes
    """
    if not isinstance(plan, dict):
        return False
    slides = plan.get("slides")
    if not isinstance(slides, list) or not slides:
        return False
    slide_req = slides[0] if isinstance(slides[0], dict) else {}
    if not isinstance(slide_req, dict):
        return False

    default_slide_style = plan.get("default_slide_style") if isinstance(plan.get("default_slide_style"), dict) else {}
    slide_style_for_artist = (
        slide_req.get("slide_style_for_artist")
        if isinstance(slide_req.get("slide_style_for_artist"), dict)
        else {}
    )
    merged_slide_style = dict(default_slide_style)
    merged_slide_style.update(slide_style_for_artist)
    if merged_slide_style:
        background_payload = {}
        if isinstance(merged_slide_style.get("gradient"), dict):
            background_payload["gradient"] = merged_slide_style.get("gradient")
        if merged_slide_style.get("background_color") is not None:
            background_payload["background_color"] = merged_slide_style.get("background_color")
        if merged_slide_style.get("color") is not None:
            background_payload["color"] = merged_slide_style.get("color")
        if background_payload:
            _set_slide_background(slide, background_payload, document=document)

    title_text = slide_req.get("inspiration_pattern") or "Slide"
    body_text = slide_req.get("content_for_artist") or slide_req.get("prompt_for_artist") or ""
    notes_text = slide_req.get("speaker_notes")

    try:
        title_shape = document.createInstance("com.sun.star.drawing.TextShape")
        title_shape.setPosition(_create_point(_pt_to_mm100(40), _pt_to_mm100(20)))
        title_shape.setSize(_create_size(_pt_to_mm100(1200), _pt_to_mm100(70)))
        title_shape.getText().setString(str(title_text))
        slide.add(title_shape)
    except Exception:
        pass

    try:
        body_shape = document.createInstance("com.sun.star.drawing.TextShape")
        body_shape.setPosition(_create_point(_pt_to_mm100(60), _pt_to_mm100(120)))
        body_shape.setSize(_create_size(_pt_to_mm100(1160), _pt_to_mm100(520)))
        body_shape.getText().setString(str(body_text))
        slide.add(body_shape)
    except Exception:
        pass

    _set_slide_notes(slide, notes_text)
    return True


def _dispatch_duplicate_page(document, controller, draw_pages, source_idx_zero_based):
    source_page = _get_page(draw_pages, source_idx_zero_based)
    if source_page is None:
        return None

    # Best path: XDrawPageDuplicator
    try:
        if hasattr(draw_pages, "duplicate"):
            duplicated = draw_pages.duplicate(source_page)
            if duplicated is not None:
                return duplicated
    except Exception:
        pass

    # Fallback: dispatch duplicate command on focused page
    frame = controller.getFrame()
    if frame is None:
        return None
    try:
        controller.setCurrentPage(source_page)
    except Exception:
        pass
    try:
        helper = document.getCurrentController().getFrame().getController().getModel().getCurrentController()
        _ = helper  # keep linter quiet in non-UNO tooling
    except Exception:
        pass
    try:
        dh = document.getCurrentController().getFrame().queryDispatch
        _ = dh  # no-op; dispatch helper path below
    except Exception:
        pass
    try:
        service_manager = XSCRIPTCONTEXT.getComponentContext().ServiceManager  # type: ignore # NOQA
        dispatcher = service_manager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper",
            XSCRIPTCONTEXT.getComponentContext(),  # type: ignore # NOQA
        )
        dispatcher.executeDispatch(frame, ".uno:DuplicatePage", "", 0, ())
        # New page is usually after source page.
        return _get_page(draw_pages, min(source_idx_zero_based + 1, draw_pages.getCount() - 1))
    except Exception:
        return None


def _is_full_component_payload(payload):
    if not isinstance(payload, dict):
        return False
    comp_type = str(payload.get("type") or "").strip().lower()
    if not comp_type:
        return False
    if comp_type == "connector":
        return bool(payload.get("source_id")) and bool(payload.get("target_id"))
    if isinstance(payload.get("children"), list) or isinstance(payload.get("items"), list):
        return True
    if _payload_has_explicit_geometry(payload):
        return True
    # Table/chart payloads are typically full object replacements.
    if comp_type == "table" and isinstance(payload.get("table_data"), dict):
        return True
    if comp_type == "chart" and isinstance(payload.get("chart_data"), dict):
        return True
    # Container/list payloads are rendered as subtrees.
    if comp_type in {"container", "vstack", "hstack", "grid", "zstack", "list"}:
        return True
    return False


def _payload_has_explicit_geometry(payload):
    if not isinstance(payload, dict):
        return False
    for key in ("geometry", "position", "size"):
        section = payload.get(key)
        if not isinstance(section, dict):
            continue
        if any(section.get(axis) is not None for axis in ("x", "y", "width", "height")):
            return True
    return False


def _shape_geometry_payload(shape):
    if shape is None:
        return {}
    try:
        pos = shape.getPosition()
        size = shape.getSize()
    except Exception:
        return {}

    x_pt = _mm100_to_pt(getattr(pos, "X", None))
    y_pt = _mm100_to_pt(getattr(pos, "Y", None))
    w_pt = _mm100_to_pt(getattr(size, "Width", None))
    h_pt = _mm100_to_pt(getattr(size, "Height", None))
    if None in (x_pt, y_pt, w_pt, h_pt):
        return {}

    return {
        "geometry": {"x": x_pt, "y": y_pt, "width": w_pt, "height": h_pt},
        "position": {"x": x_pt, "y": y_pt},
        "size": {"width": w_pt, "height": h_pt},
    }


def _action_replace_component(document, draw_pages, action):
    target = action.get("target") or {}
    payload = action.get("payload") if isinstance(action.get("payload"), dict) else {}
    slide_idx = _resolve_slide_index(target, draw_pages=draw_pages)
    element_id = target.get("element_id") if isinstance(target, dict) else None
    _log_json("V2/replace_component: action=", action)
    _log(
        "V2/replace_component: "
        f"resolved slide_idx={slide_idx} element_id={element_id} payload_type={payload.get('type') if isinstance(payload, dict) else None}"
    )
    if slide_idx is None:
        _log("V2/replace_component: missing slide index")
        return {"status": "error", "error": "missing_slide_index"}

    slide = _get_page(draw_pages, slide_idx - 1)
    if slide is None:
        _log(f"V2/replace_component: slide out of range slide_idx={slide_idx}")
        return {"status": "error", "error": "slide_out_of_range"}

    shape, shape_idx = _find_shape_by_canonical_id(slide, element_id) if element_id else (None, -1)
    full_component_payload = _is_full_component_payload(payload)

    # Full replacement path (backend parity): remove target branch then render payload subtree.
    if full_component_payload:
        component_payload = dict(payload)
        if shape is not None and not _payload_has_explicit_geometry(component_payload):
            preserved_geometry = _shape_geometry_payload(shape)
            for key, value in preserved_geometry.items():
                existing = component_payload.get(key)
                if not isinstance(existing, dict) or not existing:
                    component_payload[key] = value
            if preserved_geometry:
                _log("V2/replace_component: preserved target geometry for full_replace payload")

        removed_count = 0
        if element_id:
            matched = _collect_shapes_for_target(slide, element_id, include_descendants=True)
            removed_count = _remove_shapes(slide, matched)

        if element_id and not component_payload.get("id"):
            component_payload["id"] = element_id

        id_map = {}
        connectors = []
        _render_blueprint_component(
            document,
            slide,
            component_payload,
            0.0,
            0.0,
            id_map,
            connectors,
            ancestors=[],
        )
        for connector_component in connectors:
            _render_connector(document, slide, connector_component, id_map)

        rendered_shape = None
        rendered_id = component_payload.get("id")
        if rendered_id:
            rendered_shape, _ = _find_shape_by_canonical_id(slide, rendered_id)
        _log(
            "V2/replace_component: full_replace "
            f"removed_count={removed_count} rendered_id={rendered_id}"
        )
        return {
            "status": "ok",
            "changed": True,
            "slide_index": slide_idx,
            "_shape": rendered_shape,
        }

    # Partial replacement path (in-place mutation).
    if shape is None:
        _log(f"V2/replace_component: element_not_found element_id={element_id} slide_idx={slide_idx}")
        return {"status": "error", "error": "element_not_found", "slide_index": slide_idx}
    _log_json("V2/replace_component: matched_shape=", _shape_debug_info(shape))
    _log(f"V2/replace_component: matched_shape_index={shape_idx}")

    changed = _apply_payload_to_shape(shape, payload)
    _log(f"V2/replace_component: changed={changed}")
    return {"status": "ok", "changed": bool(changed), "slide_index": slide_idx, "_shape": shape if changed else None}


def _action_update_style(draw_pages, action):
    target = action.get("target") or {}
    payload = action.get("payload") if isinstance(action.get("payload"), dict) else {}
    slide_idx = _resolve_slide_index(target, draw_pages=draw_pages)
    element_id = target.get("element_id") if isinstance(target, dict) else None
    _log_json("V2/update_style: action=", action)
    _log(f"V2/update_style: resolved slide_idx={slide_idx} element_id={element_id}")
    if slide_idx is None or not element_id:
        _log("V2/update_style: missing target")
        return {"status": "error", "error": "missing_target"}

    slide = _get_page(draw_pages, slide_idx - 1)
    if slide is None:
        _log(f"V2/update_style: slide out of range slide_idx={slide_idx}")
        return {"status": "error", "error": "slide_out_of_range"}

    shape, shape_idx = _find_shape_by_canonical_id(slide, element_id)
    if shape is None:
        _log(f"V2/update_style: element_not_found element_id={element_id}")
        return {"status": "error", "error": "element_not_found"}
    _log_json("V2/update_style: matched_shape=", _shape_debug_info(shape))
    _log(f"V2/update_style: matched_shape_index={shape_idx}")

    changed = _apply_style(shape, payload)
    _log(f"V2/update_style: changed={changed}")
    return {"status": "ok", "changed": bool(changed), "slide_index": slide_idx}


def _action_delete_element(draw_pages, action):
    target = action.get("target") or {}
    slide_idx = _resolve_slide_index(target, draw_pages=draw_pages)
    element_id = target.get("element_id") if isinstance(target, dict) else None
    _log_json("V2/delete_element: action=", action)
    _log(f"V2/delete_element: resolved slide_idx={slide_idx} element_id={element_id}")
    if slide_idx is None or not element_id:
        _log("V2/delete_element: missing target")
        return {"status": "error", "error": "missing_target"}

    slide = _get_page(draw_pages, slide_idx - 1)
    if slide is None:
        _log(f"V2/delete_element: slide out of range slide_idx={slide_idx}")
        return {"status": "error", "error": "slide_out_of_range"}

    matched = _collect_shapes_for_target(slide, element_id, include_descendants=True)
    if not matched:
        _log(f"V2/delete_element: element_not_found element_id={element_id}")
        return {"status": "error", "error": "element_not_found"}
    removed_count = _remove_shapes(slide, matched)
    _log(f"V2/delete_element: removed_count={removed_count}")
    return {
        "status": "ok" if removed_count > 0 else "error",
        "slide_index": slide_idx,
        "removed_count": int(removed_count),
    }


def _action_insert_slide(document, draw_pages, action):
    raw_index = action.get("index")
    count = draw_pages.getCount()
    try:
        idx = int(raw_index) - 1 if raw_index is not None else count
    except Exception:
        idx = count
    idx = max(0, min(idx, count))
    slide = draw_pages.insertNewByIndex(idx)
    _prepare_blank_slide_surface(document, slide)
    rendered = False
    blueprint = action.get("blueprint")
    if isinstance(blueprint, dict):
        target_canvas = _get_page_canvas_size(slide) or _get_document_canvas_size(document, draw_pages)
        scaled_blueprint = _scale_blueprint_to_canvas(
            blueprint,
            target_canvas.get("width_pt"),
            target_canvas.get("height_pt"),
        )
        rendered = _render_blueprint_to_slide(document, slide, scaled_blueprint)
    if not rendered:
        plan = action.get("plan")
        _apply_slide_plan_to_slide(document, slide, plan if isinstance(plan, dict) else {})
    return {"status": "ok", "slide_index": idx + 1}


def _action_replace_slide(document, draw_pages, action):
    target = action.get("target") or {}
    slide_idx = _resolve_slide_index(target, draw_pages=draw_pages)
    if slide_idx is None:
        return {"status": "error", "error": "missing_slide_index"}
    idx = slide_idx - 1
    slide = _get_page(draw_pages, idx)
    if slide is None:
        return {"status": "error", "error": "slide_out_of_range"}

    # Do not remove+insert here: removing the only slide can trigger Impress to
    # auto-create a default title slide, causing count/index drift.
    _prepare_blank_slide_surface(document, slide)
    rendered = False
    blueprint = action.get("blueprint")
    if isinstance(blueprint, dict):
        target_canvas = _get_page_canvas_size(slide) or _get_document_canvas_size(document, draw_pages)
        scaled_blueprint = _scale_blueprint_to_canvas(
            blueprint,
            target_canvas.get("width_pt"),
            target_canvas.get("height_pt"),
        )
        rendered = _render_blueprint_to_slide(document, slide, scaled_blueprint)
    if not rendered:
        plan = action.get("plan")
        _apply_slide_plan_to_slide(document, slide, plan if isinstance(plan, dict) else {})
    return {"status": "ok", "slide_index": slide_idx}


def _action_delete_slide(draw_pages, action):
    target = action.get("target") or {}
    slide_idx = _resolve_slide_index(target, draw_pages=draw_pages)
    if slide_idx is None:
        return {"status": "error", "error": "missing_slide_index"}
    idx = slide_idx - 1
    page = _get_page(draw_pages, idx)
    if page is None:
        return {"status": "error", "error": "slide_out_of_range"}
    try:
        draw_pages.remove(page)
        return {"status": "ok", "slide_index": max(1, slide_idx - 1)}
    except Exception as exc:
        return {"status": "error", "error": "remove_failed", "message": str(exc)}


def _action_duplicate_slide(document, controller, draw_pages, action):
    source = action.get("source") or {}
    source_idx = _resolve_slide_index(source, draw_pages=draw_pages)
    if source_idx is None:
        return {"status": "error", "error": "missing_source"}

    duplicated = _dispatch_duplicate_page(document, controller, draw_pages, source_idx - 1)
    if duplicated is None:
        return {"status": "error", "error": "duplicate_failed"}

    # Best-effort slide index discovery and optional reposition.
    new_idx_zero = _find_page_index(draw_pages, duplicated)
    if new_idx_zero is None:
        new_idx_zero = source_idx

    target_idx = action.get("target_index")
    if target_idx is not None:
        try:
            target_zero = int(target_idx) - 1
            moved_idx = _move_page_to_index(document, controller, draw_pages, duplicated, target_zero)
            if moved_idx is not None:
                new_idx_zero = moved_idx
        except Exception:
            pass

    return {"status": "ok", "slide_index": int(new_idx_zero) + 1}


def _action_update_slide_properties(document, draw_pages, action):
    target = action.get("target") or {}
    slide_idx = _resolve_slide_index(target, draw_pages=draw_pages)
    if slide_idx is None:
        return {"status": "error", "error": "missing_slide_index"}
    slide = _get_page(draw_pages, slide_idx - 1)
    if slide is None:
        return {"status": "error", "error": "slide_out_of_range"}

    payload = action.get("payload") if isinstance(action.get("payload"), dict) else {}
    notes_value = action.get("notes")
    if notes_value is None:
        notes_value = payload.get("notes")

    background_value = action.get("background")
    if background_value is None:
        background_value = payload.get("background")
    if background_value is None and payload.get("background_color") is not None:
        background_value = {"background_color": payload.get("background_color")}

    transition_value = action.get("transition")
    if transition_value is None:
        transition_value = payload.get("transition")

    changed = False
    if _set_slide_notes(slide, notes_value):
        changed = True
    if _set_slide_background(slide, background_value, document=document):
        changed = True
    if _set_slide_transition(slide, transition_value):
        changed = True
    return {"status": "ok", "changed": changed, "slide_index": slide_idx}


def _action_navigate_to_slide(document, controller, draw_pages, action):
    target = action.get("target") or {}
    slide_idx = _resolve_slide_index(target, draw_pages=draw_pages)
    if slide_idx is None:
        return {"status": "error", "error": "missing_slide_index"}

    slide = _get_page(draw_pages, slide_idx - 1)
    if slide is None:
        return {"status": "error", "error": "slide_out_of_range"}

    try:
        _nudge_slide_selection(controller, draw_pages, slide)
    except Exception:
        try:
            controller.setCurrentPage(slide)
        except Exception as exc:
            return {"status": "error", "error": "navigation_failed", "message": str(exc)}

    return {"status": "ok", "changed": False, "slide_index": slide_idx}


# Typed replace_* actions (replace_text_box, replace_image, etc.) are
# collapsed back to "replace_component" in live_plan._collapse_typed_replace_kinds
# before the plan reaches this script, so the executor only knows one kind.
_HANDLERS = {
    "replace_component": _action_replace_component,
    "update_style": _action_update_style,
    "delete_element": _action_delete_element,
    "insert_slide": _action_insert_slide,
    "replace_slide": _action_replace_slide,
    "delete_slide": _action_delete_slide,
    "duplicate_slide": _action_duplicate_slide,
    "update_slide_properties": _action_update_slide_properties,
    "navigate_to_slide": _action_navigate_to_slide,
}


def applySlidesV2ActionPlan(plan_json: str) -> str:
    _log("ApplySlidesV2ActionPlan: received payload")
    try:
        plan = json.loads(plan_json)
    except Exception as exc:
        _log(f"ApplySlidesV2ActionPlan: invalid json: {exc}")
        return json.dumps({"ok": False, "error": "invalid_json", "message": str(exc)})

    plan_id = plan.get("plan_id") or "slides-v2-plan"
    actions = plan.get("actions")
    _log_json("ApplySlidesV2ActionPlan: parsed plan=", plan)
    raw_host_managed_save = plan.get("host_managed_save", None)
    raw_save_after_each = plan.get("save_after_each_action", None)
    # Slides V2 defaults to host-managed save with no script-side store().
    # This keeps save authority in one place (frontend Action_Save -> WOPI PutFile).
    host_managed_save = (
        True if raw_host_managed_save is None else bool(raw_host_managed_save)
    )
    save_after_each = False if raw_save_after_each is None else bool(raw_save_after_each)
    if host_managed_save and save_after_each:
        _log(
            "ApplySlidesV2ActionPlan: "
            "host_managed_save=True overrides save_after_each_action=True -> False"
        )
        save_after_each = False
    _log(
        "ApplySlidesV2ActionPlan: "
        f"plan_id={plan_id} action_count={len(actions) if isinstance(actions, list) else 'invalid'} "
        f"save_after_each_action={save_after_each} "
        f"host_managed_save={host_managed_save} "
        f"(raw_save_after_each_action={raw_save_after_each!r}, raw_host_managed_save={raw_host_managed_save!r})"
    )
    if not isinstance(actions, list):
        return json.dumps({"ok": False, "error": "invalid_actions", "plan_id": plan_id})

    try:
        document = XSCRIPTCONTEXT.getDocument()  # type: ignore # NOQA
    except Exception as exc:
        _log(f"ApplySlidesV2ActionPlan: no document: {exc}")
        return json.dumps({"ok": False, "error": "no_document", "plan_id": plan_id})

    if document is None:
        return json.dumps({"ok": False, "error": "no_document", "plan_id": plan_id})

    controller = document.getCurrentController()
    if controller is None:
        return json.dumps({"ok": False, "error": "no_controller", "plan_id": plan_id})

    draw_pages = _get_draw_pages(document)
    if draw_pages is None:
        return json.dumps({"ok": False, "error": "no_draw_pages", "plan_id": plan_id})

    document_canvas = _get_document_canvas_size(document, draw_pages, controller=controller)
    inferred_canvas_width_pt, inferred_canvas_height_pt = _infer_canvas_size_from_actions(actions)
    canvas_update = {
        "mode": "inherit_existing_document_canvas",
        "document_canvas": document_canvas,
        "inferred_blueprint_canvas": {
            "width_pt": float(inferred_canvas_width_pt),
            "height_pt": float(inferred_canvas_height_pt),
            "width_mm100": int(_pt_to_mm100(inferred_canvas_width_pt)),
            "height_mm100": int(_pt_to_mm100(inferred_canvas_height_pt)),
        },
    }
    _log_json("ApplySlidesV2ActionPlan: canvas_update=", canvas_update)

    results = []
    wrote = False
    last_slide_index = None
    deferred_shapes = []  # shapes to nudge/focus after unlockControllers

    locked = False
    try:
        document.lockControllers()
        locked = True
    except Exception:
        pass

    try:
        for idx, action in enumerate(actions):
            if not isinstance(action, dict):
                results.append(
                    {
                        "index": idx,
                        "operation_id": f"op-{idx}",
                        "status": "error",
                        "error": "invalid_action_object",
                    }
                )
                continue

            kind = action.get("kind")
            operation_id = action.get("operation_id") or f"op-{idx}"
            _log(
                "ApplySlidesV2ActionPlan: "
                f"starting action idx={idx} operation_id={operation_id} kind={kind}"
            )
            _log_json("ApplySlidesV2ActionPlan: action payload=", action)
            handler = _HANDLERS.get(kind)
            if handler is None:
                results.append(
                    {
                        "index": idx,
                        "operation_id": operation_id,
                        "kind": kind,
                        "status": "error",
                        "error": "unsupported_action",
                    }
                )
                continue

            try:
                if kind == "replace_component":
                    out = handler(document, draw_pages, action)
                elif kind in {"insert_slide", "replace_slide"}:
                    out = handler(document, draw_pages, action)
                elif kind == "update_slide_properties":
                    out = handler(document, draw_pages, action)
                elif kind in {"duplicate_slide", "navigate_to_slide"}:
                    out = handler(document, controller, draw_pages, action)
                else:
                    out = handler(draw_pages, action)
            except Exception as exc:
                out = {
                    "status": "error",
                    "error": "handler_exception",
                    "message": str(exc),
                    "traceback": traceback.format_exc(),
                }
            # Collect shapes for deferred invalidation (after unlockControllers).
            # Pop _shape BEFORE logging — UNO objects are not JSON-serializable.
            deferred_shape = out.pop("_shape", None)
            if deferred_shape is not None:
                deferred_shapes.append(deferred_shape)

            _log_json("ApplySlidesV2ActionPlan: action out=", out)

            status = out.get("status", "error")
            result = {
                "index": idx,
                "operation_id": operation_id,
                "kind": kind,
                "status": status,
            }
            if status != "ok":
                result["error"] = out.get("error", "unknown_error")
                if out.get("message"):
                    result["message"] = out.get("message")
            else:
                changed = out.get("changed")
                if changed is None or bool(changed):
                    wrote = True
            if out.get("slide_index") is not None:
                last_slide_index = out.get("slide_index")
            results.append(result)
            _log_json("ApplySlidesV2ActionPlan: normalized action result=", result)

            if status == "ok" and save_after_each and not host_managed_save:
                try:
                    document.store()
                    _log("ApplySlidesV2ActionPlan: document.store() succeeded (save_after_each_action=True)")
                except Exception as exc:
                    _log(f"ApplySlidesV2ActionPlan: store failed: {exc}")

        if wrote and not save_after_each and not host_managed_save:
            try:
                document.store()
                _log("ApplySlidesV2ActionPlan: final document.store() succeeded")
            except Exception as exc:
                _log(f"ApplySlidesV2ActionPlan: final store failed: {exc}")
        elif wrote and host_managed_save:
            # In host-managed save mode, avoid direct store() here.
            # The frontend will issue Action_Save, which persists via WOPI and
            # keeps the live editing session more stable for repaint updates.
            _log("ApplySlidesV2ActionPlan: deferring save to host (host_managed_save=True)")
    finally:
        if locked:
            try:
                document.unlockControllers()
            except Exception:
                pass

    # ── Post-unlock: Force Collabora/LOK to re-render tiles ──
    # All invalidation MUST happen here, outside lockControllers().
    # Under lock, UI events are suppressed and invalidation signals are lost.
    if wrote:
        # Step 1: Small delay to let LOK process the unlock event.
        time.sleep(0.10)
        _log("V2/post_unlock: unlock settle complete")

        # Step 2: setModified(True) — canonical LOK tile invalidation trigger.
        # This fires LOK_CALLBACK_INVALIDATE_TILES in the LOK rendering pipeline.
        try:
            document.setModified(True)
            _log("V2/post_unlock: setModified applied")
        except Exception as exc:
            _log(f"V2/post_unlock: setModified failed: {exc}")

        # Step 3: Deferred shape nudge + focus for shapes mutated under lock.
        if deferred_shapes:
            _log(f"V2/post_unlock: deferred nudge on {len(deferred_shapes)} shapes")
            for shape in deferred_shapes:
                try:
                    _nudge_shape_invalidation(shape)
                except Exception:
                    pass
            # Focus the last mutated shape to trigger a visible UI event.
            try:
                _focus_shape(document, deferred_shapes[-1])
            except Exception:
                pass

        # Step 4: Slide selection bounce (forces tile cache re-fetch).
        if last_slide_index is not None:
            target_page = None
            try:
                target_page = _get_page(draw_pages, last_slide_index - 1)
                if target_page is not None:
                    _nudge_slide_selection(controller, draw_pages, target_page)
                    _log(f"V2/post_unlock: slide bounce to slide {last_slide_index}")
            except Exception as exc:
                _log(f"V2/post_unlock: slide bounce failed: {exc}")

        # Step 5: Re-apply the current controller view state so LOK recomputes
        # visible area/zoom after any in-session canvas resize.
        try:
            if controller is not None:
                controller.restoreViewData(controller.getViewData())
                _log("V2/post_unlock: controller view restored")
        except Exception as exc:
            _log(f"V2/post_unlock: controller restore failed: {exc}")

        # Step 6: Dispatch an Impress-valid UNO command for view refresh.
        try:
            _force_ui_refresh(controller)
        except Exception as exc:
            _log(f"V2/post_unlock: refresh dispatch failed: {exc}")

    response = json.dumps(
        {
            "ok": True,
            "plan_id": plan_id,
            "file_name": plan.get("file_name"),
            "save_after_each_action": save_after_each,
            "host_managed_save": host_managed_save,
            "results": results,
            "last_slide_index": last_slide_index,
        }
    )
    _log(f"ApplySlidesV2ActionPlan: returning response={response}")
    return response


g_exportedScripts = (applySlidesV2ActionPlan,)
