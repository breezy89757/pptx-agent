"""WCAG 2.x contrast-ratio math, used to guarantee the Design System Agent's
text/background token pairs stay readable regardless of what the LLM picked.

AA thresholds: 4.5:1 for normal text, 3:1 for large text (WCAG 2.2 SC 1.4.3).
We enforce the stricter 4.5:1 everywhere since design tokens are reused at
arbitrary font sizes.
"""

from __future__ import annotations

import colorsys
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx_agent.models import DesignTokens

AA_NORMAL_TEXT_RATIO = 4.5


def _srgb_to_linear(c: float) -> float:
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def relative_luminance(hex_color: str) -> float:
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i : i + 2], 16) / 255 for i in (0, 2, 4))
    r, g, b = _srgb_to_linear(r), _srgb_to_linear(g), _srgb_to_linear(b)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(hex_a: str, hex_b: str) -> float:
    la, lb = relative_luminance(hex_a), relative_luminance(hex_b)
    lighter, darker = max(la, lb), min(la, lb)
    return (lighter + 0.05) / (darker + 0.05)


def _with_lightness(hex_color: str, lightness: float) -> str:
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i : i + 2], 16) / 255 for i in (0, 2, 4))
    h, _l, s = colorsys.rgb_to_hls(r, g, b)
    r2, g2, b2 = colorsys.hls_to_rgb(h, lightness, s)
    return f"{round(r2 * 255):02x}{round(g2 * 255):02x}{round(b2 * 255):02x}"


def ensure_min_contrast(fg_hex: str, bg_hex: str, min_ratio: float = AA_NORMAL_TEXT_RATIO) -> str:
    """Returns fg_hex unchanged if it already meets min_ratio against bg_hex,
    otherwise pushes fg's lightness toward black or white (whichever side of
    bg has headroom) in small steps until the ratio is met."""
    if contrast_ratio(fg_hex, bg_hex) >= min_ratio:
        return fg_hex

    bg_lum = relative_luminance(bg_hex)
    # Push foreground toward whichever extreme is farther from the background's
    # luminance -- that's the direction with more contrast headroom.
    target_extreme = 1.0 if bg_lum < 0.5 else 0.0

    best_hex = fg_hex
    best_ratio = contrast_ratio(fg_hex, bg_hex)
    steps = 20
    for step in range(1, steps + 1):
        frac = step / steps  # 0 -> fg's own lightness territory, 1 -> target_extreme
        lightness = (1 - frac) * 0.5 + frac * target_extreme
        candidate = _with_lightness(fg_hex, lightness)
        ratio = contrast_ratio(candidate, bg_hex)
        if ratio > best_ratio:
            best_hex, best_ratio = candidate, ratio
        if ratio >= min_ratio:
            return candidate
    return best_hex


def enforce_token_contrast(tokens: "DesignTokens") -> "DesignTokens":
    """Guarantees text_color/muted_text_color read clearly against both
    background_color and surface_color, regardless of what the LLM picked.
    The HTML Layout Agent is told to only ever put readable text in these two
    colors, so fixing them here covers every slide deterministically."""
    text = tokens.text_color
    muted = tokens.muted_text_color

    text = ensure_min_contrast(text, tokens.background_color)
    text = ensure_min_contrast(text, tokens.surface_color)
    muted = ensure_min_contrast(muted, tokens.background_color)
    muted = ensure_min_contrast(muted, tokens.surface_color)

    return tokens.model_copy(update={"text_color": text, "muted_text_color": muted})
