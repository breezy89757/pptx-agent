"""Font safety list (spec's flagged risk: "中文字型在瀏覽器與 PowerPoint 之間的
度量差異"). We restrict the Design System Agent to fonts that ship with
Windows/Office by default, so the same physical font file renders the HTML
preview (Playwright/Chromium picks up the OS-installed font too) and the
final .pptx on any machine that opens it -- no substitution, no metric drift.
"""

from __future__ import annotations

# name -> short description shown to the LLM when picking.
SAFE_FONTS: dict[str, str] = {
    "Microsoft JhengHei": "現代無襯線,Windows 內建繁中字型,標題/內文皆適用,首選",
    "Microsoft JhengHei UI": "同上的 UI 變體,字面略緊,適合密集資訊排版",
    "PMingLiU": "新細明體,Windows 內建繁中襯線字型,適合正式/傳統調性的標題",
    "DFKai-SB": "標楷體,楷書風格,只適合大字級標題或強調文字,不適合內文",
    "Segoe UI": "Windows 內建英數字型,適合純英數的標籤/數字",
    "Arial": "廣泛內建的英數無襯線字型,適合純英數的標籤/數字",
    "Georgia": "廣泛內建的英數襯線字型,適合純英數的強調標題",
}

DEFAULT_HEADING_FONT = "Microsoft JhengHei"
DEFAULT_BODY_FONT = "Microsoft JhengHei"


def clamp_to_safe_font(font_name: str, *, fallback: str) -> str:
    """Case-insensitive match against the safe list; falls back to a known-
    safe default if the LLM picked something outside it despite instructions."""
    normalized = font_name.strip().lower()
    for safe_name in SAFE_FONTS:
        if safe_name.lower() == normalized:
            return safe_name
    return fallback
