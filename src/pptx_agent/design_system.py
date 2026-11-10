from __future__ import annotations

import json
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class LoadedDesignSystem:
    prompt: str
    token_overrides: dict[str, str]


def _color_value(tokens: dict[str, Any], *names: str) -> str:
    color_tokens = tokens.get("color", {}).get("tokens", [])
    values = {item.get("name", "").casefold(): item.get("value", "") for item in color_tokens}
    for name in names:
        value = values.get(name.casefold())
        if value:
            normalized = str(value).lstrip("#")
            if not re.fullmatch(r"[0-9a-fA-F]{6}", normalized):
                raise ValueError(f"Design-system color {name!r} must be a six-digit hex value")
            return normalized.upper()
    raise ValueError(f"Design-system tokens.json is missing a color token: {' / '.join(names)}")


def load_design_system(path: Path, run_dir: Path) -> LoadedDesignSystem:
    """Load a README/tokens design system and stage its assets beside deck.html."""
    root = path.expanduser().resolve()
    readme_path = root / "README.md"
    tokens_path = root / "tokens.json"
    if not readme_path.is_file() or not tokens_path.is_file():
        raise FileNotFoundError("--design-system must be a directory containing README.md and tokens.json")

    readme = readme_path.read_text(encoding="utf-8")
    tokens: dict[str, Any] = json.loads(tokens_path.read_text(encoding="utf-8"))
    brand = _color_value(tokens, "brand-red", "brand-primary", "primary")
    rule = _color_value(tokens, "rule-red", "rule", "secondary")
    ink = _color_value(tokens, "ink", "text-primary", "text")
    surface = _color_value(tokens, "surface", "background", "white")

    asset_source = root / "assets"
    staged_assets = run_dir / "assets"
    asset_paths: list[str] = []
    if asset_source.is_dir():
        shutil.copytree(asset_source, staged_assets, dirs_exist_ok=True)
        asset_paths = [
            item.relative_to(run_dir).as_posix()
            for item in sorted(staged_assets.rglob("*"))
            if item.is_file()
        ]

    asset_list = "\n".join(f"- {asset_path}" for asset_path in asset_paths) or "- (none)"
    prompt = (
        "Treat the following design-system specification as mandatory.\n\n"
        f"README.md:\n{readme}\n\n"
        f"tokens.json:\n{json.dumps(tokens, ensure_ascii=False, indent=2)}\n\n"
        "Assets copied next to deck.html; use these relative paths in img src:\n"
        f"{asset_list}\n\n"
        "The palette is enforced by the application: brand color is for graphic accents, rule color is only for title rules, "
        "background/surface stay white, and all text stays ink. Do not add colors, gradients, or shadows."
    )
    overrides = {
        "theme_name": str(tokens.get("name", "Design system")),
        "primary_color": brand,
        "secondary_color": rule,
        "accent_color": brand,
        "background_color": surface,
        "surface_color": surface,
        "text_color": ink,
        "muted_text_color": ink,
    }
    return LoadedDesignSystem(prompt=prompt, token_overrides=overrides)