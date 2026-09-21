from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from rich.console import Console
from rich.markup import escape

from pptx_agent.agents.content_agent import run_content_agent
from pptx_agent.agents.design_agent import run_design_agent
from pptx_agent.agents.layout_agent import CANVAS_H, CANVAS_W, run_layout_agent
from pptx_agent.models import Deck, DesignTokens, SlideHtml

TEMPLATES_DIR = Path(__file__).parent / "templates"
console = Console()


def render_deck_html(run_dir: Path) -> Path:
    """Renders deck.html purely from the on-disk JSON contracts (content.json,
    design_tokens.json, slides_html.json). Reusable both for the initial
    build and for the Visual QA loop, which patches slides_html.json for the
    flagged indices and needs to re-render without redoing earlier stages."""
    deck = Deck.model_validate_json((run_dir / "content.json").read_text(encoding="utf-8"))
    tokens = DesignTokens.model_validate_json((run_dir / "design_tokens.json").read_text(encoding="utf-8"))
    slide_htmls = [
        SlideHtml.model_validate(s) for s in json.loads((run_dir / "slides_html.json").read_text(encoding="utf-8"))
    ]
    slide_htmls.sort(key=lambda s: s.index)

    env = Environment(
        loader=FileSystemLoader(str(TEMPLATES_DIR)),
        autoescape=select_autoescape(disabled_extensions=("jinja",)),
    )
    template = env.get_template("deck.html.jinja")
    deck_html = template.render(
        deck_title=deck.deck_title,
        tokens=tokens,
        slides=slide_htmls,
        canvas_w=CANVAS_W,
        canvas_h=CANVAS_H,
        variants=", ".join(sorted({s.layout_variant for s in slide_htmls})),
    )
    deck_path = run_dir / "deck.html"
    deck_path.write_text(deck_html, encoding="utf-8")
    return deck_path


def run_pipeline(
    outline: str,
    *,
    output_root: Path,
    target_slide_count: int | None = None,
    run_name: str | None = None,
) -> Path:
    run_id = run_name or dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = output_root / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    console.print("[bold cyan]1/3 Content Agent[/bold cyan] 拆解大綱為投影片內容...")
    deck = run_content_agent(outline, target_slide_count=target_slide_count)
    (run_dir / "content.json").write_text(deck.model_dump_json(indent=2), encoding="utf-8")
    console.print(escape(f"  -> {len(deck.slides)} 張投影片: " + ", ".join(f"[{s.content_type.value}] {s.title}" for s in deck.slides)))

    console.print("[bold cyan]2/3 Design System Agent[/bold cyan] 產生設計令牌...")
    tokens = run_design_agent(deck)
    (run_dir / "design_tokens.json").write_text(tokens.model_dump_json(indent=2), encoding="utf-8")
    console.print(f"  -> theme={tokens.theme_name} primary=#{tokens.primary_color} font={tokens.font_heading}/{tokens.font_body}")

    console.print("[bold cyan]3/3 HTML Layout Agent[/bold cyan] 逐頁產生 HTML 版面...")
    slide_htmls: list[SlideHtml] = []
    for slide in deck.slides:
        slide_html = run_layout_agent(slide, tokens)
        slide_htmls.append(slide_html)
        console.print(escape(f"  -> slide #{slide.index} [{slide.content_type.value}] variant={slide_html.layout_variant}"))
    slide_htmls.sort(key=lambda s: s.index)
    (run_dir / "slides_html.json").write_text(
        json.dumps([s.model_dump() for s in slide_htmls], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    deck_path = render_deck_html(run_dir)
    console.print(f"[bold green]完成[/bold green] -> {deck_path}")
    return deck_path
