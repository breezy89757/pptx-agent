from __future__ import annotations

import json
import subprocess
import webbrowser
from pathlib import Path
from typing import Annotated

import typer
from rich.console import Console
from rich.markup import escape

from pptx_agent.agents.layout_agent import run_layout_agent
from pptx_agent.agents.qa_agent import run_qa_agent
from pptx_agent.models import Deck, DesignTokens
from pptx_agent.pipeline import render_deck_html, run_pipeline
from pptx_agent.validate import validate_pptx

app = typer.Typer(add_completion=False)
console = Console()

PROJECT_ROOT = Path(__file__).resolve().parents[2]
NODE_DIR = PROJECT_ROOT / "node"


def _run_node(script: str, run_dir: Path) -> None:
    result = subprocess.run(["node", str(NODE_DIR / script), str(run_dir.resolve())], cwd=NODE_DIR)
    if result.returncode != 0:
        raise typer.Exit(result.returncode)


def _run_qa_loop(run_dir: Path, qa_rounds: int) -> None:
    if qa_rounds <= 0:
        return

    try:
        from pptx_agent.pptx_render import render_pptx_to_pngs
    except ImportError:
        console.print("[yellow]略過視覺 QA:此環境沒有 pywin32(僅支援 Windows + PowerPoint)。[/yellow]")
        return

    deck = Deck.model_validate_json((run_dir / "content.json").read_text(encoding="utf-8"))
    tokens = DesignTokens.model_validate_json((run_dir / "design_tokens.json").read_text(encoding="utf-8"))
    slide_by_index = {s.index: s for s in deck.slides}
    sorted_slides = sorted(deck.slides, key=lambda s: s.index)

    for round_num in range(1, qa_rounds + 1):
        console.print(f"[bold cyan]視覺 QA 第 {round_num}/{qa_rounds} 輪[/bold cyan] 渲染實際 PowerPoint 畫面並檢查...")
        try:
            pngs = render_pptx_to_pngs(run_dir / "deck.pptx", run_dir / f"qa_round_{round_num}")
        except Exception as exc:  # COM automation is inherently flaky
            console.print(f"[yellow]視覺 QA 渲染失敗,略過本輪:{exc}[/yellow]")
            return

        if len(pngs) != len(sorted_slides):
            console.print(
                f"[yellow]警告:渲染出 {len(pngs)} 張圖片,但有 {len(sorted_slides)} 張投影片,略過本輪 QA。[/yellow]"
            )
            return

        flagged = []
        for png, slide in zip(pngs, sorted_slides):
            verdict = run_qa_agent(png, slide, tokens)
            if verdict.has_issues:
                flagged.append(verdict)
                console.print(escape(f"  slide #{slide.index}: [問題] " + "; ".join(verdict.issues)))
            else:
                console.print(f"  slide #{slide.index}: [green]OK[/green]")

        if not flagged:
            console.print("[bold green]視覺 QA 通過,沒有發現問題[/bold green]")
            return

        console.print(f"[bold yellow]{len(flagged)} 張投影片有問題,重新產生版面中...[/bold yellow]")
        slides_html_path = run_dir / "slides_html.json"
        slide_htmls_by_index = {s["index"]: s for s in json.loads(slides_html_path.read_text(encoding="utf-8"))}
        for verdict in flagged:
            slide = slide_by_index[verdict.index]
            new_html = run_layout_agent(slide, tokens, extra_instructions=verdict.fix_instructions)
            slide_htmls_by_index[verdict.index] = new_html.model_dump()
        slides_html_path.write_text(
            json.dumps(list(slide_htmls_by_index.values()), ensure_ascii=False, indent=2), encoding="utf-8"
        )

        render_deck_html(run_dir)
        _run_node("extract.js", run_dir)
        _run_node("translate.js", run_dir)
    else:
        console.print(f"[bold yellow]已達最大 QA 輪數({qa_rounds}),仍有未解決的問題,建議人工檢查。[/bold yellow]")


@app.command()
def generate(
    outline: Annotated[
        str | None, typer.Option(help="簡報大綱文字(與 --outline-file 擇一)")
    ] = None,
    outline_file: Annotated[
        Path | None, typer.Option(help="包含大綱的文字檔路徑")
    ] = None,
    output_dir: Annotated[Path, typer.Option(help="輸出根目錄")] = Path("outputs"),
    slides: Annotated[int | None, typer.Option(help="期望的投影片數量(可略)")] = None,
    html_only: Annotated[
        bool, typer.Option(help="只做到 v0 網頁簡報,不轉成 .pptx")
    ] = False,
    qa_rounds: Annotated[
        int, typer.Option(help="視覺 QA 迴圈輪數:實際渲染 .pptx、讓模型檢查、有問題就重做該頁。0 表示關閉")
    ] = 1,
    open_result: Annotated[
        bool, typer.Option(help="完成後自動開啟結果(.pptx 或 deck.html)")
    ] = True,
) -> None:
    """大綱 -> Content/Design/Layout Agent -> deck.html -> 幾何萃取 + 轉成可編輯 .pptx -> 視覺 QA 迴圈 -> 驗證。"""
    if not outline and not outline_file:
        typer.echo("請提供 --outline 或 --outline-file", err=True)
        raise typer.Exit(1)
    text = outline_file.read_text(encoding="utf-8") if outline_file else outline
    assert text is not None

    deck_path = run_pipeline(text, output_root=output_dir, target_slide_count=slides)
    run_dir = deck_path.parent

    if html_only:
        if open_result:
            webbrowser.open(deck_path.resolve().as_uri())
        return

    console.print("[bold cyan]4/5 Extract[/bold cyan] Playwright 萃取版面幾何...")
    _run_node("extract.js", run_dir)

    console.print("[bold cyan]5/5 Translate[/bold cyan] pptxgenjs 轉成 .pptx...")
    _run_node("translate.js", run_dir)

    _run_qa_loop(run_dir, qa_rounds)

    console.print("[bold cyan]驗證[/bold cyan] 檢查文字可編輯 + 每頁有非純文字元素...")
    ok, reports = validate_pptx(run_dir)
    for r in reports:
        status = "OK" if r.ok else "FAIL"
        label = f"slide {r.index}" if r.index >= 0 else "deck"
        console.print(f"  [{'green' if r.ok else 'red'}][{status}][/] {label}: text={r.text_shape_count} shapes={r.non_text_shape_count} pictures={r.picture_count} charts={r.chart_count}")
        for p in r.problems:
            console.print(f"        - {p}")

    pptx_path = run_dir / "deck.pptx"
    if ok:
        console.print(f"[bold green]完成,驗證通過[/bold green] -> {pptx_path}")
    else:
        console.print(f"[bold yellow]完成,但驗證有警告[/bold yellow] -> {pptx_path}")

    if open_result:
        import os

        os.startfile(pptx_path)  # noqa: S606 (Windows-only convenience open)


if __name__ == "__main__":
    app()
