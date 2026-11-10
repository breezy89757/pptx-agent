"""Offline demo (no LLM): hand-written content/tokens/slide HTML that exercises
every chart type and option of the new `data-chart` pipeline.

    python scripts/demo_charts.py outputs/demo-charts
    node node/build.js outputs/demo-charts
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pptx_agent.charts import validate_chart_directives  # noqa: E402
from pptx_agent.models import Deck, DesignTokens, SlideHtml  # noqa: E402
from pptx_agent.pipeline import render_deck_html  # noqa: E402

TOKENS = DesignTokens(
    theme_name="bank-red-neutral",
    primary_color="E60012",
    secondary_color="3A3A3A",
    accent_color="9E9E9E",
    background_color="FFFFFF",
    surface_color="F5F5F5",
    text_color="222222",
    muted_text_color="595959",
    font_heading="Microsoft JhengHei",
    font_body="Microsoft JhengHei",
    spacing_unit_px=16,
    radius_px=10,
    motif="標題左側一條紅色短豎條",
)

DATASETS = [
    {"id": "d1", "title": "台股加權指數收盤(示意)", "columns": ["日期", "收盤指數"],
     "rows": [["11/6", 39701.6], ["11/9", 41661.7]]},
    {"id": "d2", "title": "各季營收組成(億元,示意)", "columns": ["季度", "放款", "手續費", "投資"],
     "rows": [["Q1", 52, 18, 12], ["Q2", 55, 20, 14], ["Q3", 58, 19, 16], ["Q4", 63, 23, 18]]},
    {"id": "d3", "title": "各分行新增戶數(示意)", "columns": ["分行", "新增戶數"],
     "rows": [["台南", 482], ["高雄", 455], ["台中", 391], ["新竹", 322], ["台北", 301], ["桃園", 268]]},
    {"id": "d4", "title": "月成長率(%,示意)", "columns": ["月份", "本行", "同業平均"],
     "rows": [["6月", 2.1, 1.8], ["7月", 2.6, 2.0], ["8月", 3.1, 2.2], ["9月", 2.8, 2.5], ["10月", 3.6, 2.6], ["11月", 4.2, 2.9]]},
    {"id": "d5", "title": "產品占比(示意)", "columns": ["產品", "占比"],
     "rows": [["房貸", 0.42], ["信貸", 0.23], ["企業放款", 0.2], ["財富管理", 0.1], ["其他", 0.05]]},
    {"id": "d6", "title": "方案評分(示意)", "columns": ["構面", "方案 A", "方案 B"],
     "rows": [["成本", 7, 5], ["速度", 6, 8], ["彈性", 8, 6], ["風險控管", 5, 9], ["擴充性", 7, 7]]},
    {"id": "d7", "title": "行銷費用與新增客戶(示意)", "columns": ["行銷費用(萬)", "新增客戶"],
     "rows": [[20, 110], [35, 160], [48, 190], [60, 260], [75, 270], [90, 340], [105, 355], [120, 420]]},
]

DECK = {
    "deck_title": "資料圖表自動生成 Demo",
    "datasets": DATASETS,
    "slides": [
        {"index": 1, "content_type": "data", "title": "11/9 台股反彈近 4.9%", "dataset_ids": ["d1"]},
        {"index": 2, "content_type": "data", "title": "營收組成逐季成長", "dataset_ids": ["d2"]},
        {"index": 3, "content_type": "data", "title": "台南分行新增戶數居冠", "dataset_ids": ["d3"]},
        {"index": 4, "content_type": "data", "title": "成長率領先同業,產品以房貸為主", "dataset_ids": ["d4", "d5"]},
        {"index": 5, "content_type": "data", "title": "方案輪廓與行銷效益", "dataset_ids": ["d6", "d7"]},
        {"index": 6, "content_type": "data", "title": "同一份資料,兩種看法", "dataset_ids": ["d2"]},
    ],
}


def chart(spec: dict, left: int, top: int, w: int, h: int) -> str:
    return (
        f"<div data-chart='{json.dumps(spec, ensure_ascii=False)}' "
        f'style="position:absolute;left:{left}px;top:{top}px;width:{w}px;height:{h}px;"></div>'
    )


def frame(title: str, subtitle: str, body: str) -> str:
    return f"""<div style="position:absolute;inset:0;width:1280px;height:720px;overflow:hidden;background:var(--bg);font-family:var(--font-body);color:var(--text);">
  <div style="position:absolute;left:64px;top:44px;width:8px;height:44px;background:var(--primary);"></div>
  <div style="position:absolute;left:88px;top:38px;width:1100px;font-family:var(--font-heading);font-size:36px;font-weight:700;line-height:1.25;color:var(--text);">{title}</div>
  <div style="position:absolute;left:88px;top:90px;width:1100px;font-size:18px;color:var(--muted);">{subtitle}</div>
  {body}
</div>"""


def card(left, top, w, h, big, label, note="") -> str:
    return f"""<div style="position:absolute;left:{left}px;top:{top}px;width:{w}px;height:{h}px;background:var(--surface);border-radius:var(--radius);">
    <div style="position:absolute;left:24px;top:20px;width:{w-48}px;font-size:16px;color:var(--muted);">{label}</div>
    <div style="position:absolute;left:24px;top:52px;width:{w-48}px;font-family:var(--font-heading);font-size:44px;font-weight:700;color:var(--primary);">{big}</div>
    <div style="position:absolute;left:24px;top:{h-56}px;width:{w-48}px;font-size:15px;line-height:1.5;color:var(--text);">{note}</div>
  </div>"""


SLIDES_HTML = {
    1: frame(
        "11/9 台股反彈近 4.9%", "加權指數收盤:11/6 → 11/9(縱軸不從 0 起算,只強調 11/9)",
        chart({"type": "bar", "dataset": "d1", "x": "日期", "y": ["收盤指數"], "title": "加權指數收盤",
               "options": {"highlight": "11/9", "y_min": 38000, "y_max": 42500, "gap_width": 80}}, 64, 150, 720, 520)
        + card(832, 150, 384, 250, "+1,960 點", "11/9 較 11/6", "政策與海外科技股反彈帶動,指數回升至 41,661.7。")
        + card(832, 424, 384, 246, "-4.7%", "11/6 單日跌幅", "題材為外資賣壓與市場恐慌,台股一度重挫。"),
    ),
    2: frame(
        "營收組成逐季成長", "各季營收組成(億元,示意數據),堆疊長條",
        chart({"type": "bar", "dataset": "d2", "x": "季度", "y": ["放款", "手續費", "投資"],
               "options": {"stacked": "stacked", "legend": "bottom", "y_title": "億元", "gap_width": 55,
                           "colors": ["primary", "secondary", "accent"]}}, 64, 150, 760, 520)
        + card(868, 150, 348, 250, "+21%", "Q4 對 Q1 總營收", "放款仍是主體,手續費與投資同步放大。")
        + card(868, 424, 348, 246, "63 億", "Q4 放款收入", "連續四季成長,為各項收入最高。"),
    ),
    3: frame(
        "台南分行新增戶數居冠", "各分行新增戶數(示意數據),橫向長條並強調單一類別",
        chart({"type": "bar", "dataset": "d3", "x": "分行", "y": ["新增戶數"],
               "options": {"horizontal": True, "highlight": "台南", "gap_width": 40, "x_title": ""}}, 64, 150, 800, 520)
        + card(908, 150, 308, 250, "482", "台南 新增戶數", "高於第二名高雄 27 戶。")
        + card(908, 424, 308, 246, "268", "桃園 新增戶數", "後段分行仍有拉升空間。"),
    ),
    4: frame(
        "成長率領先同業,產品以房貸為主", "左:折線(本行 vs 同業平均);右:環圈(產品占比)",
        chart({"type": "line", "dataset": "d4", "x": "月份", "y": ["本行", "同業平均"],
               "title": "月成長率(%)", "options": {"colors": ["primary", "accent"], "number_format": "0.0", "y_min": 0}}, 64, 150, 660, 520)
        + chart({"type": "doughnut", "dataset": "d5", "x": "產品", "y": ["占比"], "title": "產品占比",
                 "options": {"legend": "bottom", "colors": ["primary", "secondary", "accent", "muted", "text"]}}, 760, 150, 456, 520),
    ),
    5: frame(
        "方案輪廓與行銷效益", "左:雷達圖(兩方案比較);右:散佈圖(費用 vs 新增客戶)",
        chart({"type": "radar", "dataset": "d6", "x": "構面", "y": ["方案 A", "方案 B"], "title": "方案評分",
               "options": {"colors": ["primary", "secondary"], "legend": "bottom"}}, 64, 150, 560, 520)
        + chart({"type": "scatter", "dataset": "d7", "x": "行銷費用(萬)", "y": ["新增客戶"], "title": "費用與新增客戶",
                 "options": {"colors": ["primary"], "legend": "none", "x_title": "行銷費用(萬)", "y_title": "新增客戶"}}, 660, 150, 556, 520),
    ),
    6: frame(
        "同一份資料,兩種看法", "左:面積(絕對量堆疊);右:100% 堆疊長條(結構占比)",
        chart({"type": "area", "dataset": "d2", "x": "季度", "y": ["放款", "手續費", "投資"], "title": "絕對量(億元)",
               "options": {"stacked": "stacked", "legend": "bottom", "colors": ["primary", "secondary", "accent"], "data_labels": False}}, 64, 150, 560, 520)
        + chart({"type": "bar", "dataset": "d2", "x": "季度", "y": ["放款", "手續費", "投資"], "title": "結構占比",
                 "options": {"stacked": "percent", "legend": "bottom", "colors": ["primary", "secondary", "accent"], "number_format": "0%"}}, 660, 150, 556, 520),
    ),
}


def main(out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    deck = Deck.model_validate(DECK)
    (out / "content.json").write_text(deck.model_dump_json(indent=2), encoding="utf-8")
    (out / "design_tokens.json").write_text(TOKENS.model_dump_json(indent=2), encoding="utf-8")
    slides = []
    for s in deck.slides:
        html = SLIDES_HTML[s.index]
        ds = [d for d in deck.datasets if d.id in (s.dataset_ids or [])]
        errors = validate_chart_directives(html, ds)
        assert not errors, f"slide {s.index}: {errors}"
        slides.append(SlideHtml(index=s.index, layout_variant="demo", html=html).model_dump())
    (out / "slides_html.json").write_text(json.dumps(slides, ensure_ascii=False, indent=2), encoding="utf-8")
    render_deck_html(out)
    print(f"wrote demo run to {out}")


if __name__ == "__main__":
    main(Path(sys.argv[1] if len(sys.argv) > 1 else "outputs/demo-charts"))
