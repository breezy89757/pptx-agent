"""Chart directives: the declarative spec the HTML Layout Agent writes into
`data-chart='{...}'` attributes, plus validation against the deck's datasets.

Design rule: the Layout Agent never retypes numbers. It references a dataset
by id and picks columns; the Node translate stage resolves the real rows and
builds a native, data-editable PowerPoint chart."""

from __future__ import annotations

import json
from enum import Enum
from html.parser import HTMLParser
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, ValidationError

from pptx_agent.models import Dataset

ColorRole = Literal["primary", "secondary", "accent", "muted", "text"]


class ChartKind(str, Enum):
    BAR = "bar"
    LINE = "line"
    AREA = "area"
    PIE = "pie"
    DOUGHNUT = "doughnut"
    SCATTER = "scatter"
    RADAR = "radar"


class ChartOptions(BaseModel):
    model_config = ConfigDict(extra="forbid")

    horizontal: bool = False
    stacked: Literal["none", "stacked", "percent"] = "none"
    data_labels: bool = True
    legend: Literal["none", "top", "bottom", "right"] | None = None  # None = auto
    x_title: str | None = None
    y_title: str | None = None
    y_min: float | None = None
    y_max: float | None = None
    smooth: bool = False
    number_format: str | None = Field(default=None, description='e.g. "#,##0" or "0.0%"')
    gap_width: int | None = Field(default=None, ge=0, le=500)
    colors: list[ColorRole] | None = Field(default=None, description="series (or slice) color roles in order")
    highlight: str | None = Field(
        default=None, description="single-series bar only: category label drawn in the first color role, others neutral"
    )


class ChartDirective(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: ChartKind
    dataset: str
    x: str = Field(description="category column (scatter: numeric x column)")
    y: list[str] = Field(min_length=1, description="one or more numeric series columns")
    title: str | None = None
    options: ChartOptions = Field(default_factory=ChartOptions)


class _ChartScanner(HTMLParser):
    """Collects every element carrying data-chart and notes whether it has
    child elements or text (a placeholder must be empty: its contents are
    discarded by the extract stage)."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.found: list[dict] = []
        self._stack: list[dict | None] = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        # any element nested inside an open chart placeholder is a child
        for open_chart in (e for e in self._stack if e):
            open_chart["has_children"] = True
        entry = None
        if "data-chart" in a:
            entry = {"raw": a["data-chart"], "style": a.get("style", ""), "has_children": False}
            self.found.append(entry)
        if tag not in {"br", "img", "hr", "input", "meta", "link"}:
            self._stack.append(entry)

    def handle_startendtag(self, tag, attrs):
        a = dict(attrs)
        for open_chart in (e for e in self._stack if e):
            open_chart["has_children"] = True
        if "data-chart" in a:
            self.found.append({"raw": a["data-chart"], "style": a.get("style", ""), "has_children": False})

    def handle_endtag(self, tag):
        if self._stack:
            self._stack.pop()

    def handle_data(self, data):
        if data.strip():
            for open_chart in (e for e in self._stack if e):
                open_chart["has_children"] = True


def extract_chart_directives(html: str) -> list[dict]:
    scanner = _ChartScanner()
    scanner.feed(html)
    return scanner.found


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def validate_chart_directives(html: str, datasets: list[Dataset]) -> list[str]:
    """Return human-readable (Traditional Chinese) errors; empty list = OK.
    The messages are fed straight back to the Layout Agent on retry."""
    errors: list[str] = []
    by_id = {d.id: d for d in datasets}

    for n, found in enumerate(extract_chart_directives(html), start=1):
        tag = f"第 {n} 個 data-chart"
        try:
            raw = json.loads(found["raw"])
        except json.JSONDecodeError as e:
            errors.append(f"{tag}:不是合法 JSON({e.msg},位置 {e.pos})。屬性要用單引號包住,JSON 內部用雙引號。")
            continue
        try:
            d = ChartDirective.model_validate(raw)
        except ValidationError as e:
            for err in e.errors():
                loc = ".".join(str(p) for p in err["loc"])
                errors.append(f"{tag}:欄位 {loc or '(根)'} 不合法 — {err['msg']}")
            continue

        if found["has_children"]:
            errors.append(f"{tag}:圖表版位內不可放任何子元素或文字,請保持空白 div。")
        style = found["style"].replace(" ", "").lower()
        if "position:absolute" not in style or "width:" not in style or "height:" not in style:
            errors.append(f"{tag}:style 必須包含 position:absolute 與明確的 left/top/width/height(px)。")

        ds = by_id.get(d.dataset)
        if ds is None:
            errors.append(f"{tag}:找不到資料集 {d.dataset!r},可用的有 {sorted(by_id) or '(無)'}。")
            continue
        cols = {c: i for i, c in enumerate(ds.columns)}
        for col in [d.x, *d.y]:
            if col not in cols:
                errors.append(f"{tag}:資料集 {ds.id} 沒有欄位 {col!r},可用欄位:{ds.columns}。")
        if any(col not in cols for col in [d.x, *d.y]):
            continue
        for col in d.y:
            bad = [r[cols[col]] for r in ds.rows if not _is_number(r[cols[col]])]
            if bad:
                errors.append(f"{tag}:欄位 {col!r} 含非數值資料 {bad[:3]},不能當數列。")
        if d.type in (ChartKind.PIE, ChartKind.DOUGHNUT) and len(d.y) != 1:
            errors.append(f"{tag}:{d.type.value} 圖只能有 1 個數列(y 只放一個欄位)。")
        if d.type == ChartKind.SCATTER:
            xs = [r[cols[d.x]] for r in ds.rows]
            if not all(_is_number(v) for v in xs):
                errors.append(f"{tag}:scatter 的 x 欄位 {d.x!r} 必須是數值。")
        o = d.options
        if o.highlight is not None:
            cats = [str(r[cols[d.x]]) for r in ds.rows]
            if d.type != ChartKind.BAR or len(d.y) != 1 or o.stacked != "none":
                errors.append(f"{tag}:highlight 只能用在單一數列、非堆疊的 bar 圖。")
            elif o.highlight not in cats:
                errors.append(f"{tag}:highlight={o.highlight!r} 不在類別 {cats} 之中。")
        if o.y_min is not None and o.y_max is not None and o.y_min >= o.y_max:
            errors.append(f"{tag}:y_min 必須小於 y_max。")
        if o.stacked != "none" and d.type not in (ChartKind.BAR, ChartKind.AREA):
            errors.append(f"{tag}:stacked 只適用 bar / area。")
    return errors


# Prompt fragment shared with the layout agent so the grammar lives in one place.
GRAMMAR_PROMPT = """\
data-chart 語法(原生、可在 PowerPoint 內編輯資料的圖表):
`<div data-chart='{JSON}' style="position:absolute;left:..px;top:..px;width:..px;height:..px;"></div>`
- 屬性值用單引號包住,JSON 內部一律雙引號;div 內必須是空的(內容會被丟棄)。
- JSON 欄位:
  type: "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "radar"
  dataset: 資料集 id(只能用本頁提供的資料集)
  x: 類別欄位名稱(scatter 則為數值 x 欄位)
  y: 數列欄位名稱陣列(pie/doughnut 只能一個)
  title: 圖表標題(選填)
  options(皆選填):
    horizontal(bar 橫向)、stacked("none"|"stacked"|"percent")、data_labels(預設 true)、
    legend("none"|"top"|"bottom"|"right",預設自動)、x_title、y_title、y_min、y_max、smooth(line)、
    number_format(例如 "#,##0"、"0.0%")、gap_width(0-500,bar 間距)、
    colors(色彩角色陣列,值為 "primary"|"secondary"|"accent"|"muted"|"text",依數列/扇區順序)、
    highlight(單一數列 bar 專用:要強調的類別名稱,該柱用第一個色彩角色,其餘柱為中性灰)
- 你不需要也不可以重打數字:數值由系統依 dataset 取出。顏色由設計令牌決定,只能用 colors 指定「角色」。
- 一頁可放多張圖表(各自一個 div)。挑最能傳達訊息的圖表類型:比較用 bar、趨勢用 line/area、占比用 pie/doughnut、
  組成隨時間變化用 stacked bar、多維度輪廓用 radar、兩個數值關係用 scatter。
- 想突出「某一個值」時優先用 highlight,不要把每根柱都塗成強調色。
"""
