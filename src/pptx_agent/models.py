"""Structured JSON contracts passed between pipeline stages (per spec: agent
handoffs are always structured data, never prose, so each stage is
independently rerunnable)."""

from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class ContentType(str, Enum):
    TITLE = "title"
    BULLETS = "bullets"
    COMPARISON = "comparison"
    DATA = "data"
    TIMELINE = "timeline"
    QUOTE = "quote"


class DataPoint(BaseModel):
    label: str
    value: str
    unit: str | None = None


class ChartType(str, Enum):
    BAR = "bar"
    LINE = "line"
    PIE = "pie"


class ChartSeries(BaseModel):
    name: str
    values: list[float]


class ChartSpec(BaseModel):
    """Only populated when the outline contains real, comparable numeric
    data across categories -- never invented to fill space (same honesty
    rule as data_points' "待填入" placeholders). When present, this becomes
    a native PowerPoint chart (pptxgenjs addChart), not a hand-drawn one."""

    chart_type: ChartType
    categories: list[str]
    series: list[ChartSeries]
    value_label: str | None = Field(default=None, description="e.g. unit or axis label")


class ComparisonSide(BaseModel):
    label: str
    points: list[str]


class TimelineItem(BaseModel):
    label: str
    description: str


class QuoteBlock(BaseModel):
    text: str
    attribution: str | None = None


class SlideContent(BaseModel):
    """One block from the Content Agent. Only one of the type-specific
    fields is populated, matching `content_type`."""

    index: int
    content_type: ContentType
    title: str
    subtitle: str | None = None
    bullets: list[str] | None = None
    data_points: list[DataPoint] | None = None
    chart: ChartSpec | None = None
    comparison: list[ComparisonSide] | None = None
    timeline: list[TimelineItem] | None = None
    quote: QuoteBlock | None = None
    speaker_notes: str | None = None


class Deck(BaseModel):
    deck_title: str
    slides: list[SlideContent]


class DesignTokens(BaseModel):
    """Output of the Design System Agent. HTML Layout Agent must reference
    these values exclusively rather than choosing its own colors."""

    theme_name: str
    primary_color: str = Field(description="6-hex, no #")
    secondary_color: str
    accent_color: str
    background_color: str
    surface_color: str
    text_color: str
    muted_text_color: str
    font_heading: str
    font_body: str
    spacing_unit_px: int
    radius_px: int
    motif: str = Field(description="A short description of one recurring visual motif")


class SlideHtml(BaseModel):
    """Output of the HTML Layout Agent for a single slide."""

    index: int
    layout_variant: str = Field(description="Short label for which layout pattern was used")
    html: str = Field(description="Self-contained inner HTML for the 1280x720px slide canvas")


class SlideQaVerdict(BaseModel):
    """Output of the Visual QA Agent: a vision review of the *actual rendered
    .pptx slide* (not the HTML intermediate), checked against the original
    content and design tokens."""

    index: int
    has_issues: bool
    issues: list[str] = Field(default_factory=list, description="Short, concrete descriptions of each visible problem")
    fix_instructions: str | None = Field(
        default=None, description="Actionable guidance for the HTML Layout Agent to redo this slide"
    )


class SlideObjectType(str, Enum):
    TEXT = "text"
    SHAPE = "shape"
    IMAGE = "image"
    CHART = "chart"
    GROUP = "group"


class SlideObject(BaseModel):
    """The Slide Object Model: the DOM-extracted contract between the HTML
    layer and the PPTX layer (v1). Units are inches unless noted."""

    type: SlideObjectType
    x_in: float
    y_in: float
    w_in: float
    h_in: float
    text_content: str | None = None
    font_family: str | None = None
    font_size_pt: float | None = None
    font_weight: str | None = None
    color_hex: str | None = None
    fill_hex: str | None = None
    align: str | None = None
    z_index: int = 0
    is_editable: bool = True


class SlideObjectModel(BaseModel):
    slide_index: int
    width_in: float
    height_in: float
    objects: list[SlideObject]
