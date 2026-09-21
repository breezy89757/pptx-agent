/**
 * v1 Translate stage: Slide Object Model (JSON) -> native, editable .pptx
 * via pptxgenjs. type=text -> addText, type=shape -> addShape,
 * type=image -> addImage (flattened decorative layers from extract.js).
 *
 * Usage: node translate.js <run_dir>
 * Reads: <run_dir>/slide_object_model.json, <run_dir>/design_tokens.json, <run_dir>/content.json
 * Writes: <run_dir>/deck.pptx
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const CANVAS_W_IN = 1280 / 96;
const CANVAS_H_IN = 720 / 96;
// Extra width kept for text boxes: browser vs. PowerPoint font metrics differ
// enough (esp. for CJK) that a tight box risks clipped/wrapped text (see spec's
// "字型度量差異緩衝" note).
const TEXT_WIDTH_SLACK = 1.12;

function widenText(obj) {
  const newW = Math.min(obj.w_in * TEXT_WIDTH_SLACK, CANVAS_W_IN - obj.x_in);
  return { x: obj.x_in, w: Math.max(newW, obj.w_in) };
}

// pptxgenjs transparency is 0 (opaque) - 100 (fully transparent); CSS opacity is the inverse.
function transparencyOf(obj) {
  const opacity = typeof obj.opacity === "number" && !Number.isNaN(obj.opacity) ? obj.opacity : 1;
  return Math.round((1 - Math.min(Math.max(opacity, 0), 1)) * 100);
}

function addTextObject(slide, obj) {
  const { x, w } = widenText(obj);
  const transparency = transparencyOf(obj);
  slide.addText(obj.text_content || "", {
    x,
    y: obj.y_in,
    w,
    h: obj.h_in,
    fontFace: obj.font_family || "Arial",
    fontSize: obj.font_size_pt || 12,
    bold: obj.font_weight === "bold",
    color: obj.color_hex || "000000",
    ...(transparency > 0 ? { transparency } : {}),
    align: obj.align || "left",
    valign: "top",
    wrap: true,
    margin: 0,
    autoFit: false,
  });
}

function addShapeObject(slide, pptx, obj) {
  const shapeType = obj.border_radius_px > 4 ? pptx.ShapeType.roundRect : pptx.ShapeType.rect;
  const transparency = transparencyOf(obj);
  const opts = {
    x: obj.x_in,
    y: obj.y_in,
    w: obj.w_in,
    h: obj.h_in,
    fill: obj.fill_hex ? { color: obj.fill_hex, transparency } : { type: "none" },
    line: obj.color_hex ? { color: obj.color_hex, width: 1, transparency } : { type: "none" },
  };
  if (shapeType === pptx.ShapeType.roundRect) {
    opts.rectRadius = Math.min(obj.border_radius_px / 96, Math.min(obj.w_in, obj.h_in) / 2);
  }
  slide.addShape(shapeType, opts);
}

function addImageObject(slide, runDir, obj) {
  if (!obj.image_path) return;
  slide.addImage({
    path: path.join(runDir, obj.image_path),
    x: obj.x_in,
    y: obj.y_in,
    w: obj.w_in,
    h: obj.h_in,
  });
}

const CHART_TYPE_MAP = { bar: "bar", line: "line", pie: "pie" };

function addChartObject(slide, pptx, obj, chartSpec, tokens) {
  if (!chartSpec || !chartSpec.series || chartSpec.series.length === 0) return;
  const chartType = pptx.ChartType[CHART_TYPE_MAP[chartSpec.chart_type] || "bar"];
  const chartColors = [tokens.primary_color, tokens.secondary_color, tokens.accent_color];

  const chartData =
    chartSpec.chart_type === "pie"
      ? [{ name: chartSpec.value_label || chartSpec.series[0].name, labels: chartSpec.categories, values: chartSpec.series[0].values }]
      : chartSpec.series.map((s) => ({ name: s.name, labels: chartSpec.categories, values: s.values }));

  slide.addChart(chartType, chartData, {
    x: obj.x_in,
    y: obj.y_in,
    w: obj.w_in,
    h: obj.h_in,
    chartColors,
    showLegend: chartData.length > 1,
    legendColor: tokens.muted_text_color,
    legendFontFace: tokens.font_body,
    catAxisLabelColor: tokens.muted_text_color,
    catAxisLabelFontFace: tokens.font_body,
    valAxisLabelColor: tokens.muted_text_color,
    valAxisLabelFontFace: tokens.font_body,
    dataLabelColor: tokens.text_color,
    dataLabelFontFace: tokens.font_body,
    showTitle: false,
    catAxisLineColor: tokens.surface_color,
    valAxisLineColor: tokens.surface_color,
    showValue: chartSpec.chart_type !== "line",
  });
}

async function main() {
  const runDir = process.argv[2];
  if (!runDir) {
    console.error("Usage: node translate.js <run_dir>");
    process.exit(1);
  }
  const modelPath = path.resolve(runDir, "slide_object_model.json");
  const tokensPath = path.resolve(runDir, "design_tokens.json");
  const contentPath = path.resolve(runDir, "content.json");
  for (const p of [modelPath, tokensPath]) {
    if (!fs.existsSync(p)) {
      console.error(`Not found: ${p} (run extract.js first)`);
      process.exit(1);
    }
  }

  const model = JSON.parse(fs.readFileSync(modelPath, "utf-8"));
  const tokens = JSON.parse(fs.readFileSync(tokensPath, "utf-8"));
  const content = fs.existsSync(contentPath) ? JSON.parse(fs.readFileSync(contentPath, "utf-8")) : null;
  const deckTitle = content ? content.deck_title : "Deck";
  // slide_object_model.json is built from deck.html's slide order, which
  // matches content.json's slides sorted by their own `index` -- pair them
  // positionally so each chart placeholder finds its ChartSpec.
  const contentSlidesInOrder = content ? [...content.slides].sort((a, b) => a.index - b.index) : [];

  const pptx = new pptxgen();
  pptx.defineLayout({ name: "PPTX_AGENT_WIDE", width: CANVAS_W_IN, height: CANVAS_H_IN });
  pptx.layout = "PPTX_AGENT_WIDE";
  pptx.title = deckTitle;

  model.forEach((slideModel, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: tokens.background_color };
    const chartSpec = contentSlidesInOrder[i] ? contentSlidesInOrder[i].chart : null;
    const objects = [...slideModel.objects].sort((a, b) => a.z_index - b.z_index);
    for (const obj of objects) {
      if (obj.type === "text") addTextObject(slide, obj);
      else if (obj.type === "shape") addShapeObject(slide, pptx, obj);
      else if (obj.type === "image") addImageObject(slide, runDir, obj);
      else if (obj.type === "chart") addChartObject(slide, pptx, obj, chartSpec, tokens);
    }
  });

  const outPath = path.join(runDir, "deck.pptx");
  await pptx.writeFile({ fileName: outPath });
  console.log(`Wrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
