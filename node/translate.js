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
const { addChartObject, ChartError } = require("./lib/charts");

const CANVAS_W_IN = 1280 / 96;
const CANVAS_H_IN = 720 / 96;

// pptxgenjs transparency is 0 (opaque) - 100 (fully transparent); CSS opacity is the inverse.
function transparencyOf(obj) {
  const opacity = typeof obj.opacity === "number" && !Number.isNaN(obj.opacity) ? obj.opacity : 1;
  return Math.round((1 - Math.min(Math.max(opacity, 0), 1)) * 100);
}

function addTextObject(slide, obj) {
  const transparency = transparencyOf(obj);
  slide.addText(obj.text_content || "", {
    x: obj.x_in,
    y: obj.y_in,
    w: obj.w_in,
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
  // Charts reference these by id (data-chart directive); rows are never
  // retyped by the layout agent.
  const datasets = content ? content.datasets || [] : [];
  const chartErrors = [];

  const pptx = new pptxgen();
  pptx.defineLayout({ name: "PPTX_AGENT_WIDE", width: CANVAS_W_IN, height: CANVAS_H_IN });
  pptx.layout = "PPTX_AGENT_WIDE";
  pptx.title = deckTitle;

  model.forEach((slideModel, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: tokens.background_color };
    const objects = [...slideModel.objects].sort((a, b) => a.z_index - b.z_index);
    for (const obj of objects) {
      if (obj.type === "text") addTextObject(slide, obj);
      else if (obj.type === "shape") addShapeObject(slide, pptx, obj);
      else if (obj.type === "image") addImageObject(slide, runDir, obj);
      else if (obj.type === "chart") {
        try {
          addChartObject(slide, pptx, obj, datasets, tokens);
        } catch (err) {
          if (!(err instanceof ChartError)) throw err;
          chartErrors.push(`slide ${i}: ${err.message}`);
        }
      }
    }
  });

  if (chartErrors.length) {
    // fail loudly: a dropped chart should never ship silently
    console.error("Chart build errors:\n  " + chartErrors.join("\n  "));
    process.exit(2);
  }

  const outPath = path.join(runDir, "deck.pptx");
  await pptx.writeFile({ fileName: outPath });
  console.log(`Wrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
