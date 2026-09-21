/**
 * v1 Extract stage: opens deck.html in headless Chromium, and for every
 * `.slide` walks its DOM to build a Slide Object Model (per spec) — the
 * structured contract the translate stage turns into a native PPTX.
 *
 * Usage: node extract.js <run_dir>   (run_dir must contain deck.html)
 * Writes: <run_dir>/slide_object_model.json, <run_dir>/assets/*.png
 */
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const PX_PER_IN = 96;

// Runs inside the page (must be self-contained: no closures over outer scope).
function extractSlide(slideEl) {
  const PX_PER_IN = 96;
  function toHex2(n) {
    return Math.max(0, Math.min(255, Math.round(n)))
      .toString(16)
      .padStart(2, "0");
  }
  function parseColor(css) {
    if (!css) return null;
    const m = css.match(/rgba?\(([^)]+)\)/);
    if (!m) return null;
    const parts = m[1].split(",").map((s) => parseFloat(s.trim()));
    const [r, g, b, a = 1] = parts;
    if ([r, g, b].some((v) => Number.isNaN(v))) return null;
    return { hex: `${toHex2(r)}${toHex2(g)}${toHex2(b)}`, alpha: a };
  }
  function isVisible(css) {
    const c = parseColor(css);
    return !!c && c.alpha > 0.02;
  }

  const slideRect = slideEl.getBoundingClientRect();
  const objects = [];
  const imageRefs = [];
  let zCounter = 0;
  let extractIdCounter = 0;

  function toRect(rect) {
    return {
      x_in: (rect.left - slideRect.left) / PX_PER_IN,
      y_in: (rect.top - slideRect.top) / PX_PER_IN,
      w_in: rect.width / PX_PER_IN,
      h_in: rect.height / PX_PER_IN,
    };
  }

  function directTextOnlyChildren(el) {
    // true if no child element is block-level (i.e. this element's whole
    // subtree is just inline text formatting, so we treat it as one leaf).
    return Array.from(el.children).every((child) => {
      const d = getComputedStyle(child).display;
      return d === "inline" || d === "inline-block" || d === "inline-flex";
    });
  }

  function markForScreenshot(el, rect, isEditable) {
    const id = `extract-${extractIdCounter++}`;
    el.setAttribute("data-extract-id", id);
    imageRefs.push(id);
    const r = toRect(rect);
    objects.push({
      type: "image",
      ...r,
      text_content: null,
      font_family: null,
      font_size_pt: null,
      font_weight: null,
      color_hex: null,
      fill_hex: null,
      align: null,
      z_index: zCounter++,
      is_editable: isEditable,
      image_ref: id,
    });
  }

  function walk(el) {
    if (el.classList && el.classList.contains("slide-index-badge")) return;
    const style = getComputedStyle(el);
    if (style.display === "none" || style.visibility === "hidden" || parseFloat(style.opacity) === 0) {
      return;
    }
    const rect = el.getBoundingClientRect();
    if (rect.width < 1 || rect.height < 1) {
      Array.from(el.children).forEach(walk);
      return;
    }

    if (el.hasAttribute("data-chart-placeholder")) {
      // Reserved real estate for a native PowerPoint chart (translate.js
      // pairs this geometry with the matching slide's ChartSpec from
      // content.json). Nothing here is DOM content to extract -- the layout
      // agent is told to leave this div empty.
      const r = toRect(rect);
      objects.push({
        type: "chart",
        ...r,
        text_content: null,
        font_family: null,
        font_size_pt: null,
        font_weight: null,
        color_hex: null,
        fill_hex: null,
        align: null,
        z_index: zCounter++,
        is_editable: true,
      });
      return; // don't recurse -- placeholder must be empty
    }

    const isSvg = el.tagName.toLowerCase() === "svg";
    const isImg = el.tagName.toLowerCase() === "img";
    const hasGradient = style.backgroundImage && style.backgroundImage !== "none";
    const hasShadow = style.boxShadow && style.boxShadow !== "none";
    // getBoundingClientRect() is the post-transform axis-aligned bounding box.
    // Treating that AABB as an unrotated shape's x/y/w/h turns e.g. a thin
    // rotated line into a big solid rectangle covering its diagonal extent.
    // Screenshotting instead captures the actual (correctly rotated) pixels.
    const hasTransform = style.transform && style.transform !== "none";

    if (isSvg || hasGradient || hasShadow || hasTransform) {
      markForScreenshot(el, rect, false);
      return; // flattened to an image layer, don't recurse further
    }
    if (isImg) {
      markForScreenshot(el, rect, true);
      return;
    }

    const hasOwnText = directTextOnlyChildren(el) && el.textContent && el.textContent.trim() !== "";
    if (hasOwnText) {
      const r = toRect(rect);
      const color = parseColor(style.color);
      const weight = parseInt(style.fontWeight, 10) || 400;
      objects.push({
        type: "text",
        ...r,
        text_content: el.innerText.trim(),
        font_family: style.fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, ""),
        font_size_pt: Math.round(parseFloat(style.fontSize) * 0.75 * 100) / 100,
        font_weight: weight >= 600 ? "bold" : "normal",
        color_hex: color ? color.hex : "000000",
        fill_hex: null,
        align: ["left", "center", "right", "justify"].includes(style.textAlign) ? style.textAlign : "left",
        z_index: zCounter++,
        is_editable: true,
        // Text always renders fully opaque, even if the source HTML set an
        // opacity < 1: the layout agent is told never to dim readable text
        // (that's what var(--muted) is for), but this is the safety net --
        // low-opacity text is a contrast failure waiting to happen.
      });
      return; // leaf: don't also emit its inline children separately
    }

    const bgVisible = isVisible(style.backgroundColor);
    const borderWidths = [style.borderTopWidth, style.borderRightWidth, style.borderBottomWidth, style.borderLeftWidth].map(parseFloat);
    const hasBorder = borderWidths.some((w) => w > 0) && isVisible(style.borderTopColor || style.borderColor);

    if (bgVisible || hasBorder) {
      const r = toRect(rect);
      const fill = isVisible(style.backgroundColor) ? parseColor(style.backgroundColor) : null;
      objects.push({
        type: "shape",
        ...r,
        text_content: null,
        font_family: null,
        font_size_pt: null,
        font_weight: null,
        color_hex: hasBorder ? (parseColor(style.borderTopColor) || {}).hex || null : null,
        fill_hex: fill ? fill.hex : null,
        align: null,
        z_index: zCounter++,
        is_editable: true,
        border_radius_px: parseFloat(style.borderTopLeftRadius) || 0,
        opacity: parseFloat(style.opacity),
      });
    }

    Array.from(el.children).forEach(walk);
  }

  Array.from(slideEl.children).forEach(walk);

  return {
    width_in: slideRect.width / PX_PER_IN,
    height_in: slideRect.height / PX_PER_IN,
    objects,
    imageRefs,
  };
}

async function main() {
  const runDir = process.argv[2];
  if (!runDir) {
    console.error("Usage: node extract.js <run_dir>");
    process.exit(1);
  }
  const deckHtmlPath = path.resolve(runDir, "deck.html");
  if (!fs.existsSync(deckHtmlPath)) {
    console.error(`Not found: ${deckHtmlPath}`);
    process.exit(1);
  }
  const assetsDir = path.resolve(runDir, "assets");
  fs.mkdirSync(assetsDir, { recursive: true });

  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 }, deviceScaleFactor: 2 });
  await page.goto("file://" + deckHtmlPath.replace(/\\/g, "/"));
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(400); // let web fonts actually repaint/reflow

  const slideHandles = await page.$$(".slide");
  const model = [];

  for (let i = 0; i < slideHandles.length; i++) {
    const slideEl = slideHandles[i];
    const result = await slideEl.evaluate(extractSlide);
    console.log(`slide ${i}: ${result.objects.length} objects (${result.imageRefs.length} image layers)`);

    for (const obj of result.objects) {
      if (obj.type === "image" && obj.image_ref) {
        // Scoped to this slide's subtree: ids only unique per-slide (extractSlide's
        // counter resets each call), a page-wide query could match an identically
        // numbered element from an earlier slide.
        const target = await slideEl.$(`[data-extract-id="${obj.image_ref}"]`);
        const imgPath = path.join(assetsDir, `slide${i}-${obj.image_ref}.png`);
        await target.screenshot({ path: imgPath });
        obj.image_path = path.relative(runDir, imgPath).replace(/\\/g, "/");
        delete obj.image_ref;
      }
    }

    model.push({
      slide_index: i,
      width_in: result.width_in,
      height_in: result.height_in,
      objects: result.objects,
    });
  }

  await browser.close();

  const outPath = path.join(runDir, "slide_object_model.json");
  fs.writeFileSync(outPath, JSON.stringify(model, null, 2), "utf-8");
  console.log(`Wrote ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
