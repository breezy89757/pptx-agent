function toHex2(n) {
  return Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0");
}

/** Parses "rgb(r,g,b)" / "rgba(r,g,b,a)" into { hex, alpha }. Returns null on parse failure. */
function parseColor(cssColor) {
  if (!cssColor) return null;
  const m = cssColor.match(/rgba?\(([^)]+)\)/);
  if (!m) return null;
  const parts = m[1].split(",").map((s) => parseFloat(s.trim()));
  const [r, g, b, a = 1] = parts;
  if ([r, g, b].some((v) => Number.isNaN(v))) return null;
  return { hex: `${toHex2(r)}${toHex2(g)}${toHex2(b)}`, alpha: a };
}

function isVisibleColor(cssColor) {
  const parsed = parseColor(cssColor);
  return !!parsed && parsed.alpha > 0.02;
}

module.exports = { parseColor, isVisibleColor };
