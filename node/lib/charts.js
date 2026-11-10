/**
 * Native chart builder: resolves a `data-chart` directive (written by the
 * HTML Layout Agent) against content.json's datasets and emits a real,
 * data-editable PowerPoint chart through pptxgenjs.addChart.
 *
 * The directive never carries numbers -- rows come from the dataset, so the
 * embedded workbook holds the true values.
 */

const hex = (h) => String(h || "").replace("#", "").toUpperCase();

function blend(fg, bg, t) {
  // t = share of fg (1 = fg, 0 = bg)
  const a = hex(fg);
  const b = hex(bg);
  const ch = (s, i) => parseInt(s.slice(i, i + 2), 16);
  const mix = (i) =>
    Math.round(ch(a, i) * t + ch(b, i) * (1 - t))
      .toString(16)
      .padStart(2, "0");
  return `${mix(0)}${mix(2)}${mix(4)}`.toUpperCase();
}

function roleColor(role, tokens) {
  const map = {
    primary: tokens.primary_color,
    secondary: tokens.secondary_color,
    accent: tokens.accent_color,
    muted: tokens.muted_text_color,
    text: tokens.text_color,
  };
  return hex(map[role] || tokens.primary_color);
}

/** n colors: the given roles first, then progressively lighter tints of them. */
function palette(n, roles, tokens) {
  const base = (roles && roles.length ? roles : ["primary", "secondary", "accent", "muted"]).map((r) =>
    roleColor(r, tokens)
  );
  const out = [];
  for (let i = 0; i < n; i++) {
    if (i < base.length) out.push(base[i]);
    else {
      const round = Math.floor(i / base.length);
      out.push(blend(base[i % base.length], tokens.background_color, Math.max(0.2, 0.65 - 0.2 * (round - 1))));
    }
  }
  return out;
}

function luminance(h) {
  const c = hex(h);
  const v = [0, 2, 4].map((i) => {
    const s = parseInt(c.slice(i, i + 2), 16) / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
}

const readableOn = (bgHex) => (luminance(bgHex) > 0.4 ? "222222" : "FFFFFF");

class ChartError extends Error {}

function resolveData(spec, datasets) {
  const ds = (datasets || []).find((d) => d.id === spec.dataset);
  if (!ds) throw new ChartError(`dataset ${JSON.stringify(spec.dataset)} not found`);
  const idx = (col) => {
    const i = ds.columns.indexOf(col);
    if (i < 0) throw new ChartError(`dataset ${ds.id} has no column ${JSON.stringify(col)}`);
    return i;
  };
  const xi = idx(spec.x);
  const yis = (spec.y || []).map(idx);
  if (yis.length === 0) throw new ChartError("chart has no y series");
  const categories = ds.rows.map((r) => String(r[xi]));
  const xs = ds.rows.map((r) => Number(r[xi]));
  const series = yis.map((yi, k) => ({
    name: spec.y[k],
    values: ds.rows.map((r) => {
      const n = Number(r[yi]);
      if (Number.isNaN(n)) throw new ChartError(`non-numeric value in ${ds.id}.${spec.y[k]}`);
      return n;
    }),
  }));
  return { categories, xs, series };
}

const isInt = (v) => Number.isInteger(v);

function addChartObject(slide, pptx, obj, datasets, tokens) {
  if (obj.chart_error) throw new ChartError(`data-chart is not valid JSON: ${obj.chart_error}`);
  const spec = obj.chart_spec;
  if (!spec) throw new ChartError("data-chart has no spec");
  const o = spec.options || {};
  const { categories, xs, series } = resolveData(spec, datasets);
  const type = spec.type;
  const bg = hex(tokens.background_color);
  const font = tokens.font_body;
  const muted = hex(tokens.muted_text_color);
  const text = hex(tokens.text_color);

  const allValues = series.flatMap((s) => s.values);
  const integers = allValues.every(isInt);
  const defaultFmt = integers ? "#,##0" : "#,##0.0";
  const percentStacked = type === "bar" && o.stacked === "percent";
  // For a 100%-stacked bar the axis is a share (0%-100%) while labels keep the
  // raw values; a "%" format given by the model would otherwise turn 12 into 1200%.
  const labelFmt = percentStacked && (o.number_format || "").includes("%") ? defaultFmt : o.number_format || defaultFmt;
  const axisFmt = percentStacked ? "0%" : o.number_format || defaultFmt;

  const perPoint = type === "pie" || type === "doughnut";
  const highlightMode = o.highlight != null && type === "bar" && series.length === 1;
  let colors;
  if (highlightMode) {
    const hi = palette(1, o.colors, tokens)[0];
    const neutral = blend(muted, bg, 0.45);
    colors = categories.map((c) => (c === o.highlight ? hi : neutral));
  } else {
    colors = palette(perPoint ? categories.length : series.length, o.colors, tokens);
  }

  const common = {
    x: obj.x_in,
    y: obj.y_in,
    w: obj.w_in,
    h: obj.h_in,
    chartColors: colors,
    // text
    catAxisLabelColor: muted,
    catAxisLabelFontFace: font,
    catAxisLabelFontSize: 12,
    valAxisLabelColor: muted,
    valAxisLabelFontFace: font,
    valAxisLabelFontSize: 11,
    dataLabelColor: text,
    dataLabelFontFace: font,
    dataLabelFontSize: 12,
    dataLabelFormatCode: labelFmt,
    valAxisLabelFormatCode: axisFmt,
    showValue: o.data_labels !== false,
    // chrome
    valGridLine: { color: blend(muted, bg, 0.18), size: 0.5 },
    catGridLine: { style: "none" },
    catAxisLineShow: true,
    valAxisLineShow: false,
    showTitle: !!spec.title,
    ...(spec.title ? { title: spec.title, titleFontFace: tokens.font_heading, titleColor: text, titleFontSize: 16 } : {}),
  };

  // legend
  const legendMap = { top: "t", bottom: "b", right: "r" };
  const wantLegend =
    o.legend === "none" ? false : o.legend ? true : perPoint || series.length > 1;
  Object.assign(common, {
    showLegend: wantLegend,
    legendPos: legendMap[o.legend] || (perPoint ? "r" : "b"),
    legendColor: muted,
    legendFontFace: font,
    legendFontSize: 12,
  });

  // axes
  if (o.x_title) Object.assign(common, { showCatAxisTitle: true, catAxisTitle: o.x_title, catAxisTitleColor: muted, catAxisTitleFontFace: font, catAxisTitleFontSize: 12 });
  if (o.y_title) Object.assign(common, { showValAxisTitle: true, valAxisTitle: o.y_title, valAxisTitleColor: muted, valAxisTitleFontFace: font, valAxisTitleFontSize: 12 });
  if (o.y_min != null) common.valAxisMinVal = o.y_min;
  if (o.y_max != null) common.valAxisMaxVal = o.y_max;

  const labeled = (s) => ({ name: s.name, labels: categories, values: s.values });

  switch (type) {
    case "bar": {
      const stacked = o.stacked === "stacked" || o.stacked === "percent";
      slide.addChart(pptx.charts.BAR, series.map(labeled), {
        ...common,
        barDir: o.horizontal ? "bar" : "col",
        ...(o.horizontal
          ? {
              catAxisOrientation: "maxMin",
              ...(o.data_labels !== false ? { valAxisHidden: true, valGridLine: { style: "none" } } : {}),
            }
          : {}),
        barGrouping: o.stacked === "percent" ? "percentStacked" : o.stacked === "stacked" ? "stacked" : "clustered",
        barGapWidthPct: o.gap_width ?? 60,
        dataLabelPosition: stacked ? "ctr" : "outEnd",
        ...(stacked ? { dataLabelColor: readableOn(colors[0]) } : {}),
      });
      break;
    }
    case "line":
      slide.addChart(pptx.charts.LINE, series.map(labeled), {
        ...common,
        lineSize: 3,
        lineDataSymbol: "circle",
        lineDataSymbolSize: 8,
        lineSmooth: !!o.smooth,
        dataLabelPosition: "t",
      });
      break;
    case "area":
      slide.addChart(pptx.charts.AREA, series.map(labeled), {
        ...common,
        barGrouping: o.stacked === "percent" ? "percentStacked" : o.stacked === "stacked" ? "stacked" : "standard",
        chartColorsOpacity: 85,
      });
      break;
    case "pie":
    case "doughnut": {
      const s = series[0];
      slide.addChart(type === "pie" ? pptx.charts.PIE : pptx.charts.DOUGHNUT, [labeled(s)], {
        ...common,
        showPercent: o.data_labels !== false,
        showValue: false,
        dataLabelColor: readableOn(colors[0]),
        dataLabelPosition: type === "pie" ? "bestFit" : undefined,
        dataLabelFormatCode: "0%",
        dataBorder: { pt: 1.5, color: bg },
        ...(type === "doughnut" ? { holeSize: 60 } : {}),
      });
      break;
    }
    case "scatter":
      slide.addChart(
        pptx.charts.SCATTER,
        [{ name: spec.x, values: xs }, ...series.map((s) => ({ name: s.name, values: s.values }))],
        {
          ...common,
          lineSize: 0,
          lineDataSymbol: "circle",
          lineDataSymbolSize: 10,
          showValue: false,
          catAxisLabelFormatCode: "#,##0",
        }
      );
      break;
    case "radar":
      slide.addChart(pptx.charts.RADAR, series.map(labeled), {
        ...common,
        radarStyle: "marker",
        lineSize: 2.5,
        lineDataSymbolSize: 7,
        showValue: false,
      });
      break;
    default:
      throw new ChartError(`unsupported chart type ${JSON.stringify(type)}`);
  }
}

module.exports = { addChartObject, ChartError };
