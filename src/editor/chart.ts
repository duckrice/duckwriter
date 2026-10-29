// Word's charts, drawn from the numbers Word keeps in the chart part (its caches), as DuckSheets
// draws Excel's. The chart itself is kept in the file as it is; this is only how it is shown.

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E", "#636363", "#997300"];

export type ChartKind = "column" | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter";

export interface ChartData {
  kind: ChartKind;
  title: string;
  xTitle?: string;
  yTitle?: string;
  categories: string[];
  /** For a scatter chart, the x of each point. */
  xs: (number | null)[];
  series: { name: string; values: (number | null)[]; color?: string }[];
  /** Pie slice colours, where set. */
  points?: (string | null)[];
  decimals: number | null;
  percent: boolean;
}

const kids = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const kid = (el: Element | null | undefined, name: string) => kids(el, name)[0] ?? null;
const deep = (el: Element | null | undefined, name: string) => (el ? (Array.from(el.getElementsByTagNameNS("*", name))[0] ?? null) : null);
const words = (el: Element | null) =>
  Array.from(el?.getElementsByTagNameNS(A, "t") ?? [])
    .map((t) => t.textContent ?? "")
    .join("");
const rgb = (el: Element | null) => {
  const c = el && deep(el, "srgbClr")?.getAttribute("val");
  return c ? `#${c.toUpperCase()}` : null;
};

/** A cache's points by index (strRef/numRef/strLit/numLit all hold c:pt elements). */
function cached(el: Element | null): string[] {
  const out: string[] = [];
  const cache = el && (deep(el, "strCache") ?? deep(el, "numCache") ?? deep(el, "strLit") ?? deep(el, "numLit"));
  const count = Number(kid(cache, "ptCount")?.getAttribute("val") ?? 0);
  for (let i = 0; i < count; i++) out.push("");
  for (const pt of kids(cache, "pt")) out[Number(pt.getAttribute("idx") ?? 0)] = kid(pt, "v")?.textContent ?? "";
  return out;
}
const numbers = (el: Element | null) => cached(el).map((v) => (v === "" || !Number.isFinite(Number(v)) ? null : Number(v)));

/** The chart in a chart part, or null when it isn't one DuckWriter can draw. */
export function readChart(text: string): ChartData | null {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const plot = doc.getElementsByTagNameNS(C, "plotArea")[0];
  // shortcut: a combined chart (columns and a line) shows only its first kind; draw each kind if people use them.
  const type = Array.from(plot?.children ?? []).find((el) => /Chart$/.test(el.localName));
  if (!type) return null;
  const kind: ChartKind | null =
    type.localName === "barChart"
      ? kid(type, "barDir")?.getAttribute("val") === "bar"
        ? "bar"
        : "column"
      : ({ lineChart: "line", areaChart: "area", pieChart: "pie", doughnutChart: "doughnut", scatterChart: "scatter" } as Record<string, ChartKind>)[type.localName] ?? null;
  if (!kind) return null;
  const sers = kids(type, "ser");
  if (!sers.length) return null;
  const scatter = kind === "scatter";
  const format = deep(kid(sers[0], scatter ? "yVal" : "val"), "formatCode")?.textContent ?? "General";
  const data: ChartData = {
    kind,
    title: words(kid(doc.getElementsByTagNameNS(C, "chart")[0], "title")),
    categories: cached(kid(sers[0], scatter ? "xVal" : "cat")),
    xs: scatter ? numbers(kid(sers[0], "xVal")) : [],
    series: sers.map((s, i) => ({
      name: cached(kid(s, "tx"))[0] || `Series ${i + 1}`,
      values: numbers(kid(s, scatter ? "yVal" : "val")),
      color: rgb(kid(s, "spPr")) ?? rgb(kid(kid(s, "marker"), "spPr")) ?? undefined,
    })),
    decimals: /0\.(0+)/.exec(format)?.[1].length ?? (/^0%?$|#,##0(?!\.)/.test(format) ? 0 : null),
    percent: format.includes("%"),
  };
  if (!data.categories.length) data.categories = data.series[0].values.map((_, i) => String(i + 1));
  if (kind === "pie" || kind === "doughnut") {
    data.points = [];
    for (const pt of kids(sers[0], "dPt")) data.points[Number(kid(pt, "idx")?.getAttribute("val") ?? 0)] = rgb(kid(pt, "spPr"));
  }
  for (const axis of Array.from(plot.children).filter((el) => /Ax$/.test(el.localName))) {
    const name = words(kid(axis, "title"));
    if (!name || kid(axis, "delete")?.getAttribute("val") === "1") continue;
    if (["b", "t"].includes(kid(axis, "axPos")?.getAttribute("val") ?? "")) data.xTitle = name;
    else data.yTitle = name;
  }
  return data;
}

/** Round steps for an axis: 0, 2 000, 4 000 … rather than 0, 1 873 … */
function niceScale(min: number, max: number, ticks = 5): { lo: number; hi: number; step: number } {
  if (min === max) {
    min = Math.min(0, min);
    max = max === 0 ? 1 : Math.max(0, max);
  }
  const lo0 = Math.min(0, min);
  const hi0 = Math.max(0, max);
  const raw = (hi0 - lo0) / ticks;
  const magnitude = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * magnitude).find((s) => s >= raw) ?? 10 * magnitude;
  return { lo: Math.floor(lo0 / step) * step, hi: Math.ceil(hi0 / step) * step, step };
}

const SVG = "http://www.w3.org/2000/svg";
function svg(name: string, attrs: Record<string, string | number> = {}): SVGElement {
  const el = document.createElementNS(SVG, name);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
}
const text = (x: number, y: number, content: string, attrs: Record<string, string | number> = {}) => {
  const el = svg("text", { x, y, "font-size": 11, fill: "#404040", "font-family": "Calibri, Carlito, sans-serif", ...attrs });
  el.textContent = content;
  return el;
};

/** The chart drawn as an SVG, w × h (in points; the SVG scales with the page). */
export function drawChart(chart: ChartData, w: number, h: number): SVGSVGElement {
  const root = svg("svg", { viewBox: `0 0 ${w} ${h}`, width: "100%", height: "100%", class: "chart-svg" }) as SVGSVGElement;
  root.append(svg("rect", { x: 0, y: 0, width: w, height: h, fill: "#ffffff", stroke: "#d9d9d9" }));
  const round = chart.kind === "pie" || chart.kind === "doughnut";
  const color = (i: number) => (round ? chart.points?.[i] : chart.series[i]?.color) ?? COLORS[i % COLORS.length];
  const shown = (n: number) => {
    const v = chart.percent ? n * 100 : n;
    const s = chart.decimals === null ? String(Math.round(v * 1e6) / 1e6) : v.toFixed(chart.decimals);
    return chart.percent ? `${s}%` : s;
  };
  const { categories, series, xs } = chart;
  let top = 12;
  if (chart.title) {
    root.append(text(w / 2, 26, chart.title, { "font-size": 14, "text-anchor": "middle", fill: "#262626" }));
    top = 42;
  }
  const legendItems = round ? categories : series.length > 1 ? series.map((s) => s.name) : [];
  const legendH = legendItems.length ? 22 : 0;
  if (legendItems.length) {
    let x = 0;
    const items = legendItems.map((name, i) => {
      const g = svg("g");
      g.append(svg("rect", { x, y: 0, width: 9, height: 9, fill: color(i) }), text(x + 13, 9, name));
      x += 26 + name.length * 6.2;
      return g;
    });
    const legend = svg("g", { transform: `translate(${Math.max(8, (w - x) / 2)},${h - 16})` });
    legend.append(...items);
    root.append(legend);
  }
  const bottom = h - 10 - legendH - (chart.xTitle && !round ? 18 : 0);
  const left = chart.yTitle && !round ? 18 : 0;
  const axisTitles = (plot: { x: number; y: number; w: number; h: number }) => {
    if (chart.xTitle) root.append(text(plot.x + plot.w / 2, bottom + 15, chart.xTitle, { "text-anchor": "middle", "font-weight": "bold", fill: "#595959" }));
    if (chart.yTitle) root.append(text(0, 0, chart.yTitle, { "text-anchor": "middle", "font-weight": "bold", fill: "#595959", transform: `translate(12,${plot.y + plot.h / 2}) rotate(-90)` }));
  };

  if (round) {
    const values = series[0].values.map((v) => Math.max(0, v ?? 0));
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const r = Math.max(10, Math.min(w - 20, bottom - top) / 2 - 4);
    const cx = w / 2;
    const cy = (top + bottom) / 2;
    let angle = -Math.PI / 2;
    values.forEach((v, i) => {
      const a = (v / total) * Math.PI * 2;
      const end = angle + a;
      const p = (t: number) => `${cx + r * Math.cos(t)},${cy + r * Math.sin(t)}`;
      const d = values.filter((x) => x > 0).length === 1 && v > 0 ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` : `M${cx},${cy}L${p(angle)}A${r},${r} 0 ${a > Math.PI ? 1 : 0} 1 ${p(end)}Z`;
      root.append(svg("path", { d, fill: color(i), stroke: "#ffffff", "stroke-width": 1.5 }));
      angle = end;
    });
    if (chart.kind === "doughnut") root.append(svg("circle", { cx, cy, r: r * 0.5, fill: "#ffffff" }));
    return root;
  }

  const all = series.flatMap((s) => s.values).filter((v): v is number => v !== null);
  if (!all.length) return root;
  if (chart.kind === "scatter") {
    const xv = xs.filter((v): v is number => v !== null);
    if (!xv.length) return root;
    const sx = niceScale(Math.min(...xv), Math.max(...xv));
    const sy = niceScale(Math.min(...all), Math.max(...all));
    const labelW = Math.max(...[sy.lo, sy.hi].map((n) => shown(n).length)) * 6.5 + 10;
    const plot = { x: left + labelW + 6, y: top, w: w - left - labelW - 20, h: bottom - top - 20 };
    if (plot.w < 20 || plot.h < 20) return root;
    const X = (v: number) => plot.x + ((v - sx.lo) / (sx.hi - sx.lo)) * plot.w;
    const Y = (v: number) => plot.y + plot.h - ((v - sy.lo) / (sy.hi - sy.lo)) * plot.h;
    for (let v = sy.lo; v <= sy.hi + sy.step / 2; v += sy.step) root.append(svg("line", { x1: plot.x, y1: Y(v), x2: plot.x + plot.w, y2: Y(v), stroke: "#d9d9d9" }), text(plot.x - 6, Y(v) + 4, shown(v), { "text-anchor": "end" }));
    for (let v = sx.lo; v <= sx.hi + sx.step / 2; v += sx.step) root.append(text(X(v), plot.y + plot.h + 15, String(Math.round(v * 1e6) / 1e6), { "text-anchor": "middle" }));
    root.append(svg("line", { x1: plot.x, y1: plot.y + plot.h, x2: plot.x + plot.w, y2: plot.y + plot.h, stroke: "#bfbfbf" }));
    series.forEach((s, k) =>
      s.values.forEach((v, i) => {
        const x = xs[i];
        if (v !== null && x !== null && x !== undefined) root.append(svg("circle", { cx: X(x), cy: Y(v), r: 3.5, fill: color(k) }));
      }),
    );
    axisTitles(plot);
    return root;
  }

  const scale = niceScale(Math.min(...all), Math.max(...all));
  const horizontal = chart.kind === "bar";
  const labelW = Math.max(...[scale.lo, scale.hi].map((n) => shown(n).length)) * 6.5 + 10;
  const catW = horizontal ? Math.min(w * 0.3, Math.max(...categories.map((c) => c.length)) * 6.2 + 12) : 0;
  const plot = { x: left + (horizontal ? catW + 8 : labelW + 6), y: top, w: 0, h: 0 };
  plot.w = w - plot.x - 14;
  plot.h = bottom - top - (horizontal ? 18 : 20);
  if (plot.w < 20 || plot.h < 20) return root;
  const pos = (v: number) => (v - scale.lo) / (scale.hi - scale.lo);
  for (let v = scale.lo; v <= scale.hi + scale.step / 2; v += scale.step) {
    const t = pos(v);
    if (horizontal) {
      const x = plot.x + t * plot.w;
      root.append(svg("line", { x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: "#d9d9d9" }), text(x, plot.y + plot.h + 14, shown(v), { "text-anchor": "middle" }));
    } else {
      const y = plot.y + plot.h - t * plot.h;
      root.append(svg("line", { x1: plot.x, y1: y, x2: plot.x + plot.w, y2: y, stroke: "#d9d9d9" }), text(plot.x - 6, y + 4, shown(v), { "text-anchor": "end" }));
    }
  }
  const zero = pos(Math.max(scale.lo, Math.min(0, scale.hi)));
  const n = categories.length;
  const band = (horizontal ? plot.h : plot.w) / n;
  const every = Math.max(1, Math.ceil((horizontal ? 14 : Math.max(...categories.map((c) => c.length)) * 6.2 + 6) / band));
  categories.forEach((name, i) => {
    if (i % every) return;
    const mid = i * band + band / 2;
    if (horizontal) root.append(text(plot.x - 6, plot.y + mid + 4, name, { "text-anchor": "end" }));
    else root.append(text(plot.x + mid, plot.y + plot.h + 15, name, { "text-anchor": "middle" }));
  });
  if (chart.kind === "area") {
    const base = plot.y + plot.h - zero * plot.h;
    series.forEach((s, k) => {
      const points = s.values.map((v, i) => [plot.x + i * band + band / 2, plot.y + plot.h - pos(v ?? 0) * plot.h]);
      const d = `M${points[0][0]},${base}${points.map((p) => `L${p[0]},${p[1]}`).join("")}L${points[points.length - 1][0]},${base}Z`;
      root.append(svg("path", { d, fill: color(k), "fill-opacity": 0.75, stroke: color(k) }));
    });
  } else if (chart.kind === "line") {
    series.forEach((s, k) => {
      const points = s.values.map((v, i) => (v === null ? null : [plot.x + i * band + band / 2, plot.y + plot.h - pos(v) * plot.h]));
      let d = "";
      points.forEach((p, i) => {
        if (p) d += `${d && points[i - 1] ? "L" : "M"}${p[0]},${p[1]}`;
      });
      root.append(svg("path", { d, fill: "none", stroke: color(k), "stroke-width": 2.25, "stroke-linejoin": "round" }));
    });
  } else {
    // shortcut: stacked bars are drawn side by side; stack them if a document needs it.
    const gap = band * 0.25;
    const thick = (band - gap * 2) / series.length;
    series.forEach((s, k) =>
      s.values.forEach((v, i) => {
        if (v === null) return;
        const a = Math.min(pos(v), zero);
        const b = Math.max(pos(v), zero);
        const along = i * band + gap + k * thick;
        root.append(
          horizontal
            ? svg("rect", { x: plot.x + a * plot.w, y: plot.y + along, width: Math.max(0.5, (b - a) * plot.w), height: Math.max(1, thick - 1), fill: color(k) })
            : svg("rect", { x: plot.x + along, y: plot.y + plot.h - b * plot.h, width: Math.max(1, thick - 1), height: Math.max(0.5, (b - a) * plot.h), fill: color(k) }),
        );
      }),
    );
  }
  if (horizontal) root.append(svg("line", { x1: plot.x + zero * plot.w, y1: plot.y, x2: plot.x + zero * plot.w, y2: plot.y + plot.h, stroke: "#bfbfbf" }));
  else root.append(svg("line", { x1: plot.x, y1: plot.y + plot.h - zero * plot.h, x2: plot.x + plot.w, y2: plot.y + plot.h - zero * plot.h, stroke: "#bfbfbf" }));
  axisTitles(plot);
  return root;
}
