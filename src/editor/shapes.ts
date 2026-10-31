// Word's preset shapes (arrows, chevrons, stars …) as SVG outlines, and SmartArt diagrams drawn from
// the shapes Word keeps with them (word/diagrams/drawingN.xml). Only how they look: the file keeps them.

/** A shape's outline in a w × h box, as an SVG path, or null for a preset not drawn here. */
export function presetPath(kind: string, w: number, h: number): string | null {
  const m = Math.min(w, h);
  const poly = (pts: number[][]) => `M${pts.map((p) => p.join(" ")).join("L")}Z`;
  switch (kind) {
    case "rect":
      return poly([[0, 0], [w, 0], [w, h], [0, h]]);
    case "roundRect": {
      const r = m / 6;
      return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w} ${r}V${h - r}A${r} ${r} 0 0 1 ${w - r} ${h}H${r}A${r} ${r} 0 0 1 0 ${h - r}V${r}A${r} ${r} 0 0 1 ${r} 0Z`;
    }
    case "ellipse":
      return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
    case "triangle":
      return poly([[w / 2, 0], [w, h], [0, h]]);
    case "rtTriangle":
      return poly([[0, 0], [w, h], [0, h]]);
    case "diamond":
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case "parallelogram":
      return poly([[m / 4, 0], [w, 0], [w - m / 4, h], [0, h]]);
    case "trapezoid":
      return poly([[m / 4, 0], [w - m / 4, 0], [w, h], [0, h]]);
    case "hexagon":
      return poly([[m / 4, 0], [w - m / 4, 0], [w, h / 2], [w - m / 4, h], [m / 4, h], [0, h / 2]]);
    case "octagon": {
      const d = m * 0.29;
      return poly([[d, 0], [w - d, 0], [w, d], [w, h - d], [w - d, h], [d, h], [0, h - d], [0, d]]);
    }
    case "pentagon":
      return poly([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]]);
    case "plus": {
      const d = m / 4;
      return poly([[d, 0], [w - d, 0], [w - d, d], [w, d], [w, h - d], [w - d, h - d], [w - d, h], [d, h], [d, h - d], [0, h - d], [0, d], [d, d]]);
    }
    case "star5": {
      const pts: number[][] = [];
      for (let i = 0; i < 10; i++) {
        const a = -Math.PI / 2 + (i * Math.PI) / 5;
        const r = i % 2 ? 0.38 : 1;
        pts.push([w / 2 + (w / 2) * r * Math.cos(a), h / 2 + (h / 2) * r * Math.sin(a) + h * 0.05]);
      }
      return poly(pts);
    }
    case "chevron":
      return poly([[0, 0], [w - m / 2, 0], [w, h / 2], [w - m / 2, h], [0, h], [m / 2, h / 2]]);
    case "homePlate":
      return poly([[0, 0], [w - m / 2, 0], [w, h / 2], [w - m / 2, h], [0, h]]);
    case "rightArrow":
      return poly([[0, h / 4], [w - m / 2, h / 4], [w - m / 2, 0], [w, h / 2], [w - m / 2, h], [w - m / 2, (3 * h) / 4], [0, (3 * h) / 4]]);
    case "leftArrow":
      return poly([[w, h / 4], [m / 2, h / 4], [m / 2, 0], [0, h / 2], [m / 2, h], [m / 2, (3 * h) / 4], [w, (3 * h) / 4]]);
    case "upArrow":
      return poly([[w / 4, h], [w / 4, m / 2], [0, m / 2], [w / 2, 0], [w, m / 2], [(3 * w) / 4, m / 2], [(3 * w) / 4, h]]);
    case "downArrow":
      return poly([[w / 4, 0], [w / 4, h - m / 2], [0, h - m / 2], [w / 2, h], [w, h - m / 2], [(3 * w) / 4, h - m / 2], [(3 * w) / 4, 0]]);
    case "leftRightArrow":
      return poly([[0, h / 2], [m / 2, 0], [m / 2, h / 4], [w - m / 2, h / 4], [w - m / 2, 0], [w, h / 2], [w - m / 2, h], [w - m / 2, (3 * h) / 4], [m / 2, (3 * h) / 4], [m / 2, h]]);
    case "upDownArrow":
      return poly([[w / 2, 0], [w, m / 2], [(3 * w) / 4, m / 2], [(3 * w) / 4, h - m / 2], [w, h - m / 2], [w / 2, h], [0, h - m / 2], [w / 4, h - m / 2], [w / 4, m / 2], [0, m / 2]]);
    case "flowChartProcess":
      return presetPath("rect", w, h);
    case "flowChartAlternateProcess":
      return presetPath("roundRect", w, h);
    case "flowChartDecision":
      return presetPath("diamond", w, h);
    case "flowChartTerminator": {
      const r = Math.min(h / 2, w / 2);
      return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w - r} ${h}H${r}A${r} ${r} 0 0 1 ${r} 0Z`;
    }
    default:
      return null;
  }
}

/** One shape to draw: its place and size in points, its colours and lines of text. */
export interface DrawnShape {
  kind: string;
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string | null;
  stroke: string | null;
  thickness: number;
  text: string[];
  color: string;
  size: number;
}

const SVG = "http://www.w3.org/2000/svg";

/** Shapes drawn into a box w × h points, each with its text in the middle. */
export function drawShapes(shapes: DrawnShape[], w: number, h: number): HTMLElement {
  const box = document.createElement("span");
  box.className = "shape-group";
  box.style.cssText = `width:${w}pt;height:${h}pt`;
  const svg = document.createElementNS(SVG, "svg");
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.setAttribute("width", "100%");
  svg.setAttribute("height", "100%");
  for (const s of shapes) {
    const d = presetPath(s.kind, s.w, s.h) ?? presetPath("rect", s.w, s.h)!;
    const path = document.createElementNS(SVG, "path");
    path.setAttribute("d", d);
    path.setAttribute("transform", `translate(${s.x} ${s.y})`);
    path.setAttribute("fill", s.fill ?? "none");
    if (s.stroke) {
      path.setAttribute("stroke", s.stroke);
      path.setAttribute("stroke-width", String(s.thickness));
    }
    svg.append(path);
  }
  box.append(svg);
  for (const s of shapes) {
    if (!s.text.some((t) => t.trim())) continue;
    const label = document.createElement("span");
    label.className = "shape-label";
    label.style.cssText = `left:${s.x}pt;top:${s.y}pt;width:${s.w}pt;height:${s.h}pt;color:${s.color};font-size:${s.size}pt`;
    for (const line of s.text) label.append(Object.assign(document.createElement("span"), { textContent: line || " " }));
    box.append(label);
  }
  return box;
}

// shortcut: theme colours are Office 2013–2022's; read the document's theme if a template differs.
const SCHEME: Record<string, string> = { accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5", accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47", lt1: "FFFFFF", bg1: "FFFFFF", dk1: "000000", tx1: "000000", lt2: "E7E6E6", dk2: "44546A" };
export const colorIn = (xml: string | undefined) => {
  const m = /<a:(srgbClr|schemeClr|prstClr) val="(\w+)"\s*(\/>|>([\s\S]*?)<\/a:\1>)/.exec(xml ?? "");
  if (!m) return null;
  const hex = m[1] === "srgbClr" ? m[2] : m[1] === "schemeClr" ? (SCHEME[m[2]] ?? "4472C4") : m[2] === "white" ? "FFFFFF" : "000000";
  return adjusted(hex, m[4] ?? "");
};

/** A colour with Word's adjustments applied: shade (darker), tint (lighter), lumMod/lumOff ("lighter 40%"). */
function adjusted(hex: string, mods: string): string {
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const val = (name: string) => {
    const v = new RegExp(`<a:${name} val="(\\d+)"`).exec(mods)?.[1];
    return v === undefined ? null : Number(v) / 100000;
  };
  const shade = val("shade"), tint = val("tint"), lumMod = val("lumMod"), lumOff = val("lumOff");
  // Office shades and tints in linear light, not on the sRGB numbers.
  const linear = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const gamma = (c: number) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
  if (shade !== null) [r, g, b] = [r, g, b].map((c) => gamma(linear(c) * shade));
  if (tint !== null) [r, g, b] = [r, g, b].map((c) => gamma(linear(c) + (1 - linear(c)) * (1 - tint)));
  if (lumMod !== null || lumOff !== null) {
    // Through hue, saturation and lightness, as the luminance is changed.
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let l = (max + min) / 2;
    const d = max - min;
    const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
    let h = d === 0 ? 0 : max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h = (h * 60 + 360) % 360;
    l = Math.min(1, Math.max(0, l * (lumMod ?? 1) + (lumOff ?? 0)));
    const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m0 = l - c / 2;
    const [r1, g1, b1] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
    [r, g, b] = [r1 + m0, g1 + m0, b1 + m0];
  }
  return `#${[r, g, b].map((c) => Math.round(Math.min(1, Math.max(0, c)) * 255).toString(16).padStart(2, "0").toUpperCase()).join("")}`;
}

/** The shapes of a SmartArt diagram's drawing part, in points from its top-left corner. */
export function readDiagram(xml: string): DrawnShape[] {
  const emu = (v: string | undefined) => (Number(v) || 0) / 12700;
  const out: DrawnShape[] = [];
  for (const sp of xml.split(/<dsp:sp\b/).slice(1)) {
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(sp);
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(sp);
    if (!off || !ext) continue;
    const spPr = /<dsp:spPr>([\s\S]*?)<\/dsp:spPr>/.exec(sp)?.[1] ?? "";
    const ln = /<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/.exec(spPr);
    const own = spPr.replace(/<a:ln\b[\s\S]*?<\/a:ln>/, "");
    const text = (/<dsp:txBody>([\s\S]*?)<\/dsp:txBody>/.exec(sp)?.[1] ?? "").split(/<\/a:p>/).slice(0, -1).map((p) => [...p.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => t[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")).join(""));
    const size = Number(/<a:(?:rPr|endParaRPr)\b[^>]*\bsz="(\d+)"/.exec(sp)?.[1] ?? 1100) / 100;
    out.push({
      kind: /<a:prstGeom prst="(\w+)"/.exec(spPr)?.[1] ?? "rect",
      x: emu(off[1]),
      y: emu(off[2]),
      w: emu(ext[1]),
      h: emu(ext[2]),
      fill: /<a:noFill\/>/.test(own) ? null : colorIn(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(own)?.[1]),
      stroke: ln && !/<a:noFill\/>/.test(ln[2]) ? colorIn(ln[2]) : null,
      thickness: emu(/w="(\d+)"/.exec(ln?.[1] ?? "")?.[1] ?? "12700"),
      text,
      color: colorIn(/<a:fontRef[^>]*>([\s\S]*?)<\/a:fontRef>/.exec(sp)?.[1]) ?? "#000000",
      size: Math.min(size, 40),
    });
  }
  return out;
}
