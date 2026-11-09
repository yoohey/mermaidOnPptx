/**
 * Mermaid ソースを SVG に描画する。タスクペイン（WebView2）の DOM 上で mermaid.js を動かす。
 * PowerPoint は SVG 内の foreignObject（HTML ラベル）を描けない可能性が高いため htmlLabels を切る。
 * 実機（Windows 版 2610）で、この設定のまま CSS をインライン化せずに色・日本語が出ることを確認済み。
 */
import mermaid from "mermaid";
import { mermaidThemeConfig, themeKey, type TemplateColors, type Theme } from "../core/theme";
import { DEFAULT_FONT_PT, PX_TO_PT, boxToPx, type Size } from "../core/layout";

declare const __MERMAID_VERSION__: string;

/** ビルド時に node_modules/mermaid/package.json から埋め込む */
export const MERMAID_VERSION: string = __MERMAID_VERSION__;

const FONT_FAMILY = '"Yu Gothic UI", "Yu Gothic", Meiryo, "Segoe UI", sans-serif';

/**
 * ガントの設定。既定（文字 14px、幅 800px、棒 24px）を、指定の文字の大きさに比例して大きくする。
 * 幅を文字に比例させないと、日付の目盛りが重なる
 */
function ganttConfig(fontPx: number) {
  const k = fontPx / 14;
  return {
    fontSize: fontPx,
    sectionFontSize: fontPx,
    useWidth: Math.round(800 * k),
    barHeight: Math.round(24 * k),
    barGap: Math.round(4 * k),
    leftPadding: Math.round(75 * k),
    rightPadding: Math.round(75 * k),
    gridLineStartPadding: Math.round(35 * k),
  };
}

/** ガントの日付の目盛りは mermaid が 10px 固定で描くため、指定の文字の大きさに合わせて描き直す */
function fixGanttAxis(svg: Element, fontPx: number) {
  if (svg.getAttribute("aria-roledescription") !== "gantt") return;
  const size = String(Math.round(fontPx * 0.8));
  for (const text of Array.from(svg.querySelectorAll("g.grid .tick text"))) text.setAttribute("font-size", size);
}

export interface RenderOptions {
  /** 文字の大きさ（pt）。等倍で挿入したときにこの大きさに見えるよう、mermaid には px で渡す */
  fontPt: number;
}

/** 独自スタイルでは themeVariables 側のフォント指定が優先されるため、そちらにも入れる */
function withFont(config: ReturnType<typeof mermaidThemeConfig>, fontPx: number) {
  return {
    ...config,
    themeVariables: {
      ...config.themeVariables,
      fontFamily: FONT_FAMILY,
      fontSize: `${fontPx}px`,
      // 円グラフは文字の大きさを個別の変数で持つ
      pieSectionTextSize: `${fontPx}px`,
      pieLegendTextSize: `${fontPx}px`,
      pieTitleTextSize: `${Math.round(fontPx * 1.4)}px`,
    },
  };
}

let initializedKey: string | null = null;

function ensureInitialized(theme: Theme, colors: TemplateColors | null, options: RenderOptions) {
  const key = JSON.stringify([themeKey(theme, colors), options.fontPt]);
  if (initializedKey === key) return;
  const fontPx = options.fontPt / PX_TO_PT;
  // ブロック先頭の %%{init: {"theme": ...}}%% は、ここで指定した theme より優先される（mermaid の仕様）
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: "strict",
    ...withFont(mermaidThemeConfig(theme, colors), fontPx),
    fontFamily: FONT_FAMILY,
    fontSize: fontPx,
    htmlLabels: false,
    flowchart: { htmlLabels: false },
    class: { htmlLabels: false },
    state: { htmlLabels: false },
    // シーケンス図は themeVariables の文字の大きさが効かないため、個別に指定する
    sequence: { actorFontSize: fontPx, messageFontSize: fontPx, noteFontSize: fontPx },
    // ガントは既定だとパネルの幅に合わせて描かれ、文字も小さい。幅を固定して読める大きさにする
    gantt: ganttConfig(fontPx),
  } as Parameters<typeof mermaid.initialize>[0]);
  initializedKey = key;
}

export interface RenderedSvg {
  svg: string;
  /** SVG の自然な大きさ（px） */
  width: number;
  height: number;
  viewBox: [number, number, number, number];
  /** 文字の高さ（px、中央値）。文字の無い図では 0 */
  textHeight: number;
}

let seq = 0;

/** colors は「テンプレートの色」「カスタム」で使う配色（paletteFor の値） */
export async function renderMermaid(
  code: string,
  theme: Theme,
  colors: TemplateColors | null = null,
  options: RenderOptions = { fontPt: DEFAULT_FONT_PT },
): Promise<RenderedSvg> {
  ensureInitialized(theme, colors, options);
  // 構文エラーはここで「Parse error on line N」のようなメッセージとして投げられる
  await mermaid.parse(code);
  const { svg } = await mermaid.render(`m4p-${Date.now()}-${seq++}`, code);
  return normalizeSvg(svg, theme, options.fontPt / PX_TO_PT);
}

const SHAPES = "rect, path, circle, ellipse, polygon, polyline, line, text, tspan";

/**
 * CSS（<style>）で付いている塗りと線の色・文字の大きさを、図形の属性にも書き写す。
 * PowerPoint はマインドマップなどの CSS の色を読めず真っ黒に描く（実機で確認）ため。
 * ブラウザでは CSS が属性より優先されるので、見た目は変わらない
 */
const SHAPE_PROPS = ["fill", "stroke", "stroke-width", "stroke-dasharray", "fill-opacity", "stroke-opacity", "opacity"];
const TEXT_PROPS = ["font-size", "font-family", "font-weight", "font-style", "text-anchor"];

function inlineColors(svg: SVGSVGElement) {
  for (const el of Array.from(svg.querySelectorAll<SVGElement>(SHAPES))) {
    if (el.closest("mask, clipPath, pattern")) continue;
    const style = getComputedStyle(el);
    const props = el.tagName === "text" || el.tagName === "tspan" ? [...SHAPE_PROPS, ...TEXT_PROPS] : SHAPE_PROPS;
    // 属性（stroke="none" など）を CSS が上書きしている図形もあるので、最終的な値で書き直す
    for (const prop of props) {
      const value = style.getPropertyValue(prop);
      if (value) el.setAttribute(prop, value);
    }
  }
}

/** 文字の高さの中央値（px）。文字の無い図では 0 */
function textHeightOf(svg: SVGSVGElement): number {
  const heights = Array.from(svg.querySelectorAll("text"))
    .map((t) => t.getBBox().height)
    .filter((h) => h > 0)
    .sort((a, b) => a - b);
  return heights.length > 0 ? heights[Math.floor(heights.length / 2)] : 0;
}

/** SVG を見えない場所に一時的に置いて fn を呼ぶ。計算済みスタイルや getBBox は画面に置いた要素でしか得られないため */
function inDom<T>(svgText: string, fn: (svg: SVGSVGElement) => T): T {
  const holder = document.createElement("div");
  holder.style.cssText = "position:absolute;left:-10000px;top:0;visibility:hidden;";
  holder.innerHTML = svgText;
  document.body.appendChild(holder);
  try {
    return fn(holder.querySelector("svg")!);
  } finally {
    holder.remove();
  }
}

/** 色の書き写しと文字の高さの計測をする */
function finishInDom(svgText: string): { svg: string; textHeight: number } {
  return inDom(svgText, (svg) => {
    inlineColors(svg);
    return { svg: new XMLSerializer().serializeToString(svg), textHeight: textHeightOf(svg) };
  });
}

const SVG_NS = "http://www.w3.org/2000/svg";
const num = (v: number) => String(+v.toFixed(4));

/**
 * 図を box（pt）の大きさに合わせる。箱・線などの図形は枠に合わせて縦横別々に伸び縮みし
 * （PowerPoint で図の大きさを変えたときと同じ形）、文字だけは歪めずに描き直す。
 * 文字の倍率は縦横の倍率の小さい方（1 を超えるときは 1 = 指定の大きさ）なので、伸びた箱からはみ出さない。
 * 新規挿入では枠が図と同じ縦横比のため、図全体を同じ倍率にしたものと同じになる。
 * 同じソース・同じ枠なら、何度作り直しても同じ図になる。
 * textScale は、描いたときの文字の大きさに対する、指定の文字の大きさの比。図形はそのままで、文字だけこの倍率にする
 * （文字の大きさを変えても図形の大きさを変えないため。文字を大きくすると箱からはみ出すことがある）
 */
export function svgForBox(rendered: RenderedSvg, box: Size, textScale = 1): string {
  const px = boxToPx(box);
  const [x, y, w, h] = rendered.viewBox;
  const sx = px.width / w;
  const sy = px.height / h;
  const t = Math.min(1, sx, sy) * textScale;
  return inDom(rendered.svg, (svg) => {
    // 文字の位置（左寄せは左端、中央寄せは中央、右寄せは右端と、縦の中央）を動かさずに、伸ばした分を打ち消す
    // 縦横を同じ倍率で縮めるだけ（新規挿入など）なら、文字も図と一緒に縮めればよい
    const distorted = Math.abs(t / sx - 1) > 1e-4 || Math.abs(t / sy - 1) > 1e-4;
    for (const text of distorted ? Array.from(svg.querySelectorAll<SVGTextElement>("text")) : []) {
      if (text.closest("defs, marker, mask, clipPath, pattern")) continue;
      const bb = text.getBBox();
      if (bb.width === 0 && bb.height === 0) continue;
      const anchor = text.getAttribute("text-anchor") || getComputedStyle(text).getPropertyValue("text-anchor");
      const fx = anchor === "start" ? bb.x : anchor === "end" ? bb.x + bb.width : bb.x + bb.width / 2;
      const fy = bb.y + bb.height / 2;
      const counter = `translate(${num(fx)} ${num(fy)}) scale(${num(t / sx)} ${num(t / sy)}) translate(${num(-fx)} ${num(-fy)})`;
      text.setAttribute("transform", [text.getAttribute("transform"), counter].filter(Boolean).join(" "));
    }
    // 図全体を枠の大きさ（px）に伸ばす
    const g = document.createElementNS(SVG_NS, "g");
    g.setAttribute("transform", `scale(${num(sx)} ${num(sy)}) translate(${num(-x)} ${num(-y)})`);
    while (svg.firstChild) g.appendChild(svg.firstChild);
    svg.appendChild(g);
    svg.setAttribute("viewBox", `0 0 ${num(px.width)} ${num(px.height)}`);
    svg.setAttribute("width", String(Math.ceil(px.width)));
    svg.setAttribute("height", String(Math.ceil(px.height)));
    return new XMLSerializer().serializeToString(svg);
  });
}

/**
 * ER 図の関係ラベルは、htmlLabels を切ると背景の四角と文字が同じ色（.edgeLabel .label の fill）で塗られ、
 * 文字が見えなくなる（mermaid 11.12 で確認）。背景だけ塗り直す。
 */
function fixErLabels(svg: Element, theme: Theme) {
  if (svg.getAttribute("aria-roledescription") !== "er") return;
  const fill = theme === "dark" ? "#333333" : "#FFFFFF";
  for (const rect of Array.from(svg.querySelectorAll(".edgeLabel .label rect.background"))) {
    rect.setAttribute("style", `fill: ${fill}; opacity: 0.85;`);
  }
}

/** dominant-baseline の値ごとに、同じ見た目になる下方向のずらし量（文字の高さに対する割合） */
const BASELINE_SHIFT: Record<string, number> = { central: 0.35, middle: 0.3 };

/** dy に em 単位で shift を足す。px（単位なし）なら文字の大きさから px に直して足す */
function addShift(dy: string | null, shift: number, fontPx: number): string {
  const m = /^(-?[\d.]+)(em|px)?$/.exec(dy ?? "0");
  if (!m) return `${shift}em`;
  const value = Number(m[1]);
  if (m[2] === "em") return `${+(value + shift).toFixed(3)}em`;
  return `${+(value + shift * fontPx).toFixed(2)}`;
}

/**
 * シーケンス図の箱の文字などは dominant-baseline="central" で縦の中央に置かれるが、
 * PowerPoint ではこの指定が効かず、文字が上にずれる（実機の画面で確認。指定を無視していると推測）。
 * dy でずらす形に置き換える。
 */
function fixBaselines(svg: Element) {
  for (const text of Array.from(svg.querySelectorAll("text[dominant-baseline]"))) {
    const shift = BASELINE_SHIFT[text.getAttribute("dominant-baseline") ?? ""];
    if (!shift) continue;
    text.removeAttribute("dominant-baseline");
    text.removeAttribute("alignment-baseline");
    // 1 行目だけずらせば、2 行目以降（dy="1em" など相対指定）も一緒にずれる
    const target = text.querySelector("tspan") ?? text;
    const fontPx = Number(/font-size:\s*([\d.]+)px/.exec(text.getAttribute("style") ?? "")?.[1] ?? 16);
    target.setAttribute("dy", addShift(target.getAttribute("dy"), shift, fontPx));
  }
}

/** PowerPoint に渡しやすいよう、幅・高さを viewBox から明示した単体の SVG に整える */
function normalizeSvg(svgText: string, theme: Theme, fontPx: number): RenderedSvg {
  const doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
  const svg = doc.documentElement;
  const viewBox = (svg.getAttribute("viewBox") ?? "").split(/[\s,]+/).map(Number);
  const width = viewBox.length === 4 && viewBox[2] > 0 ? viewBox[2] : 400;
  const height = viewBox.length === 4 && viewBox[3] > 0 ? viewBox[3] : 300;
  const [x, y] = viewBox.length === 4 && viewBox.every(Number.isFinite) ? viewBox : [0, 0];
  svg.setAttribute("viewBox", `${x} ${y} ${width} ${height}`);

  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", String(Math.ceil(width)));
  svg.setAttribute("height", String(Math.ceil(height)));
  svg.removeAttribute("style"); // mermaid は max-width を style で付ける
  fixErLabels(svg, theme);
  fixBaselines(svg);
  fixGanttAxis(svg, fontPx);

  const done = finishInDom(new XMLSerializer().serializeToString(svg));
  return { svg: done.svg, width, height, viewBox: [x, y, width, height], textHeight: done.textHeight };
}
