/**
 * 描画チェック（scripts/render-check.js）用のページ。アドインと同じ renderMermaid で描き、結果を返す。
 */
import { renderMermaid, svgForBox } from "../../src/office/renderer";
import { SAMPLES } from "../../src/core/samples";
import { PX_TO_PT, placeNew } from "../../src/core/layout";
import { DEFAULT_CUSTOM_COLORS, THEMES, paletteFor, type TemplateColors, type Theme } from "../../src/core/theme";

/** Office の既定テーマ（Office 2013-2022）の配色。「テンプレートの色」の確認に使う */
const OFFICE_COLORS: TemplateColors = {
  dk1: "#000000",
  lt1: "#FFFFFF",
  dk2: "#44546A",
  lt2: "#E7E6E6",
  accents: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
};

const SHAPES = "rect, path, polygon, circle, ellipse, line, polyline";

export interface RenderResult {
  ok: boolean;
  error?: string;
  width?: number;
  height?: number;
  bytes?: number;
  foreignObjects?: number;
  /** 塗り・線の色を属性（fill / stroke / style）で持つ図形の数と、図形の総数 */
  shapesWithColorAttr?: number;
  shapes?: number;
  svg?: string;
}

async function render(code: string, theme: Theme): Promise<RenderResult> {
  try {
    const r = await renderMermaid(code, theme, paletteFor(theme, OFFICE_COLORS, DEFAULT_CUSTOM_COLORS));
    const svg = new DOMParser().parseFromString(r.svg, "image/svg+xml").documentElement;
    const shapes = Array.from(svg.querySelectorAll(SHAPES)).filter((el) => !el.closest("defs, marker"));
    const colored = shapes.filter(
      (el) => el.hasAttribute("fill") || el.hasAttribute("stroke") || /(^|;)\s*(fill|stroke)\s*:/.test(el.getAttribute("style") ?? ""),
    );
    return {
      ok: true,
      width: r.width,
      height: r.height,
      bytes: r.svg.length,
      foreignObjects: svg.querySelectorAll("foreignObject").length,
      shapesWithColorAttr: colored.length,
      shapes: shapes.length,
      svg: r.svg,
    };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}

const MODERN: Theme = "modern";
const FONT_PT = 14;
const renderModern = (code: string) =>
  renderMermaid(code, MODERN, paletteFor(MODERN, OFFICE_COLORS, DEFAULT_CUSTOM_COLORS), { fontPt: FONT_PT });
/**
 * 図の大きさに関わる部分（SVG の大きさ・図全体の伸ばし方・文字の打ち消し方）。
 * mermaid は描くたびに id や線の細かな曲がり方が変わるので、SVG の文字列全体では比べない
 */
function sizeSignature(svgText: string): string {
  const svg = new DOMParser().parseFromString(svgText, "image/svg+xml").documentElement;
  const root = ["width", "height", "viewBox"].map((a) => svg.getAttribute(a));
  const stretchG = svg.firstElementChild?.getAttribute("transform");
  const texts = Array.from(svg.querySelectorAll("text")).map((t) => t.getAttribute("transform"));
  return JSON.stringify([root, stretchG, texts]);
}

export interface RoundTripResult {
  ok: boolean;
  error?: string;
  /** 新規挿入した図と、何も変えずに作り直した図が同じか */
  same?: boolean;
  /** 違ったときの最初の差（前後 80 文字） */
  diff?: string[];
}

/** 新規挿入（placeNew）したあと、同じ枠のまま作り直すと同じ図になるか。スライドは 16:9（960x540pt） */
async function roundTrip(code: string): Promise<RoundTripResult> {
  try {
    const first = await renderModern(code);
    const box = placeNew(first, { width: 960, height: 540 });
    const inserted = svgForBox(first, box);
    const replaced = svgForBox(await renderModern(code), box);
    const a = sizeSignature(inserted);
    const b = sizeSignature(replaced);
    if (a === b) return { ok: true, same: true };
    let i = 0;
    while (a[i] === b[i]) i++;
    return { ok: true, same: false, diff: [a.slice(i - 40, i + 80), b.slice(i - 40, i + 80)] };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}

export interface StretchResult {
  ok: boolean;
  error?: string;
  /** 伸ばした図の文字の高さ ÷ 等倍の文字の高さ（中央値）。期待値は縦横の倍率の小さい方（1 が上限） */
  textScale?: number;
  expected?: number;
  /** 図が枠の幅・高さを占める割合の、等倍のときに対する比（1 なら枠いっぱいに伸びている） */
  fillW?: number;
  fillH?: number;
}

function medianTextHeight(svg: SVGSVGElement): number {
  const hs = Array.from(svg.querySelectorAll("text"))
    .map((t) => t.getBoundingClientRect().height)
    .filter((h) => h > 0)
    .sort((a, b) => a - b);
  return hs.length ? hs[Math.floor(hs.length / 2)] : 0;
}

/** 図形（文字を除く）が占める範囲（viewBox の幅・高さに対する割合） */
function contentRatio(svg: SVGSVGElement): { w: number; h: number } {
  const texts = Array.from(svg.querySelectorAll<SVGTextElement>("text"));
  texts.forEach((t) => (t.style.display = "none"));
  const bb = svg.getBBox();
  texts.forEach((t) => (t.style.display = ""));
  const vb = svg.viewBox.baseVal;
  return { w: bb.width / vb.width, h: bb.height / vb.height };
}

/** 等倍の枠を幅 widthRatio 倍・高さ heightRatio 倍にして作り直したとき、図形は枠に伸び、文字は歪まないか */
async function stretch(code: string, widthRatio: number, heightRatio: number): Promise<StretchResult> {
  try {
    const r = await renderModern(code);
    const box = { width: r.width * PX_TO_PT * widthRatio, height: r.height * PX_TO_PT * heightRatio };
    const holder = document.createElement("div");
    holder.style.cssText = "position:absolute;left:0;top:0;";
    document.body.appendChild(holder);
    try {
      holder.innerHTML = r.svg;
      const natural = holder.querySelector("svg")!;
      const naturalText = medianTextHeight(natural);
      const naturalFill = contentRatio(natural);
      holder.innerHTML = svgForBox(r, box);
      const svg = holder.querySelector("svg")!;
      const fill = contentRatio(svg);
      return {
        ok: true,
        textScale: naturalText > 0 ? medianTextHeight(svg) / naturalText : 1,
        expected: Math.min(1, widthRatio, heightRatio),
        // 等倍のときと同じ割合なら、図形が枠いっぱいに伸びている
        fillW: fill.w / naturalFill.w,
        fillH: fill.h / naturalFill.h,
      };
    } finally {
      holder.remove();
    }
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}

export interface FontChangeResult {
  ok: boolean;
  error?: string;
  /** 図形の伸ばし方と SVG の大きさが、文字の大きさを変える前と同じか */
  sameShapes?: boolean;
  /** 変えたあとの文字の高さ ÷ 変える前の文字の高さ（中央値） */
  textScale?: number;
}

/** 入れたあとに文字の大きさを 14pt から toPt に変えて作り直すと、図形はそのままで文字だけ変わるか */
async function fontChange(code: string, toPt: number): Promise<FontChangeResult> {
  try {
    const r = await renderModern(code);
    const box = placeNew(r, { width: 960, height: 540 });
    const before = svgForBox(r, box);
    const after = svgForBox(await renderModern(code), box, toPt / FONT_PT);
    const shapes = (svg: string) => JSON.parse(sizeSignature(svg)).slice(0, 2);
    const holder = document.createElement("div");
    holder.style.cssText = "position:absolute;left:0;top:0;";
    document.body.appendChild(holder);
    try {
      holder.innerHTML = before;
      const a = medianTextHeight(holder.querySelector("svg")!);
      holder.innerHTML = after;
      const b = medianTextHeight(holder.querySelector("svg")!);
      return {
        ok: true,
        sameShapes: JSON.stringify(shapes(before)) === JSON.stringify(shapes(after)),
        textScale: a > 0 ? b / a : toPt / FONT_PT,
      };
    } finally {
      holder.remove();
    }
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}

Object.assign(window, { m4pFontChange: fontChange, m4pRoundTrip: roundTrip, m4pStretch: stretch, m4pRender: render, m4pSamples: SAMPLES, m4pThemes: THEMES });
