/**
 * 描画チェック（scripts/render-check.js）用のページ。アドインと同じ renderMermaid で描き、結果を返す。
 */
import { renderMermaid } from "../../src/office/renderer";
import { SAMPLES } from "../../src/core/samples";
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

Object.assign(window, { m4pRender: render, m4pSamples: SAMPLES, m4pThemes: THEMES });
