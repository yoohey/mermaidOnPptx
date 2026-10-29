/**
 * Mermaid ソースを SVG に描画する。タスクペイン（WebView2）の DOM 上で mermaid.js を動かす。
 * PowerPoint は SVG 内の foreignObject（HTML ラベル）を描けない可能性が高いため htmlLabels を切る。
 * 実機（Windows 版 2610）で、この設定のまま CSS をインライン化せずに色・日本語が出ることを確認済み。
 */
import mermaid from "mermaid";
import { mermaidThemeConfig, themeKey, type TemplateColors, type Theme } from "../core/theme";

declare const __MERMAID_VERSION__: string;

/** ビルド時に node_modules/mermaid/package.json から埋め込む */
export const MERMAID_VERSION: string = __MERMAID_VERSION__;

const FONT_FAMILY = '"Yu Gothic UI", "Yu Gothic", Meiryo, "Segoe UI", sans-serif';

/** 独自スタイルでは themeVariables 側のフォント指定が優先されるため、そちらにも入れる */
function withFont(config: ReturnType<typeof mermaidThemeConfig>) {
  return config.themeVariables
    ? { ...config, themeVariables: { ...config.themeVariables, fontFamily: FONT_FAMILY } }
    : config;
}

let initializedKey: string | null = null;

function ensureInitialized(theme: Theme, colors: TemplateColors | null) {
  const key = themeKey(theme, colors);
  if (initializedKey === key) return;
  // ブロック先頭の %%{init: {"theme": ...}}%% は、ここで指定した theme より優先される（mermaid の仕様）
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: "strict",
    ...withFont(mermaidThemeConfig(theme, colors)),
    fontFamily: FONT_FAMILY,
    htmlLabels: false,
    flowchart: { htmlLabels: false },
    class: { htmlLabels: false },
    state: { htmlLabels: false },
    // ガントは既定だとパネルの幅に合わせて描かれ、文字も小さい。幅を固定して読める大きさにする
    gantt: { useWidth: 800, fontSize: 14, sectionFontSize: 14, barHeight: 24 },
  } as Parameters<typeof mermaid.initialize>[0]);
  initializedKey = key;
}

export interface RenderedSvg {
  svg: string;
  /** SVG の自然な大きさ（px） */
  width: number;
  height: number;
}

let seq = 0;

/** colors は「テンプレートの色」「カスタム」で使う配色（paletteFor の値） */
export async function renderMermaid(code: string, theme: Theme, colors: TemplateColors | null = null): Promise<RenderedSvg> {
  ensureInitialized(theme, colors);
  // 構文エラーはここで「Parse error on line N」のようなメッセージとして投げられる
  await mermaid.parse(code);
  const { svg } = await mermaid.render(`m4p-${Date.now()}-${seq++}`, code);
  return normalizeSvg(svg, theme);
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

/** PowerPoint に渡しやすいよう、幅・高さを viewBox から明示した単体の SVG に整える */
function normalizeSvg(svgText: string, theme: Theme): RenderedSvg {
  const doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
  const svg = doc.documentElement;
  const viewBox = (svg.getAttribute("viewBox") ?? "").split(/[\s,]+/).map(Number);
  const width = viewBox.length === 4 && viewBox[2] > 0 ? viewBox[2] : 400;
  const height = viewBox.length === 4 && viewBox[3] > 0 ? viewBox[3] : 300;

  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", String(Math.ceil(width)));
  svg.setAttribute("height", String(Math.ceil(height)));
  svg.removeAttribute("style"); // mermaid は max-width を style で付ける
  fixErLabels(svg, theme);

  return { svg: new XMLSerializer().serializeToString(svg), width, height };
}
