/**
 * 図を枠いっぱいに広げる。文字の大きさは変えず、mermaid の間隔の設定（箱と箱の間など）を広げて、
 * 枠の幅と高さに近い図を描かせる。Office・mermaid に依存しない純粋関数。
 */
import type { Size } from "./layout";

export type DiagramKind = "flowchart" | "sequence" | "state" | "class" | "er" | "gantt" | "other";

/** 幅・高さを広げる設定の値。どちらも無い図（円グラフなど）は広げない */
export interface Spacing {
  w?: number;
  h?: number;
}

interface Knob {
  /** mermaid の設定の場所（セクション名とキー） */
  section: string;
  key: string;
  base: number;
  min: number;
  max: number;
}

interface Knobs {
  w?: Knob;
  h?: Knob;
}

const knob = (section: string, key: string, base: number, min = base * 0.6, max = base * 20): Knob => ({
  section,
  key,
  base,
  min,
  max,
});

/** ノードの間（nodeSpacing）と段の間（rankSpacing）。横向き（LR / RL）なら段の間が幅になる */
function graphKnobs(section: string, horizontal: boolean, node = 50, rank = 50): Knobs {
  const n = knob(section, "nodeSpacing", node);
  const r = knob(section, "rankSpacing", rank);
  return horizontal ? { w: r, h: n } : { w: n, h: r };
}

/** 図の種類と向き。%% で始まる行（%%{init}%% やコメント）は飛ばす */
export function diagramKind(code: string): { kind: DiagramKind; horizontal: boolean } {
  const lines = code.split("\n").map((l) => l.trim()).filter((l) => l !== "" && !l.startsWith("%%"));
  const first = lines[0] ?? "";
  const keyword = first.split(/\s+/)[0].toLowerCase();
  const direction = (dir: string | undefined) => /^(LR|RL)$/i.test(dir ?? "");
  // 状態遷移図・クラス図・ER 図は、本文の `direction LR` で向きを変える
  const inner = lines.slice(1).find((l) => /^direction\s+/i.test(l))?.split(/\s+/)[1];
  if (keyword === "flowchart" || keyword === "graph") return { kind: "flowchart", horizontal: direction(first.split(/\s+/)[1]) };
  if (keyword === "sequencediagram") return { kind: "sequence", horizontal: false };
  if (keyword.startsWith("statediagram")) return { kind: "state", horizontal: direction(inner) };
  if (keyword === "classdiagram") return { kind: "class", horizontal: direction(inner) };
  if (keyword === "erdiagram") return { kind: "er", horizontal: direction(inner) };
  if (keyword === "gantt") return { kind: "gantt", horizontal: false };
  return { kind: "other", horizontal: false };
}


function knobsFor(code: string): Knobs {
  const { kind, horizontal } = diagramKind(code);
  switch (kind) {
    case "flowchart":
      return graphKnobs("flowchart", horizontal);
    case "state":
    case "class":
      // mermaid 11.12 の dagre は、状態遷移図・クラス図でも flowchart の間隔の設定を優先して使う
      return graphKnobs("flowchart", horizontal);
    case "er":
      // mermaid 11.12 の ER 図の既定値
      return graphKnobs("er", horizontal, 140, 80);
    case "sequence":
      // 登場人物の間（actorMargin）が幅、メッセージの間（boxMargin）が高さに効く（描画チェックで確認）
      return { w: knob("sequence", "actorMargin", 50, 10), h: knob("sequence", "boxMargin", 10, 4, 200) };
    case "gantt":
      // ガントは幅（useWidth）と棒の太さ（barHeight）を、文字の大きさに合わせた既定値に対する倍率で変える。
      // 実際の値は描画側（renderer.ts の ganttConfig）で決める
      return { w: knob("gantt", "widthScale", 1, 0.5, 25), h: knob("gantt", "barScale", 1, 0.6, 5) };
    default:
      return {};
  }
}

/** 広げない（既定の間隔の）設定 */
export function baseSpacing(code: string): Spacing {
  const k = knobsFor(code);
  return { w: k.w?.base, h: k.h?.base };
}

/** Spacing を mermaid.initialize に渡す設定に直す（セクションごとのオブジェクト） */
export function spacingConfig(code: string, spacing: Spacing): Record<string, Record<string, number>> {
  const k = knobsFor(code);
  const config: Record<string, Record<string, number>> = {};
  for (const [dim, def] of [["w", k.w], ["h", k.h]] as const) {
    const value = spacing[dim];
    if (!def || value === undefined) continue;
    // 倍率（〜Scale）は小数のまま、間隔（px）は整数にする
    config[def.section] = { ...config[def.section], [def.key]: def.key.endsWith("Scale") ? value : Math.round(value) };
  }
  return config;
}

/** 文字を何倍に縮めれば枠に収まるか（1 なら縮めない） */
export function shrinkToFit(natural: Size, target: Size): number {
  return Math.min(1, target.width / natural.width, target.height / natural.height);
}

/** どれだけ枠を埋めているか。縮めずに済むことを最優先し、次に幅、その次に高さを埋めているものを良しとする */
function score(natural: Size, target: Size): number {
  const s = shrinkToFit(natural, target);
  const fillW = (natural.width * s) / target.width;
  const fillH = (natural.height * s) / target.height;
  return Math.round(s * 50) * 100 + fillW * 2 + fillH;
}

/** 描く回数の上限 */
const MAX_RENDERS = 4;

/**
 * target（px）に近い大きさになるよう、間隔を変えて描く。
 * 間隔と図の大きさはほぼ比例するので、2 回描いて傾きを求め、狙いの間隔を決める。ずれれば 1 回直す。
 * 広げる設定の無い向き（縦一列の図の幅など）は広がらず、余りは余白になる
 */
export async function renderToFill<T extends Size>(
  code: string,
  render: (spacing: Spacing) => Promise<T>,
  target: Size,
): Promise<{ rendered: T; spacing: Spacing }> {
  const k = knobsFor(code);
  const base = baseSpacing(code);
  const first = await render(base);
  if (!k.w && !k.h) return { rendered: first, spacing: base };

  let best = { rendered: first, spacing: base };
  let prev = { rendered: first, spacing: base };
  // 2 回目は間隔を 2 倍にして傾きを測る
  let next: Spacing = { w: k.w && k.w.base * 2, h: k.h && k.h.base * 2 };
  for (let i = 1; i < MAX_RENDERS; i++) {
    let rendered: T;
    try {
      rendered = await render(next);
    } catch {
      // 間隔を詰めすぎると mermaid が線のラベルを置けずに失敗することがある。それまでの最良を使う
      break;
    }
    const current = { rendered, spacing: next };
    if (score(rendered, target) > score(best.rendered, target)) best = current;

    // 向きごとに、間隔を変えたときの大きさの傾きを求める
    const dims = (
      [
        ["w", "width", k.w],
        ["h", "height", k.h],
      ] as const
    ).map(([dim, size, def]) => {
      const a = prev.spacing[dim];
      const b = current.spacing[dim];
      const slope = !def || a === undefined || b === undefined || a === b ? 0 : (rendered[size] - prev.rendered[size]) / (b - a);
      return { dim, size, def, b: b ?? 0, slope, now: rendered[size] };
    });
    // 間隔では縮められない向き（設定が無い・効かない・最小まで詰めた）の大きさで、図を縮める倍率が決まる
    let shrink = 1;
    for (const d of dims) {
      const smallest =
        !d.def || d.slope < 0.05 ? d.now : d.now + (Math.max(d.def.min, d.b + (target[d.size] - d.now) / d.slope) - d.b) * d.slope;
      shrink = Math.min(shrink, target[d.size] / smallest);
    }
    // 縮めるなら、縮めたあとに枠いっぱいになるよう、ほかの向きは狙いを大きくする（横長の枠なら横に広げる）
    const aim: Spacing = {};
    let moved = false;
    for (const d of dims) {
      if (!d.def) continue;
      // 間隔を変えても大きさが変わらない向き（縦一列の図の幅など）は、既定の間隔に戻す
      if (d.slope < 0.05) {
        aim[d.dim] = d.def.base;
        continue;
      }
      const value = Math.min(d.def.max, Math.max(d.def.min, d.b + (target[d.size] / shrink - d.now) / d.slope));
      aim[d.dim] = value;
      if (Math.abs(value - d.b) > Math.max(1, d.b * 0.02)) moved = true;
    }
    if (!moved) break;
    prev = current;
    next = aim;
  }
  return best;
}
