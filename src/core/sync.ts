/**
 * ノートのブロックとスライド上の図を突き合わせ、何をするかを決める（設計書「同期ルール」）。
 * Office に依存しない純粋関数。
 */
import type { MermaidBlock, ParseResult } from "./notesParser";
import type { Box } from "./layout";

/** スライド上にある、このアドインが管理する図（タグ M4P_ID を持つ図形） */
export interface ExistingDiagram extends Box {
  shapeId: string;
  blockId: string;
  hash: string;
  /** タグ M4P_THEME。検証版で作った図には無いので default とみなす */
  theme: string;
  /** 挿入したあとに利用者が大きさを変えた。変えた大きさのまま作り直す */
  resized?: boolean;
  /** タグ M4P_FONT_PT（指定した文字の大きさ）。1.1 以前に作った図には無い */
  fontPt?: number | null;
  /** グループの中にある図なら、そのグループの図形 id */
  groupId?: string;
}

export type Action =
  | { kind: "insert"; block: MermaidBlock }
  | { kind: "replace"; block: MermaidBlock; diagram: ExistingDiagram }
  | { kind: "upToDate"; block: MermaidBlock; diagram: ExistingDiagram }
  /** ノートに対応するブロックが無い図。削除はせず知らせるだけ */
  | { kind: "orphan"; diagram: ExistingDiagram }
  | { kind: "error"; line: number; message: string; blockId?: string };

export interface PlanInput {
  parsed: ParseResult;
  /** ブロック id → ソースのハッシュ */
  hashes: Map<string, string>;
  diagrams: ExistingDiagram[];
  /** themeKey() の値。図のタグ M4P_THEME と比べる */
  theme: string;
  /** true ならハッシュが一致していても作り直す */
  force: boolean;
  /** パネルの「文字の大きさ」（pt）。ブロックに fontsize= があればそちらを使う */
  fontPt?: number;
}

/** ブロックの文字の大きさ（pt）。fontsize= が無ければパネルの設定 */
export function blockFontPt(block: MermaidBlock, panelFontPt: number): number {
  return block.fontPt ?? panelFontPt;
}

export function planSlide({ parsed, hashes, diagrams, theme, force, fontPt }: PlanInput): Action[] {
  const actions: Action[] = [];

  // id が重複していると、どの図がどのブロックか決められないため、スライドごと処理しない
  const duplicate = parsed.errors.some((e) => e.message.includes("重複"));
  for (const e of parsed.errors) actions.push({ kind: "error", line: e.line, message: e.message });
  if (duplicate) return actions;

  const erroredLines = new Set(parsed.errors.map((e) => e.line));
  const blockIds = new Set(parsed.blocks.map((b) => b.id));

  for (const block of parsed.blocks) {
    // 図をコピーして同じスライドに貼ると、同じ id の図が複数になる。どれもこのブロックの図として扱い、全部更新する
    const copies = diagrams.filter((d) => d.blockId === block.id);
    // 空ブロックなど記法エラーのあるブロックは、図があっても触らない
    if (erroredLines.has(block.line)) continue;
    if (copies.length === 0) {
      actions.push({ kind: "insert", block });
      continue;
    }
    for (const diagram of copies) {
      // 文字の大きさの記録が無い図（1.1 以前）は、文字の大きさの設定では作り直さない
      const sameFont = diagram.fontPt == null || fontPt === undefined || diagram.fontPt === blockFontPt(block, fontPt);
      const same =
        diagram.hash === hashes.get(block.id) &&
        (diagram.theme || "default") === theme &&
        !diagram.resized &&
        sameFont;
      actions.push(same && !force ? { kind: "upToDate", block, diagram } : { kind: "replace", block, diagram });
    }
  }

  for (const d of diagrams) {
    if (!blockIds.has(d.blockId)) actions.push({ kind: "orphan", diagram: d });
  }
  return actions;
}

/**
 * mermaid の構文エラーに含まれる「line N」（ブロック内の行番号）を、ノート全体での行番号に直す。
 * 行番号が読み取れなければ開始フェンスの行を返す。
 */
export function notesLineOfError(block: MermaidBlock, message: string): number {
  const m = /line (\d+)/i.exec(message);
  return m ? block.line + Number(m[1]) : block.line;
}
