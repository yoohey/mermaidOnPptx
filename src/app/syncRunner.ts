/**
 * ノートに合わせて図を生成・更新する一連の処理。タスクペインとリボンの両方から呼ぶ。
 */
import { parseMermaidBlocks } from "../core/notesParser";
import { sourceHash } from "../core/hash";
import { boxToPx, fitScale, placeNew, type Box } from "../core/layout";
import { renderToFill, type Spacing } from "../core/spacing";
import { blockFontPt, notesLineOfError, planSlide } from "../core/sync";
import { getCurrentSlide, getSlideIds, readPresentationNotes, type CurrentSlide } from "../office/notesSource";
import { readDiagrams } from "../office/diagrams";
import { renderMermaid, svgForBox, MERMAID_VERSION, type RenderedSvg } from "../office/renderer";
import { insertDiagram, replaceDiagram, selectSlide } from "../office/slideWriter";
import type { SyncProgress } from "../core/progress";
import { getCustomColors, getFontPt, getTheme } from "../office/settings";
import { paletteFor, themeKey, type TemplateColors } from "../core/theme";

export type ItemStatus = "inserted" | "replaced" | "upToDate" | "orphan" | "error";

export interface ItemReport {
  status: ItemStatus;
  blockId?: string;
  /** ノート全体での行番号（エラー時） */
  line?: number;
  message?: string;
  /** 生成した図（プレビュー用） */
  svg?: string;
}

export interface SlideReport {
  /** 1 始まりのスライド番号 */
  number: number;
  items: ItemReport[];
}

export interface SyncReport {
  scope: "current" | "all";
  slides: SlideReport[];
  counts: Record<ItemStatus, number>;
  /** 途中でキャンセルした（slides はそこまでの結果） */
  cancelled: boolean;
}

export interface SyncOptions {
  scope: "current" | "all";
  force: boolean;
  /** 進み具合の通知（パネルの進捗表示用） */
  onProgress?: (p: SyncProgress) => void;
  /** キャンセル。ファイルの読み込み中は途中で止め、図の生成中はそのスライドを終えてから止める */
  signal?: AbortSignal;
}

function errorMessage(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

async function syncSlide(
  slide: CurrentSlide,
  notes: string | null,
  slideSize: { width: number; height: number },
  templateColors: TemplateColors | null,
  force: boolean,
): Promise<ItemReport[]> {
  const theme = getTheme();
  const palette = paletteFor(theme, templateColors, getCustomColors());
  const key = themeKey(theme, palette);
  const parsed = parseMermaidBlocks(notes ?? "");
  const diagrams = await readDiagrams(slide.id);
  if (parsed.blocks.length === 0 && parsed.errors.length === 0 && diagrams.length === 0) return [];

  const hashes = new Map<string, string>();
  for (const block of parsed.blocks) hashes.set(block.id, await sourceHash(block.code));

  const panelFontPt = getFontPt();
  const actions = planSlide({ parsed, hashes, diagrams, theme: key, force, fontPt: panelFontPt });
  const newCount = actions.filter((a) => a.kind === "insert").length;
  let newIndex = 0;
  const items: ItemReport[] = [];
  // 同じ id の図が複数ある（コピーした）ときは、同じ設定での描画は1回だけにする
  const renders = new Map<string, Promise<RenderedSvg>>();
  const renderAt = (id: string, code: string, fontPt: number, spacing?: Spacing) => {
    const cacheKey = JSON.stringify([id, fontPt, spacing ?? null]);
    if (!renders.has(cacheKey)) renders.set(cacheKey, renderMermaid(code, theme, palette, { fontPt, spacing }));
    return renders.get(cacheKey)!;
  };
  const reportedErrors = new Set<string>();

  for (const action of actions) {
    if (action.kind === "error") {
      items.push({ status: "error", line: action.line, message: action.message });
      continue;
    }
    if (action.kind === "orphan") {
      items.push({ status: "orphan", blockId: action.diagram.blockId, message: "ノートに対応するブロックがありません" });
      continue;
    }
    if (action.kind === "upToDate") {
      items.push({ status: "upToDate", blockId: action.block.id });
      continue;
    }

    const { block } = action;
    const slot = action.kind === "insert" ? newIndex++ : 0;
    const fontPt = blockFontPt(block, panelFontPt);
    let rendered: RenderedSvg;
    let box: Box;
    try {
      if (action.kind === "replace") {
        // 置き換えでは、利用者が変えたかもしれない位置と大きさをそのまま使い、
        // 文字の大きさは変えずに間隔を広げて、図を枠いっぱいに描く
        const { left, top, width, height } = action.diagram;
        box = { left, top, width, height };
        rendered = (await renderToFill(block.code, (sp) => renderAt(block.id, block.code, fontPt, sp), boxToPx(box))).rendered;
      } else {
        rendered = await renderAt(block.id, block.code, fontPt);
        box = placeNew(rendered, slideSize, slot, newCount);
      }
    } catch (e) {
      // 構文エラーのときは既にある図に触らない。コピーが複数あってもエラーは1回だけ出す
      if (reportedErrors.has(block.id)) continue;
      reportedErrors.add(block.id);
      const message = errorMessage(e);
      items.push({ status: "error", blockId: block.id, line: notesLineOfError(block, message), message });
      continue;
    }
    try {
      const req = {
        slideId: slide.id,
        blockId: block.id,
        source: block.code,
        hash: hashes.get(block.id)!,
        theme: key,
        mermaidVersion: MERMAID_VERSION,
        svg: svgForBox(rendered, box),
        box,
        fontPt,
        // 図が枠に収まらず縮めたときは、指定より小さく見える
        textPt: fontPt * Math.min(1, fitScale(rendered, box)),
      };
      // 置き換えでは新しい図を入れてから古い図を消す。挿入に失敗しても古い図は残る
      if (action.kind === "replace") await replaceDiagram(req, action.diagram);
      else await insertDiagram(req);
      items.push({ status: action.kind === "replace" ? "replaced" : "inserted", blockId: block.id, svg: rendered.svg });
    } catch (e) {
      items.push({ status: "error", blockId: block.id, line: block.line, message: errorMessage(e) });
    }
  }
  return items;
}

export async function runSync({ scope, force, onProgress, signal }: SyncOptions): Promise<SyncReport> {
  const current = await getCurrentSlide();
  const notes = await readPresentationNotes(onProgress, signal);
  const targets: CurrentSlide[] =
    scope === "current" ? [current] : (await getSlideIds()).map((id, index) => ({ id, index }));

  const slides: SlideReport[] = [];
  let cancelled = false;
  try {
    for (const [done, slide] of targets.entries()) {
      // 1 枚の中で止めると古い図と新しい図が両方残りうるため、スライドの切れ目でだけ止める
      if (signal?.aborted) {
        cancelled = true;
        break;
      }
      onProgress?.({ phase: "slides", done, total: targets.length });
      let items: ItemReport[];
      try {
        items = await syncSlide(slide, notes.notesOf(slide), notes.slideSize, notes.themeColors, force);
      } catch (e) {
        items = [{ status: "error", message: errorMessage(e) }];
      }
      if (items.length > 0) slides.push({ number: slide.index + 1, items });
    }
  } finally {
    // 一括更新では挿入のたびに表示スライドが移るため、最初のスライドに戻す
    if (scope === "all") await selectSlide(current.id).catch(() => undefined);
  }

  const counts: SyncReport["counts"] = { inserted: 0, replaced: 0, upToDate: 0, orphan: 0, error: 0 };
  for (const s of slides) for (const item of s.items) counts[item.status]++;
  return { scope, slides, counts, cancelled };
}
