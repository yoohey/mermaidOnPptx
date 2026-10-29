/**
 * ノートに合わせて図を生成・更新する一連の処理。タスクペインとリボンの両方から呼ぶ。
 */
import { parseMermaidBlocks } from "../core/notesParser";
import { sourceHash } from "../core/hash";
import { placeNew, placeReplacing } from "../core/layout";
import { notesLineOfError, planSlide } from "../core/sync";
import { getCurrentSlide, getSlideIds, readPresentationNotes, type CurrentSlide } from "../office/notesSource";
import { readDiagrams } from "../office/diagrams";
import { renderMermaid, MERMAID_VERSION, type RenderedSvg } from "../office/renderer";
import { deleteShape, insertDiagram, selectSlide } from "../office/slideWriter";
import { getCustomColors, getTheme } from "../office/settings";
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
}

export interface SyncOptions {
  scope: "current" | "all";
  force: boolean;
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

  const actions = planSlide({ parsed, hashes, diagrams, theme: key, force });
  const newCount = actions.filter((a) => a.kind === "insert").length;
  let newIndex = 0;
  const items: ItemReport[] = [];
  // 同じ id の図が複数ある（コピーした）ときは、描画は1回だけにする
  const renders = new Map<string, Promise<RenderedSvg>>();
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
    if (!renders.has(block.id)) renders.set(block.id, renderMermaid(block.code, theme, palette));
    let rendered;
    try {
      rendered = await renders.get(block.id)!;
    } catch (e) {
      // 構文エラーのときは既にある図に触らない。コピーが複数あってもエラーは1回だけ出す
      if (reportedErrors.has(block.id)) continue;
      reportedErrors.add(block.id);
      const message = errorMessage(e);
      items.push({ status: "error", blockId: block.id, line: notesLineOfError(block, message), message });
      continue;
    }
    const box =
      action.kind === "replace"
        ? placeReplacing(action.diagram, rendered)
        : placeNew(rendered, slideSize, slot, newCount);
    try {
      await insertDiagram({
        slideId: slide.id,
        blockId: block.id,
        source: block.code,
        hash: hashes.get(block.id)!,
        theme: key,
        mermaidVersion: MERMAID_VERSION,
        svg: rendered.svg,
        box,
      });
      // 新しい図を入れてから古い図を消す。挿入に失敗しても古い図は残る
      if (action.kind === "replace") await deleteShape(slide.id, action.diagram.shapeId);
      items.push({ status: action.kind === "replace" ? "replaced" : "inserted", blockId: block.id, svg: rendered.svg });
    } catch (e) {
      items.push({ status: "error", blockId: block.id, line: block.line, message: errorMessage(e) });
    }
  }
  return items;
}

export async function runSync({ scope, force }: SyncOptions): Promise<SyncReport> {
  const current = await getCurrentSlide();
  const notes = await readPresentationNotes();
  const targets: CurrentSlide[] =
    scope === "current" ? [current] : (await getSlideIds()).map((id, index) => ({ id, index }));

  const slides: SlideReport[] = [];
  try {
    for (const slide of targets) {
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
  return { scope, slides, counts };
}
