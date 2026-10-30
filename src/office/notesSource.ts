/**
 * ノートの読み取り（Office.js 側）。
 * 公式 API が無いため、getFileAsync で pptx 全体を取って解析する。
 * 実機（Windows 版 2610）で、保存前のノート編集も反映されることを確認済み。
 * 非公式の slide.notesSlide は存在しなかったため使わない。
 */
import { readPresentationInfo, sldIdFromOfficeSlideId, type SlideNotes } from "../core/pptxNotes";
import { CancelledError, type SyncProgress } from "../core/progress";
import type { TemplateColors } from "../core/theme";

export interface CurrentSlide {
  id: string;
  index: number;
}

/** 表示順のスライド id（Office.js の Slide.id） */
export async function getSlideIds(): Promise<string[]> {
  return PowerPoint.run(async (context) => {
    const all = context.presentation.slides;
    all.load("items/id");
    await context.sync();
    return all.items.map((s) => s.id);
  });
}

export async function getCurrentSlide(): Promise<CurrentSlide> {
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedSlides();
    selected.load("items/id");
    const all = context.presentation.slides;
    all.load("items/id");
    await context.sync();
    if (selected.items.length === 0) throw new Error("スライドが選択されていません");
    const id = selected.items[0].id;
    return { id, index: all.items.findIndex((s) => s.id === id) };
  });
}

/**
 * プレゼンテーション全体を pptx（Compressed）として取得する。onSlice には受け取った分割数を渡す。
 * signal で中止すると、次の分割を取りに行かずに CancelledError で終わる
 */
export function readPresentationBytes(
  sliceSize = 4 * 1024 * 1024,
  onSlice?: (done: number, total: number) => void,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(`getFileAsync 失敗: ${result.error.message}`));
        return;
      }
      const file = result.value;
      const chunks: Uint8Array[] = [];
      let received = 0;
      const finish = (err?: Error) => {
        // 閉じないと次回の getFileAsync が失敗する
        file.closeAsync(() => {
          if (err) {
            reject(err);
            return;
          }
          const total = chunks.reduce((n, c) => n + c.length, 0);
          const bytes = new Uint8Array(total);
          let offset = 0;
          for (const c of chunks) {
            bytes.set(c, offset);
            offset += c.length;
          }
          resolve(bytes);
        });
      };
      const next = (i: number) => {
        if (signal?.aborted) {
          finish(new CancelledError());
          return;
        }
        file.getSliceAsync(i, (sliceResult) => {
          if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
            finish(new Error(`getSliceAsync 失敗: ${sliceResult.error.message}`));
            return;
          }
          chunks[i] = new Uint8Array(sliceResult.value.data as number[]);
          received++;
          onSlice?.(received, file.sliceCount);
          if (received === file.sliceCount) finish();
          else next(i + 1);
        });
      };
      onSlice?.(0, file.sliceCount);
      next(0);
    });
  });
}

export interface PresentationNotes {
  slideSize: { width: number; height: number };
  /** テーマの配色（「テンプレートの色」で使う）。読めなければ null */
  themeColors: TemplateColors | null;
  /** Office.js の Slide.id から、そのスライドのノートを引く */
  notesOf(slide: CurrentSlide): string | null;
}

/** pptx 全体を取得し、全スライドのノートとスライドの大きさを読む */
export async function readPresentationNotes(
  onProgress?: (p: SyncProgress) => void,
  signal?: AbortSignal,
): Promise<PresentationNotes> {
  const bytes = await readPresentationBytes(
    undefined,
    (done, total) => onProgress?.({ phase: "read", done, total }),
    signal,
  );
  if (signal?.aborted) throw new CancelledError();
  onProgress?.({ phase: "parse" });
  const info = await readPresentationInfo(bytes);
  const bySldId = new Map<number, SlideNotes>(info.slides.map((s) => [s.sldId, s]));
  return {
    slideSize: info.slideSize,
    themeColors: info.themeColors,
    notesOf(slide) {
      // Slide.id は "259#878835286" の形で、# の前が pptx 内の sldId（実機で確認）。取れなければ表示順で対応付ける
      const sldId = sldIdFromOfficeSlideId(slide.id);
      const match = (sldId !== null ? bySldId.get(sldId) : undefined) ?? info.slides[slide.index];
      return match?.notesText ?? null;
    },
  };
}
