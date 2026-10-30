/**
 * 図の生成・更新の進み具合。大きいファイルでは pptx の読み込みに時間がかかるため、パネルに表示する。
 */

export type SyncProgress =
  /** pptx を分割して受け取っている（done / total 個目） */
  | { phase: "read"; done: number; total: number }
  /** 受け取った pptx からノートを読んでいる */
  | { phase: "parse" }
  /** スライドごとに図を作っている（done / total 枚目まで終わった） */
  | { phase: "slides"; done: number; total: number };

/** 利用者がキャンセルしたことを表す */
export class CancelledError extends Error {
  constructor() {
    super("キャンセルしました");
    this.name = "CancelledError";
  }
}

export interface ProgressView {
  label: string;
  /** 0～1。進み具合が分からない段階は null（バーを流れる表示にする） */
  value: number | null;
}

export function progressView(p: SyncProgress): ProgressView {
  switch (p.phase) {
    case "read":
      return { label: `ファイルを読み込み中（${p.done}/${p.total}）`, value: p.total > 0 ? p.done / p.total : null };
    case "parse":
      return { label: "ノートを読み取り中", value: null };
    case "slides":
      return {
        label: p.total > 1 ? `図を生成中（スライド ${p.done}/${p.total}）` : "図を生成中",
        value: p.total > 0 ? p.done / p.total : null,
      };
  }
}

/**
 * 一定時間たっても終わらないときだけ表示する。すぐ終わる処理でちらつかせないため。
 * start() から delayMs 経つと show を呼び、stop() で（表示していれば）hide を呼ぶ。
 */
export function delayedIndicator(delayMs: number, show: () => void, hide: () => void) {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let shown = false;
  return {
    start() {
      if (timer !== null || shown) return;
      timer = setTimeout(() => {
        timer = null;
        shown = true;
        show();
      }, delayMs);
    },
    stop() {
      if (timer !== null) clearTimeout(timer);
      timer = null;
      if (shown) hide();
      shown = false;
    },
  };
}
