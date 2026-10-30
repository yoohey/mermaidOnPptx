/**
 * 図の配置とサイズを決める。単位は pt（スライド座標）。
 */

export interface Size {
  width: number;
  height: number;
}

export interface Box extends Size {
  left: number;
  top: number;
}

/** 新規の図が使ってよい範囲（スライドに対する割合） */
const MAX_W_RATIO = 0.8;
const MAX_H_RATIO = 0.7;
/** mermaid の SVG は px（96dpi）、スライドは pt */
export const PX_TO_PT = 0.75;

/**
 * 新規の図を置く位置を決める。count 個の新しい図を横に並べ、index 番目の枠に収める。
 * 文字を指定の大きさで見せるため拡大はせず、枠より大きい図だけ縮める。広げたいときは図の枠を広げて作り直す。
 */
export function placeNew(natural: Size, slide: Size, index = 0, count = 1): Box {
  const cellW = (slide.width * MAX_W_RATIO) / count;
  const maxH = slide.height * MAX_H_RATIO;
  const w = natural.width * PX_TO_PT;
  const h = natural.height * PX_TO_PT;
  const scale = Math.min(1, cellW / w, maxH / h);
  const width = Math.round(w * scale);
  const height = Math.round(h * scale);
  const cellLeft = (slide.width * (1 - MAX_W_RATIO)) / 2 + cellW * index;
  return {
    left: Math.round(cellLeft + (cellW - width) / 2),
    top: Math.round((slide.height - height) / 2),
    width,
    height,
  };
}

/** 縦横比を保って box に収めたときの倍率（等倍 = 1）。余りは余白になる */
export function fitScale(natural: Size, box: Size): number {
  return Math.min(box.width / (natural.width * PX_TO_PT), box.height / (natural.height * PX_TO_PT));
}

/** 図の文字の大きさ（pt）の既定値と、パネルで選べる値 */
export const DEFAULT_FONT_PT = 14;
export const FONT_PT_CHOICES = [10, 12, 14, 16, 18, 20, 24, 28, 32] as const;
/** ブロックの fontsize= で指定できる範囲 */
export const MIN_FONT_PT = 6;
export const MAX_FONT_PT = 72;

/** 枠（pt）を、等倍で描いたときの SVG の大きさ（px）に直す */
export function boxToPx(box: Size): Size {
  return { width: box.width / PX_TO_PT, height: box.height / PX_TO_PT };
}

/**
 * viewBox を box の縦横比まで広げ、図が中央に来るようにする。
 * PowerPoint に box の大きさで挿入しても図が歪まず、余りは余白になる。
 */
export function viewBoxForBox(viewBox: [number, number, number, number], box: Size): [number, number, number, number] {
  const [x, y, w, h] = viewBox;
  const boxRatio = box.width / box.height;
  if (w / h < boxRatio) {
    const newW = h * boxRatio;
    return [x - (newW - w) / 2, y, newW, h];
  }
  const newH = w / boxRatio;
  return [x, y - (newH - h) / 2, w, newH];
}

/** タグ M4P_BOX の値（挿入したときの幅と高さ、pt） */
export function formatBox(box: Size): string {
  return `${box.width.toFixed(1)},${box.height.toFixed(1)}`;
}

export function parseBox(value: string): Size | null {
  const [width, height] = value.split(",").map(Number);
  return width > 0 && height > 0 ? { width, height } : null;
}

/** 挿入したあとに利用者が図の大きさを変えたか。1pt 未満の差は丸めの誤差とみなす */
export function isResized(inserted: Size | null, current: Size): boolean {
  if (!inserted) return false;
  return Math.abs(inserted.width - current.width) >= 1 || Math.abs(inserted.height - current.height) >= 1;
}
