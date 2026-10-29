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
/** 小さい図を拡大する上限。これ以上大きくすると文字だけが極端に大きくなる */
const MAX_UPSCALE = 2;
/** mermaid の SVG は px（96dpi）、スライドは pt */
const PX_TO_PT = 0.75;

/**
 * 新規の図を置く位置を決める。count 個の新しい図を横に並べ、index 番目の枠に収める。
 * 枠より小さい図は MAX_UPSCALE 倍まで拡大し、大きい図は縮める。
 */
export function placeNew(natural: Size, slide: Size, index = 0, count = 1): Box {
  const cellW = (slide.width * MAX_W_RATIO) / count;
  const maxH = slide.height * MAX_H_RATIO;
  const w = natural.width * PX_TO_PT;
  const h = natural.height * PX_TO_PT;
  const scale = Math.min(MAX_UPSCALE, cellW / w, maxH / h);
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

/** 既存の図を置き換えるときは、左上と横幅を引き継ぎ、高さは新しい図の縦横比で決める */
export function placeReplacing(old: Box, natural: Size): Box {
  return {
    left: old.left,
    top: old.top,
    width: old.width,
    height: Math.round((old.width * natural.height) / natural.width),
  };
}
