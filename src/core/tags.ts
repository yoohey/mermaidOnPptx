/** 図形に付けるタグのキー。PowerPoint はタグのキーを大文字で保持するため大文字で定義する。 */
export const TAG = {
  ID: "M4P_ID",
  SRC: "M4P_SRC",
  HASH: "M4P_HASH",
  VER: "M4P_VER",
  MERMAID: "M4P_MERMAID",
  THEME: "M4P_THEME",
  /** 挿入したときの幅と高さ（pt）。利用者が大きさを変えたかの判断に使う */
  BOX: "M4P_BOX",
  /** 指定した文字の大きさ（pt）。パネルの「文字の大きさ」やブロックの fontsize= を変えたかの判断に使う */
  FONT_PT: "M4P_FONT_PT",
  /** スライド上での実際の文字の大きさ（pt）。図が枠に収まらず縮めたときは FONT_PT より小さい */
  TEXT_PT: "M4P_TEXT_PT",
  /**
   * 箱や線の配置を決めた文字の大きさ（pt）。最初に入れたときの文字の大きさで、作り直しても変えない。
   * 文字の大きさを変えて作り直しても、図形の大きさが変わらないようにするため。1.2 以前の図には無く、FONT_PT で代える
   */
  LAYOUT_PT: "M4P_LAYOUT_PT",
} as const;

/** タグのデータ形式の版 */
export const DATA_FORMAT_VERSION = "1";
