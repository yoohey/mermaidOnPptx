/**
 * プレゼン単位の設定。Office.context.document.settings は pptx の中に保存されるため、
 * ファイルと一緒に持ち運べる。
 */
import {
  DEFAULT_CUSTOM_COLORS,
  DEFAULT_THEME,
  isCustomColors,
  isTheme,
  type CustomColors,
  type Theme,
} from "../core/theme";
import { DEFAULT_FONT_PT, MAX_FONT_PT, MIN_FONT_PT } from "../core/layout";

const THEME_KEY = "m4p.theme";
const CUSTOM_COLORS_KEY = "m4p.customColors";
const FONT_PT_KEY = "m4p.fontPt";

export function getTheme(): Theme {
  const value = Office.context.document.settings.get(THEME_KEY);
  return isTheme(value) ? value : DEFAULT_THEME;
}

export function setTheme(theme: Theme): Promise<void> {
  Office.context.document.settings.set(THEME_KEY, theme);
  return save();
}

export function getCustomColors(): CustomColors {
  const value = Office.context.document.settings.get(CUSTOM_COLORS_KEY);
  return isCustomColors(value) ? value : DEFAULT_CUSTOM_COLORS;
}

/** カスタムの色をまだ一度も保存していなければ false */
export function hasCustomColors(): boolean {
  return isCustomColors(Office.context.document.settings.get(CUSTOM_COLORS_KEY));
}

export function setCustomColors(colors: CustomColors): Promise<void> {
  Office.context.document.settings.set(CUSTOM_COLORS_KEY, colors);
  return save();
}

/** 図の文字の大きさ（pt）。ブロックの fontsize= があればそちらが優先 */
export function getFontPt(): number {
  const value = Office.context.document.settings.get(FONT_PT_KEY);
  return typeof value === "number" && value >= MIN_FONT_PT && value <= MAX_FONT_PT ? value : DEFAULT_FONT_PT;
}

export function setFontPt(pt: number): Promise<void> {
  Office.context.document.settings.set(FONT_PT_KEY, pt);
  return save();
}

function save(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(`設定の保存に失敗: ${result.error.message}`));
    });
  });
}
