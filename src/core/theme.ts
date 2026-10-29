/**
 * タスクペインで選べるテーマ。modern / mono / navy はこのアドイン独自のスタイルで、
 * mermaid の base テーマに色（themeVariables）を指定して作る。template はプレゼンのテーマの色から、
 * custom はパネルで選んだ色から作る。
 * 残りは mermaid の標準テーマ。
 */
export const THEMES = ["template", "modern", "mono", "navy", "default", "neutral", "dark", "forest", "custom"] as const;
export type Theme = (typeof THEMES)[number];

export const DEFAULT_THEME: Theme = "modern";

export const THEME_LABELS: Record<Theme, string> = {
  modern: "モダン（青）",
  mono: "モノトーン",
  navy: "ネイビー",
  template: "テンプレートの色",
  custom: "カスタム（色を選ぶ）",
  default: "mermaid 標準",
  neutral: "neutral",
  dark: "dark",
  forest: "forest",
};

type ThemeVariables = Record<string, string>;

const COMMON: ThemeVariables = { edgeLabelBackground: "#FFFFFF", fontSize: "16px" };

/** 円グラフの色。base テーマのままだと primaryColor から作った淡い色ばかりになり、区別しにくい */
function pieColors(colors: string[]): ThemeVariables {
  return Object.fromEntries(colors.map((c, i) => [`pie${i + 1}`, c]));
}

/** 独自スタイルの色。キーは mermaid の themeVariables */
export const CUSTOM_THEMES: Partial<Record<Theme, ThemeVariables>> = {
  modern: {
    ...COMMON,
    primaryColor: "#E8F1FB",
    primaryBorderColor: "#2F6FB3",
    primaryTextColor: "#1F2937",
    secondaryColor: "#F1F5F9",
    tertiaryColor: "#FFFFFF",
    lineColor: "#475569",
    textColor: "#1F2937",
    actorBkg: "#DBEAFE",
    actorBorder: "#2F6FB3",
    actorTextColor: "#1E3A8A",
    actorLineColor: "#94A3B8",
    signalColor: "#334155",
    signalTextColor: "#1F2937",
    noteBkgColor: "#FFF7D6",
    noteBorderColor: "#E0B100",
    ...pieColors(["#2F6FB3", "#5B9BD5", "#1E3A8A", "#0E7490", "#3B82F6", "#475569", "#0369A1", "#64748B"]),
    pieSectionTextColor: "#FFFFFF",
  },
  mono: {
    ...COMMON,
    primaryColor: "#FFFFFF",
    primaryBorderColor: "#334155",
    primaryTextColor: "#111827",
    secondaryColor: "#F3F4F6",
    tertiaryColor: "#FFFFFF",
    lineColor: "#4B5563",
    textColor: "#111827",
    actorBkg: "#F3F4F6",
    actorBorder: "#334155",
    actorTextColor: "#111827",
    actorLineColor: "#9CA3AF",
    signalColor: "#374151",
    signalTextColor: "#111827",
    noteBkgColor: "#F9FAFB",
    noteBorderColor: "#9CA3AF",
    ...pieColors(["#1F2937", "#6B7280", "#374151", "#8B939E", "#4B5563", "#111827", "#9CA3AF", "#525A66"]),
    pieSectionTextColor: "#FFFFFF",
  },
  navy: {
    ...COMMON,
    primaryColor: "#EEF2F7",
    primaryBorderColor: "#1E3A5F",
    primaryTextColor: "#1E293B",
    secondaryColor: "#FDEBD8",
    tertiaryColor: "#FFFFFF",
    lineColor: "#1E3A5F",
    textColor: "#1E293B",
    actorBkg: "#DCE3EC",
    actorBorder: "#1E3A5F",
    actorTextColor: "#1E3A5F",
    actorLineColor: "#94A3B8",
    signalColor: "#1E3A5F",
    signalTextColor: "#1E293B",
    noteBkgColor: "#FDEBD8",
    noteBorderColor: "#E07B24",
    ...pieColors(["#1E3A5F", "#E07B24", "#5B7DA3", "#B45F1A", "#334155", "#C2410C", "#475569", "#7C94B4"]),
    pieSectionTextColor: "#FFFFFF",
  },
};

/** プレゼンのテーマの配色（<a:clrScheme>）。色は "#RRGGBB" */
export interface TemplateColors {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  /** accent1〜accent6 */
  accents: string[];
}

function rgb(hex: string): [number, number, number] {
  const n = parseInt(hex.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

/** color を white に向かって ratio（0〜1）だけ薄める */
function tint(color: string, ratio: number): string {
  const mixed = rgb(color).map((c) => Math.round(c + (255 - c) * ratio));
  return "#" + mixed.map((c) => c.toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** 相対輝度（0 が黒、1 が白）。WCAG の式 */
export function luminance(color: string): number {
  const [r, g, b] = rgb(color).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * テーマの配色から themeVariables を作る。図形の塗りはアクセント1を薄めた色、枠はアクセント1、
 * 文字は「濃色1」と「濃色2」のうち暗い方。円グラフはアクセント1〜6をそのまま使う。
 */
export function templateThemeVariables(c: TemplateColors): ThemeVariables {
  const [a1, a2 = a1, , a4 = a1] = c.accents;
  const text = luminance(c.dk1) <= luminance(c.dk2) ? c.dk1 : c.dk2;
  // アクセントが明るい配色では、円グラフの白文字が読めないので文字色を濃くする
  const avgAccent = c.accents.reduce((n, a) => n + luminance(a), 0) / Math.max(c.accents.length, 1);
  return {
    ...COMMON,
    primaryColor: tint(a1, 0.85),
    primaryBorderColor: a1,
    primaryTextColor: text,
    secondaryColor: tint(a2, 0.85),
    tertiaryColor: "#FFFFFF",
    lineColor: tint(text, 0.25),
    textColor: text,
    actorBkg: tint(a1, 0.75),
    actorBorder: a1,
    actorTextColor: text,
    actorLineColor: tint(text, 0.55),
    signalColor: tint(text, 0.15),
    signalTextColor: text,
    noteBkgColor: tint(a4, 0.85),
    noteBorderColor: a4,
    ...pieColors(c.accents),
    pieSectionTextColor: avgAccent > 0.45 ? text : "#FFFFFF",
  };
}

/** パネルで選ぶ色（カスタム）。"#RRGGBB" */
export interface CustomColors {
  main: string;
  sub: string;
  text: string;
}

export const DEFAULT_CUSTOM_COLORS: CustomColors = { main: "#2F6FB3", sub: "#E07B24", text: "#1F2937" };

export function isCustomColors(value: unknown): value is CustomColors {
  const hex = (v: unknown) => typeof v === "string" && /^#[0-9A-Fa-f]{6}$/.test(v);
  const c = value as CustomColors | null;
  return !!c && typeof c === "object" && hex(c.main) && hex(c.sub) && hex(c.text);
}

/** color を black に向かって ratio（0〜1）だけ濃くする */
function shade(color: string, ratio: number): string {
  const mixed = rgb(color).map((c) => Math.round(c * (1 - ratio)));
  return "#" + mixed.map((c) => c.toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** カスタムの3色を配色の形に広げる。円グラフ用のアクセント3〜6はメイン色・サブ色の濃淡で作る */
export function customPalette(c: CustomColors): TemplateColors {
  return {
    dk1: c.text,
    lt1: "#FFFFFF",
    dk2: c.text,
    lt2: "#F2F2F2",
    accents: [c.main, c.sub, shade(c.main, 0.35), shade(c.sub, 0.35), tint(c.main, 0.4), tint(c.sub, 0.4)],
  };
}

/** テンプレートの配色からカスタムの初期値を作る（パネルの「テンプレートの色から始める」） */
export function customFromTemplate(c: TemplateColors): CustomColors {
  const text = luminance(c.dk1) <= luminance(c.dk2) ? c.dk1 : c.dk2;
  return { main: c.accents[0], sub: c.accents[1] ?? c.accents[0], text };
}

/**
 * mermaid 標準テーマのおおよその色（カスタムの初期値用）。標準テーマは themeVariables を持たないため、
 * mermaid の各テーマの枠線・塗りの色に近い値を手で決めている。
 */
const STANDARD_THEME_COLORS: Partial<Record<Theme, CustomColors>> = {
  default: { main: "#9370DB", sub: "#D4B000", text: "#333333" },
  neutral: { main: "#999999", sub: "#666666", text: "#333333" },
  dark: { main: "#81B1DB", sub: "#C9A15A", text: "#333333" },
  forest: { main: "#13540C", sub: "#6EAA49", text: "#333333" },
};

/**
 * 図のタグ M4P_THEME（themeKey の値）から、その図の色合いに近いカスタムの初期値を作る。
 * 読み取れなければ null。タグが空の図（検証版で作った図）は mermaid 標準（default）とみなす。
 */
export function customFromThemeKey(key: string): CustomColors | null {
  const m = /^(template|custom):(.+)$/.exec(key);
  if (m) {
    const [dk1, dk2, ...accents] = m[2].split(",");
    if (!dk1 || !dk2 || accents.length < 2) return null;
    return customFromTemplate({ dk1, dk2, lt1: "#FFFFFF", lt2: "#F2F2F2", accents });
  }
  const theme = key || "default";
  if (!isTheme(theme)) return null;
  const custom = CUSTOM_THEMES[theme];
  if (custom) return { main: custom.primaryBorderColor, sub: custom.pie2, text: custom.textColor };
  return STANDARD_THEME_COLORS[theme] ?? null;
}

/** template / custom で使う配色を決める。ほかのテーマでは null */
export function paletteFor(
  theme: Theme,
  templateColors: TemplateColors | null,
  customColors: CustomColors,
): TemplateColors | null {
  if (theme === "template") return templateColors;
  if (theme === "custom") return customPalette(customColors);
  return null;
}

/**
 * mermaid.initialize に渡す theme と themeVariables。colors は template / custom の配色（paletteFor）。
 * template で配色が読めなかったとき（テーマの無いファイルなど）はモダンにする。
 */
export function mermaidThemeConfig(
  theme: Theme,
  colors?: TemplateColors | null,
): { theme: string; themeVariables?: ThemeVariables } {
  if (theme === "template" || theme === "custom") {
    return { theme: "base", themeVariables: colors ? templateThemeVariables(colors) : CUSTOM_THEMES.modern };
  }
  const custom = CUSTOM_THEMES[theme];
  return custom ? { theme: "base", themeVariables: custom } : { theme };
}

/**
 * 図のタグ M4P_THEME に入れ、更新の要否を判断する値。template / custom は配色ごと記録し、
 * 色を変えたら「全スライドを更新」で作り直されるようにする。
 */
export function themeKey(theme: Theme, colors?: TemplateColors | null): string {
  if (theme !== "template" && theme !== "custom") return theme;
  return colors ? `${theme}:${[colors.dk1, colors.dk2, ...colors.accents].join(",")}` : theme;
}

export function isTheme(value: unknown): value is Theme {
  return typeof value === "string" && (THEMES as readonly string[]).includes(value);
}
