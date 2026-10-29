import { describe, expect, it } from "vitest";
import {
  CUSTOM_THEMES,
  DEFAULT_CUSTOM_COLORS,
  DEFAULT_THEME,
  customFromTemplate,
  customFromThemeKey,
  customPalette,
  isCustomColors,
  paletteFor,
  THEMES,
  THEME_LABELS,
  luminance,
  mermaidThemeConfig,
  themeKey,
  type TemplateColors,
} from "../src/core/theme";

const office2007: TemplateColors = {
  dk1: "#000000",
  lt1: "#FFFFFF",
  dk2: "#1F497D",
  lt2: "#EEECE1",
  accents: ["#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646"],
};

describe("mermaidThemeConfig", () => {
  it("独自スタイルは base テーマに色を指定する", () => {
    const c = mermaidThemeConfig("modern");
    expect(c.theme).toBe("base");
    expect(c.themeVariables?.primaryBorderColor).toBe("#2F6FB3");
  });

  it("mermaid の標準テーマはそのまま渡す", () => {
    expect(mermaidThemeConfig("forest")).toEqual({ theme: "forest" });
  });

  it("既定はモダン、全テーマに表示名がある", () => {
    expect(DEFAULT_THEME).toBe("modern");
    for (const t of THEMES) expect(THEME_LABELS[t]).toBeTruthy();
  });
});

describe("テンプレートの色", () => {
  it("アクセント1を枠、薄めた色を塗りに使い、円グラフはアクセント1〜6", () => {
    const v = mermaidThemeConfig("template", office2007).themeVariables!;
    expect(v.primaryBorderColor).toBe("#4F81BD");
    expect(luminance(v.primaryColor)).toBeGreaterThan(0.8);
    expect(v.textColor).toBe("#000000");
    expect([1, 2, 3, 4, 5, 6].map((i) => v[`pie${i}`])).toEqual(office2007.accents);
  });

  it("アクセントが明るい配色では円グラフの文字を濃くする", () => {
    const pastel = { ...office2007, accents: Array(6).fill("#FFF2CC") };
    expect(mermaidThemeConfig("template", pastel).themeVariables!.pieSectionTextColor).toBe("#000000");
    expect(mermaidThemeConfig("template", office2007).themeVariables!.pieSectionTextColor).toBe("#FFFFFF");
  });

  it("配色が読めなければモダンの色にする", () => {
    expect(mermaidThemeConfig("template", null)).toEqual({ theme: "base", themeVariables: CUSTOM_THEMES.modern });
  });

  it("themeKey は配色が変わると変わり、ほかのテーマはテーマ名のまま", () => {
    const other = { ...office2007, accents: ["#123456", ...office2007.accents.slice(1)] };
    expect(themeKey("template", office2007)).not.toBe(themeKey("template", other));
    expect(themeKey("template", office2007).length).toBeLessThan(100);
    expect(themeKey("modern", office2007)).toBe("modern");
  });
});

describe("カスタム", () => {
  it("メイン色を枠に、文字色を文字に使い、円グラフの6色は互いに違う", () => {
    const palette = customPalette({ main: "#E60012", sub: "#333333", text: "#222222" });
    const v = mermaidThemeConfig("custom", palette).themeVariables!;
    expect(v.primaryBorderColor).toBe("#E60012");
    expect(v.textColor).toBe("#222222");
    expect(new Set([1, 2, 3, 4, 5, 6].map((i) => v[`pie${i}`])).size).toBe(6);
  });

  it("paletteFor は template / custom だけ配色を返す", () => {
    expect(paletteFor("template", office2007, DEFAULT_CUSTOM_COLORS)).toBe(office2007);
    expect(paletteFor("custom", office2007, DEFAULT_CUSTOM_COLORS)?.accents[0]).toBe(DEFAULT_CUSTOM_COLORS.main);
    expect(paletteFor("modern", office2007, DEFAULT_CUSTOM_COLORS)).toBeNull();
  });

  it("色を変えると themeKey が変わる", () => {
    const a = themeKey("custom", customPalette(DEFAULT_CUSTOM_COLORS));
    const b = themeKey("custom", customPalette({ ...DEFAULT_CUSTOM_COLORS, main: "#000000" }));
    expect(a).not.toBe(b);
    expect(a.startsWith("custom:")).toBe(true);
  });

  it("テンプレートの配色から初期値を作る", () => {
    expect(customFromTemplate(office2007)).toEqual({ main: "#4F81BD", sub: "#C0504D", text: "#000000" });
  });

  it("保存値の検査", () => {
    expect(isCustomColors(DEFAULT_CUSTOM_COLORS)).toBe(true);
    expect(isCustomColors({ main: "red", sub: "#000000", text: "#000000" })).toBe(false);
    expect(isCustomColors(null)).toBe(false);
  });
});

describe("customFromThemeKey（図のタグからカスタムの初期値を作る）", () => {
  it("独自スタイルは枠の色・円グラフ2色目・文字色", () => {
    expect(customFromThemeKey("modern")).toEqual({
      main: CUSTOM_THEMES.modern!.primaryBorderColor,
      sub: CUSTOM_THEMES.modern!.pie2,
      text: CUSTOM_THEMES.modern!.textColor,
    });
  });

  it("テンプレートの色・カスタムで作った図は、記録した配色から戻す", () => {
    expect(customFromThemeKey(themeKey("template", office2007))).toEqual(customFromTemplate(office2007));
    const custom = { main: "#E60012", sub: "#333333", text: "#222222" };
    expect(customFromThemeKey(themeKey("custom", customPalette(custom)))).toEqual(custom);
  });

  it("mermaid 標準テーマとタグが空の図はおおよその色、読めない値は null", () => {
    expect(customFromThemeKey("forest")?.main).toBe("#13540C");
    expect(customFromThemeKey("")).toEqual(customFromThemeKey("default"));
    expect(customFromThemeKey("template")).toBeNull();
    expect(customFromThemeKey("unknown")).toBeNull();
  });
});
