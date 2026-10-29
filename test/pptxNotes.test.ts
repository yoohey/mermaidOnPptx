import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  parseThemeColors,
  readPresentationInfo,
  readSlideNotes,
  resolvePartPath,
  sldIdFromOfficeSlideId,
} from "../src/core/pptxNotes";
import { parseMermaidBlocks } from "../src/core/notesParser";

// test/fixtures/make_fixture.py で生成。表示順は 257, 258, 256（ファイル名順とずらしてある）
const fixture = readFileSync(join(__dirname, "fixtures/notes.pptx"));

describe("readPresentationInfo", () => {
  it("スライドの大きさを pt で返す（python-pptx の既定は 4:3 の 720x540）", async () => {
    const info = await readPresentationInfo(fixture);
    expect(info.slideSize).toEqual({ width: 720, height: 540 });
    expect(info.slides).toHaveLength(3);
  });

  it("マスターのテーマの配色を読む（python-pptx の既定テーマは Office 2007 の配色）", async () => {
    const info = await readPresentationInfo(fixture);
    expect(info.themeColors).toEqual({
      dk1: "#000000",
      lt1: "#FFFFFF",
      dk2: "#1F497D",
      lt2: "#EEECE1",
      accents: ["#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646"],
    });
  });
});

describe("parseThemeColors", () => {
  const scheme = (inner: string) =>
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x">${inner}</a:clrScheme></a:themeElements></a:theme>`;
  const slot = (name: string, color: string) => `<a:${name}><a:srgbClr val="${color}"/></a:${name}>`;
  const full = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

  it("srgbClr は小文字でも大文字にそろえる", () => {
    const c = parseThemeColors(scheme(full.map((n) => slot(n, "aabbcc")).join("")));
    expect(c?.accents[0]).toBe("#AABBCC");
  });

  it("色が欠けていれば null", () => {
    expect(parseThemeColors(scheme(full.slice(0, 9).map((n) => slot(n, "112233")).join("")))).toBeNull();
    expect(parseThemeColors("<a:theme xmlns:a='x'/>")).toBeNull();
  });
});

describe("readSlideNotes", () => {
  it("表示順にスライドとノートを返す", async () => {
    const slides = await readSlideNotes(fixture);
    expect(slides.map((s) => [s.sldId, s.index])).toEqual([
      [257, 0],
      [258, 1],
      [256, 2],
    ]);
  });

  it("ノートの無いスライドは null", async () => {
    const slides = await readSlideNotes(fixture);
    expect(slides[0].notesText).toBeNull();
  });

  it("複数段落を改行で連結する", async () => {
    const slides = await readSlideNotes(fixture);
    const s = slides.find((x) => x.sldId === 256)!;
    expect(s.notesText).toBe("冒頭のメモ\n```mermaid\nsequenceDiagram\n  User->>App: ログイン\n  App-->>User: OK\n```");
    expect(parseMermaidBlocks(s.notesText!).blocks[0].code).toContain("User->>App");
  });

  it("段落内改行（a:br）を改行として扱い、記法として解析できる", async () => {
    const slides = await readSlideNotes(fixture);
    const s = slides.find((x) => x.sldId === 258)!;
    const parsed = parseMermaidBlocks(s.notesText!);
    expect(parsed.errors).toEqual([]);
    expect(parsed.blocks[0]).toMatchObject({ id: "flow", code: 'flowchart LR\n  A["開始"] --> B' });
  });
});

describe("resolvePartPath", () => {
  it("相対・絶対の Target を解決する", () => {
    expect(resolvePartPath("ppt/presentation.xml", "slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
    expect(resolvePartPath("ppt/slides/slide1.xml", "../notesSlides/notesSlide1.xml")).toBe(
      "ppt/notesSlides/notesSlide1.xml",
    );
    expect(resolvePartPath("ppt/slides/slide1.xml", "/ppt/notesSlides/n.xml")).toBe("ppt/notesSlides/n.xml");
  });
});

describe("sldIdFromOfficeSlideId", () => {
  it("先頭の数値を取り出す", () => {
    expect(sldIdFromOfficeSlideId("267#")).toBe(267);
    expect(sldIdFromOfficeSlideId("267#1234567")).toBe(267);
    expect(sldIdFromOfficeSlideId("abc")).toBeNull();
  });
});
