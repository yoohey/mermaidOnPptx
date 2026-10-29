import { describe, expect, it } from "vitest";
import { normalizeNotesText, parseMermaidBlocks } from "../src/core/notesParser";

describe("parseMermaidBlocks", () => {
  it("フェンス外のメモを無視してブロックを取り出す", () => {
    const r = parseMermaidBlocks("メモ\n```mermaid\nsequenceDiagram\n  A->>B: hi\n```\n後ろのメモ");
    expect(r.errors).toEqual([]);
    expect(r.blocks).toEqual([{ id: "1", code: "sequenceDiagram\n  A->>B: hi", line: 2 }]);
  });

  it("id 指定と自動採番を併用できる", () => {
    const r = parseMermaidBlocks("```mermaid id=login\ngraph TD\n```\n```mermaid\ngraph LR\n```");
    expect(r.blocks.map((b) => b.id)).toEqual(["login", "2"]);
  });

  it("引用符付きの id も読む", () => {
    const r = parseMermaidBlocks('```mermaid id="a-1"\ngraph TD\n```');
    expect(r.blocks[0].id).toBe("a-1");
  });

  it("全角バッククォートと段落内改行（\\v）を受け付ける", () => {
    const r = parseMermaidBlocks("｀｀｀mermaid\vgraph TD\v｀｀｀");
    expect(r.blocks).toEqual([{ id: "1", code: "graph TD", line: 1 }]);
  });

  it("~~~ フェンスと4つ以上のバッククォートを受け付ける", () => {
    expect(parseMermaidBlocks("~~~mermaid\ngraph TD\n~~~").blocks).toHaveLength(1);
    expect(parseMermaidBlocks("````mermaid\ngraph TD\n```\n````").blocks[0].code).toBe("graph TD\n```");
  });

  it("スマートクォートを半角に戻す", () => {
    const r = parseMermaidBlocks("```mermaid\nA[“開始”] --> B[‘x’]\n```");
    expect(r.blocks[0].code).toBe("A[\"開始\"] --> B['x']");
  });

  it("終了行が無ければエラー", () => {
    const r = parseMermaidBlocks("```mermaid\ngraph TD");
    expect(r.blocks).toEqual([]);
    expect(r.errors).toEqual([{ line: 1, message: expect.stringContaining("終了行") }]);
  });

  it("id の重複と空ブロックはエラー", () => {
    const r = parseMermaidBlocks("```mermaid id=x\ngraph TD\n```\n```mermaid id=x\n\n```");
    expect(r.errors.map((e) => e.line)).toEqual([4, 4]);
  });

  it("mermaid 以外のフェンスは無視する", () => {
    expect(parseMermaidBlocks("```js\nconst a = 1\n```").blocks).toEqual([]);
  });

  it("mermaidjs のような別名は mermaid と見なさない", () => {
    expect(parseMermaidBlocks("```mermaidjs\ngraph TD\n```").blocks).toEqual([]);
  });
});

describe("normalizeNotesText", () => {
  it("CRLF と NBSP を正規化する", () => {
    expect(normalizeNotesText("a\r\nb c")).toBe("a\nb c");
  });
});
