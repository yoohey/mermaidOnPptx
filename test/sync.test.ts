import { describe, expect, it } from "vitest";
import { notesLineOfError, planSlide, type ExistingDiagram } from "../src/core/sync";
import { parseMermaidBlocks } from "../src/core/notesParser";

const box = { left: 0, top: 0, width: 100, height: 100 };
const diagram = (blockId: string, hash: string, theme = "default", shapeId = `s-${blockId}`): ExistingDiagram => ({
  ...box,
  shapeId,
  blockId,
  hash,
  theme,
});

const notes = "```mermaid id=a\ngraph TD\n```\n```mermaid id=b\ngraph LR\n```";
const hashes = new Map([
  ["a", "ha"],
  ["b", "hb"],
]);

function kinds(actions: ReturnType<typeof planSlide>) {
  return actions.map((a) => {
    if (a.kind === "orphan") return `orphan:${a.diagram.blockId}`;
    if (a.kind === "error") return `error:${a.line}`;
    return `${a.kind}:${a.block.id}`;
  });
}

describe("planSlide", () => {
  const base = { parsed: parseMermaidBlocks(notes), hashes, theme: "default" as const, force: false };

  it("ブロックあり・図なし → 新規挿入", () => {
    expect(kinds(planSlide({ ...base, diagrams: [] }))).toEqual(["insert:a", "insert:b"]);
  });

  it("ハッシュ一致 → 何もしない、不一致 → 置き換え", () => {
    const actions = planSlide({ ...base, diagrams: [diagram("a", "ha"), diagram("b", "old")] });
    expect(kinds(actions)).toEqual(["upToDate:a", "replace:b"]);
  });

  it("強制再生成ならハッシュ一致でも置き換え", () => {
    const actions = planSlide({ ...base, force: true, diagrams: [diagram("a", "ha")] });
    expect(kinds(actions)).toEqual(["replace:a", "insert:b"]);
  });

  it("テーマが変わったら置き換え。タグが無い（検証版の）図は default とみなす", () => {
    expect(kinds(planSlide({ ...base, diagrams: [diagram("a", "ha", "")] }))).toEqual(["upToDate:a", "insert:b"]);
    expect(kinds(planSlide({ ...base, theme: "forest", diagrams: [diagram("a", "ha")] }))).toEqual([
      "replace:a",
      "insert:b",
    ]);
  });

  it("ノートに無い図は削除せず orphan として知らせる", () => {
    const actions = planSlide({ ...base, diagrams: [diagram("a", "ha"), diagram("gone", "x")] });
    expect(kinds(actions)).toEqual(["upToDate:a", "insert:b", "orphan:gone"]);
  });

  it("同じ id の図が複数あれば（コピーした図）、すべて更新する", () => {
    const actions = planSlide({
      ...base,
      diagrams: [diagram("b", "old", "default", "s1"), diagram("b", "hb", "default", "s2"), diagram("b", "old", "default", "s3")],
    });
    expect(kinds(actions)).toEqual(["insert:a", "replace:b", "upToDate:b", "replace:b"]);
    expect(actions.filter((a) => a.kind === "replace").map((a) => a.kind === "replace" && a.diagram.shapeId)).toEqual([
      "s1",
      "s3",
    ]);
  });

  it("id 重複はスライドごと処理しない", () => {
    const parsed = parseMermaidBlocks("```mermaid id=a\ngraph TD\n```\n```mermaid id=a\ngraph LR\n```");
    expect(kinds(planSlide({ ...base, parsed, diagrams: [diagram("a", "old")] }))).toEqual(["error:4"]);
  });

  it("空ブロックはエラーにして図は触らない。ほかのブロックは処理する", () => {
    const parsed = parseMermaidBlocks("```mermaid id=a\n```\n```mermaid id=b\ngraph LR\n```");
    expect(kinds(planSlide({ ...base, parsed, diagrams: [diagram("a", "ha")] }))).toEqual(["error:1", "insert:b"]);
  });
});

describe("notesLineOfError", () => {
  it("ブロック内の行番号をノート全体の行番号に直す", () => {
    const block = { id: "a", code: "graph TD\n  A-->", line: 5 };
    expect(notesLineOfError(block, "Parse error on line 2:\n...")).toBe(7);
    expect(notesLineOfError(block, "No diagram type detected")).toBe(5);
  });
});
