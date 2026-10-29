import { describe, expect, it } from "vitest";
import { parseMermaidBlocks } from "../src/core/notesParser";
import { SAMPLES, fencedSample } from "../src/core/samples";

describe("SAMPLES", () => {
  it("id が重複しない", () => {
    expect(new Set(SAMPLES.map((s) => s.id)).size).toBe(SAMPLES.length);
  });

  it("ノートに貼ると、そのまま1つのブロックとして読める", () => {
    for (const sample of SAMPLES) {
      const r = parseMermaidBlocks(`メモ\n${fencedSample(sample)}\n`);
      expect(r.errors).toEqual([]);
      expect(r.blocks.map((b) => b.code)).toEqual([sample.code]);
    }
  });
});
