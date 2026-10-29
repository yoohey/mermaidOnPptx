import { readFileSync } from "fs";
import { describe, expect, it } from "vitest";
// @ts-expect-error 型定義の無い CommonJS スクリプト
import { publicReadme } from "../scripts/public-readme.js";

const readme = readFileSync(new URL("../README.md", import.meta.url), "utf8");

describe("publicReadme", () => {
  const text: string = publicReadme(readme, "https://yoohey.github.io/mermaidOnPptx");

  it("開発用リポジトリの話を除く", () => {
    for (const s of ["## Windows で試す", "## 開発", "## 配布版の公開", "> 設計:", "git clone", "mermaid4pptx"]) {
      expect(text).not.toContain(s);
    }
  });

  it("利用者向けの節は残し、導入はサイトへ案内する", () => {
    for (const s of ["# Mermaid on PowerPoint\n", "## 使い方", "## 制限事項", "## ライセンス"]) expect(text).toContain(s);
    expect(text).toContain("## 導入\n\n導入の手順は <https://yoohey.github.io/mermaidOnPptx/> をご覧ください。\n\n## ライセンス");
  });
});
