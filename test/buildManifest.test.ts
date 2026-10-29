import { readFileSync } from "fs";
import { describe, expect, it } from "vitest";
// @ts-expect-error 型定義の無い CommonJS スクリプト
import { buildReleaseManifest, fillSiteTemplate, RELEASE_ID } from "../scripts/build-manifest.js";

const devXml = readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");

describe("buildReleaseManifest", () => {
  const xml: string = buildReleaseManifest(devXml);

  it("localhost を配信元の URL に置き換える", () => {
    expect(xml).not.toContain("localhost");
    expect(xml).toContain('<SourceLocation DefaultValue="https://yoohey.github.io/mermaidOnPptx/taskpane.html"/>');
    expect(xml).toContain("https://yoohey.github.io/mermaidOnPptx/assets/icon-32.png");
  });

  it("AppDomain はパスを含まないオリジンにする", () => {
    expect(xml).toContain("<AppDomain>https://yoohey.github.io</AppDomain>");
  });

  it("開発版と別の Id にし、表示名から (dev) を外す", () => {
    expect(xml).toContain(`<Id>${RELEASE_ID}</Id>`);
    expect(xml).not.toContain("0a3dd15a-988d-4e65-85ee-4ce438a84175");
    expect(xml).toContain('<DisplayName DefaultValue="Mermaid on PowerPoint"/>');
  });

  it("バージョンは開発用 manifest と同じ", () => {
    const version = (s: string) => /<Version>([^<]*)<\/Version>/.exec(s)?.[1];
    expect(version(xml)).toBe(version(devXml));
  });

  it("配信元を変えられる。https 以外は拒否する", () => {
    expect(buildReleaseManifest(devXml, "https://example.com/m4p/")).toContain(
      'DefaultValue="https://example.com/m4p/taskpane.html"',
    );
    expect(() => buildReleaseManifest(devXml, "http://example.com")).toThrow();
  });

  it("サポートと「詳細」は非公開リポジトリではなく配布サイトに向ける", () => {
    expect(xml).not.toContain("github.com/yoohey/mermaid4pptx");
    expect(xml).toContain('<SupportUrl DefaultValue="https://yoohey.github.io/mermaidOnPptx/support.html"/>');
    expect(xml).toContain('<bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://yoohey.github.io/mermaidOnPptx/"/>');
  });
});

describe("fillSiteTemplate", () => {
  it("配信元と公開用リポジトリの URL を埋める", () => {
    expect(fillSiteTemplate("__BASE_URL__/manifest.xml __REPO_URL__/issues", "https://yoohey.github.io/mermaidOnPptx/")).toBe(
      "https://yoohey.github.io/mermaidOnPptx/manifest.xml https://github.com/yoohey/mermaidOnPptx/issues",
    );
  });
});
