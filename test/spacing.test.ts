import { describe, expect, it } from "vitest";
import { baseSpacing, diagramKind, renderToFill, shrinkToFit, spacingConfig, type Spacing } from "../src/core/spacing";

describe("diagramKind", () => {
  it("フローチャートの向きを読む。%% の行は飛ばす", () => {
    expect(diagramKind('%%{init: {"theme": "forest"}}%%\nflowchart LR\n  A --> B')).toEqual({ kind: "flowchart", horizontal: true });
    expect(diagramKind("graph TD\n  A --> B")).toEqual({ kind: "flowchart", horizontal: false });
  });

  it("状態遷移図などは本文の direction で向きを読む", () => {
    expect(diagramKind("stateDiagram-v2\n  direction LR\n  A --> B")).toEqual({ kind: "state", horizontal: true });
    expect(diagramKind("classDiagram\n  A <|-- B").kind).toBe("class");
    expect(diagramKind("sequenceDiagram\n  A->>B: hi").kind).toBe("sequence");
    expect(diagramKind("pie\n  \"a\": 1").kind).toBe("other");
  });
});

describe("spacingConfig", () => {
  it("横向きのフローチャートは段の間で幅、ノードの間で高さを変える", () => {
    expect(spacingConfig("flowchart LR\n A-->B", { w: 120, h: 80 })).toEqual({
      flowchart: { rankSpacing: 120, nodeSpacing: 80 },
    });
  });

  it("シーケンス図は登場人物の間で幅、メッセージの間で高さを変える", () => {
    expect(spacingConfig("sequenceDiagram\n A->>B: x", { w: 100, h: 20 })).toEqual({
      sequence: { actorMargin: 100, boxMargin: 20 },
    });
  });

  it("円グラフは広げる設定が無い", () => {
    expect(baseSpacing("pie\n \"a\": 1")).toEqual({ w: undefined, h: undefined });
  });
});

describe("shrinkToFit", () => {
  it("収まるなら 1、はみ出すなら収まる倍率", () => {
    expect(shrinkToFit({ width: 100, height: 100 }, { width: 200, height: 200 })).toBe(1);
    expect(shrinkToFit({ width: 400, height: 100 }, { width: 200, height: 200 })).toBe(0.5);
  });
});

/** 大きさが間隔に比例する図。fixed は間隔に依らない部分、gaps は間隔の数（0 ならその向きは広がらない） */
function fakeRender(fixed: { w: number; h: number }, gaps: { w: number; h: number }) {
  const calls: Spacing[] = [];
  const render = async (s: Spacing) => {
    calls.push(s);
    return { width: fixed.w + gaps.w * (s.w ?? 0), height: fixed.h + gaps.h * (s.h ?? 0) };
  };
  return { render, calls };
}

const LR = "flowchart LR\n A-->B-->C";

describe("renderToFill", () => {
  it("横長の枠に合わせて、幅も高さもほぼいっぱいに広げる", async () => {
    const { render } = fakeRender({ w: 300, h: 100 }, { w: 2, h: 1 });
    const target = { width: 1200, height: 400 };
    const { rendered } = await renderToFill(LR, render, target);
    expect(rendered.width).toBeGreaterThan(1150);
    expect(rendered.width).toBeLessThanOrEqual(1200);
    expect(rendered.height).toBeGreaterThan(380);
  });

  it("幅を広げる設定が効かない図（縦一列）は、高さだけ広げる", async () => {
    const { render } = fakeRender({ w: 100, h: 300 }, { w: 0, h: 2 });
    const { rendered, spacing } = await renderToFill(LR, render, { width: 800, height: 600 });
    expect(rendered.width).toBe(100);
    expect(rendered.height).toBeGreaterThan(580);
    expect(spacing.w).toBe(50);
  });

  it("枠より大きい図は間隔を詰める（最小は既定の 0.6 倍）", async () => {
    const { render } = fakeRender({ w: 300, h: 100 }, { w: 4, h: 1 });
    const { spacing } = await renderToFill(LR, render, { width: 350, height: 150 });
    expect(spacing.w).toBe(30);
  });

  it("間隔を変えた描画が失敗したら、それまでの最良を使う", async () => {
    const { render } = fakeRender({ w: 300, h: 100 }, { w: 2, h: 1 });
    let n = 0;
    const failing = async (s: Spacing) => {
      if (n++ > 0) throw new Error("Could not find a suitable point");
      return render(s);
    };
    const { spacing } = await renderToFill(LR, failing, { width: 1200, height: 400 });
    expect(spacing).toEqual({ w: 50, h: 50 });
  });

  it("広げる設定の無い図は 1 回だけ描く", async () => {
    const { render, calls } = fakeRender({ w: 300, h: 300 }, { w: 0, h: 0 });
    await renderToFill('pie\n "a": 1', render, { width: 800, height: 600 });
    expect(calls).toHaveLength(1);
  });
});

describe("renderToFill（縮めるとき）", () => {
  it("高さが収まらず縮めるときは、縮めたあとに幅いっぱいになるよう横に広げる", async () => {
    // 高さは間隔では変わらず 800px（枠 400px の 2 倍）→ 0.5 倍に縮める。幅は 1200 / 0.5 = 2400px を狙う
    const { render } = fakeRender({ w: 300, h: 800 }, { w: 2, h: 0 });
    const target = { width: 1200, height: 400 };
    const { rendered } = await renderToFill(LR, render, target);
    const s = shrinkToFit(rendered, target);
    expect(s).toBeCloseTo(0.5, 2);
    expect((rendered.width * s) / target.width).toBeGreaterThan(0.95);
  });
});
