import { describe, expect, it } from "vitest";
import { boxToPx, fitScale, isResized, parseBox, formatBox, placeNew, viewBoxForBox } from "../src/core/layout";

const slide = { width: 960, height: 540 };

describe("placeNew", () => {
  it("小さい図は拡大せず、等倍で中央に置く", () => {
    // 100x50px → 75x38pt（文字を指定の大きさで見せるため）
    expect(placeNew({ width: 100, height: 50 }, slide)).toEqual({ left: 443, top: 251, width: 75, height: 38 });
  });

  it("縦長で高すぎる図は高さ（スライドの70%）に合わせて縮める", () => {
    const r = placeNew({ width: 200, height: 1000 }, slide);
    expect(r.height).toBe(378);
    expect(r.width).toBe(76);
  });

  it("大きい図は縦横比を保って幅80%に縮める", () => {
    const r = placeNew({ width: 4000, height: 1000 }, slide);
    expect(r.width).toBe(768);
    expect(r.height).toBe(192);
  });

  it("複数の新しい図は横に並べる", () => {
    const a = placeNew({ width: 4000, height: 1000 }, slide, 0, 2);
    const b = placeNew({ width: 4000, height: 1000 }, slide, 1, 2);
    expect(a.width).toBe(384);
    expect(a.left).toBe(96);
    expect(b.left).toBe(480);
  });
});

describe("fitScale / boxToPx", () => {
  it("縦横比を保って枠に収める倍率は、幅と高さの小さい方で決まる", () => {
    // 200x100px = 150x75pt を 300x300pt の枠へ → 幅で 2 倍
    expect(fitScale({ width: 200, height: 100 }, { width: 300, height: 300 })).toBe(2);
  });

  it("枠（pt）を等倍の SVG の大きさ（px）に直す", () => {
    expect(boxToPx({ width: 75, height: 150 })).toEqual({ width: 100, height: 200 });
  });
});

describe("viewBoxForBox", () => {
  it("横長の枠では左右に余白を足して中央に置く", () => {
    expect(viewBoxForBox([0, 0, 100, 100], { width: 200, height: 100 })).toEqual([-50, 0, 200, 100]);
  });

  it("縦長の枠では上下に余白を足す", () => {
    expect(viewBoxForBox([10, 10, 100, 50], { width: 100, height: 100 })).toEqual([10, -15, 100, 100]);
  });
});

describe("M4P_BOX と isResized", () => {
  it("幅と高さを文字列にして読み戻せる", () => {
    expect(parseBox(formatBox({ width: 300, height: 150.25 }))).toEqual({ width: 300, height: 150.3 });
    expect(parseBox("")).toBeNull();
  });

  it("1pt 以上変わっていれば大きさを変えたとみなす", () => {
    const inserted = { width: 300, height: 150 };
    expect(isResized(inserted, { width: 300.4, height: 150 })).toBe(false);
    expect(isResized(inserted, { width: 300, height: 200 })).toBe(true);
  });

  it("記録の無い図（1.0 で作った図）は変えていないとみなす", () => {
    expect(isResized(null, { width: 1, height: 1 })).toBe(false);
  });
});
