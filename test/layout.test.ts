import { describe, expect, it } from "vitest";
import { placeNew, placeReplacing } from "../src/core/layout";

const slide = { width: 960, height: 540 };

describe("placeNew", () => {
  it("小さい図は2倍まで拡大して中央に置く", () => {
    // 100x50px → 75x37.5pt → 2倍で 150x75pt
    expect(placeNew({ width: 100, height: 50 }, slide)).toEqual({ left: 405, top: 233, width: 150, height: 75 });
  });

  it("縦長の図は高さ（スライドの70%）に合わせる", () => {
    // 状態遷移図の実例 98x304pt 相当 → 高さ 378pt まで
    const r = placeNew({ width: 98 / 0.75, height: 304 / 0.75 }, slide);
    expect(r.height).toBe(378);
    expect(r.width).toBe(122);
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

describe("placeReplacing", () => {
  it("左上と横幅を引き継ぎ、高さを新しい縦横比で決める", () => {
    const old = { left: 10, top: 20, width: 300, height: 100 };
    expect(placeReplacing(old, { width: 200, height: 100 })).toEqual({ left: 10, top: 20, width: 300, height: 150 });
  });
});
