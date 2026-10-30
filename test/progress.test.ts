import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { delayedIndicator, progressView } from "../src/core/progress";

describe("progressView", () => {
  it("読み込みは受け取った分割数で進む", () => {
    expect(progressView({ phase: "read", done: 3, total: 12 })).toEqual({
      label: "ファイルを読み込み中（3/12）",
      value: 0.25,
    });
  });

  it("ノートの読み取りは進み具合が分からない表示にする", () => {
    expect(progressView({ phase: "parse" }).value).toBeNull();
  });

  it("全スライドの更新ではスライドの枚数を出す", () => {
    expect(progressView({ phase: "slides", done: 5, total: 40 })).toEqual({
      label: "図を生成中（スライド 5/40）",
      value: 0.125,
    });
  });

  it("1 枚だけのときは枚数を出さない", () => {
    expect(progressView({ phase: "slides", done: 0, total: 1 }).label).toBe("図を生成中");
  });
});

describe("delayedIndicator", () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it("すぐ終われば表示しない", () => {
    const show = vi.fn();
    const hide = vi.fn();
    const ind = delayedIndicator(300, show, hide);
    ind.start();
    vi.advanceTimersByTime(299);
    ind.stop();
    vi.advanceTimersByTime(1000);
    expect(show).not.toHaveBeenCalled();
    expect(hide).not.toHaveBeenCalled();
  });

  it("時間がかかれば表示し、終わったら消す", () => {
    const show = vi.fn();
    const hide = vi.fn();
    const ind = delayedIndicator(300, show, hide);
    ind.start();
    vi.advanceTimersByTime(300);
    expect(show).toHaveBeenCalledTimes(1);
    ind.stop();
    expect(hide).toHaveBeenCalledTimes(1);
  });

  it("続けて実行しても使い回せる", () => {
    const show = vi.fn();
    const hide = vi.fn();
    const ind = delayedIndicator(300, show, hide);
    for (let i = 0; i < 2; i++) {
      ind.start();
      vi.advanceTimersByTime(500);
      ind.stop();
    }
    expect(show).toHaveBeenCalledTimes(2);
    expect(hide).toHaveBeenCalledTimes(2);
  });
});
