import { describe, expect, it } from "vitest";
import { sourceHash } from "../src/core/hash";

describe("sourceHash", () => {
  it("SHA-256 の先頭16文字を返す", async () => {
    // echo -n "graph TD" | sha256sum
    expect(await sourceHash("graph TD")).toBe("ad6586f310d87461");
    expect(await sourceHash("graph TD")).toBe(await sourceHash("graph TD"));
    expect(await sourceHash("graph TD")).not.toBe(await sourceHash("graph LR"));
  });
});
