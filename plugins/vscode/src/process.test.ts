import { describe, expect, test } from "bun:test";

import { stderrTail } from "./process";

describe("stderrTail", () => {
  test("keeps the last lines", () => {
    expect(stderrTail("a\nb\nc\nd\n", 2)).toBe("c\nd");
    expect(stderrTail("  only  \n")).toBe("only");
    expect(stderrTail("")).toBe("");
  });
});
