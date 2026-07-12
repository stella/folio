import { describe, expect, test } from "bun:test";

import { createLatestRequestGate } from "./latestRequestGate";

describe("createLatestRequestGate", () => {
  test("keeps the newest request current", () => {
    const gate = createLatestRequestGate();
    const first = gate.begin();

    expect(first()).toBe(true);

    const second = gate.begin();

    expect(first()).toBe(false);
    expect(second()).toBe(true);
  });

  test("invalidates a request without starting another one", () => {
    const gate = createLatestRequestGate();
    const request = gate.begin();

    gate.invalidate();

    expect(request()).toBe(false);
  });

  test("an old guard never becomes current again", () => {
    const gate = createLatestRequestGate();
    const first = gate.begin();

    gate.invalidate();
    const second = gate.begin();

    expect(first()).toBe(false);
    expect(second()).toBe(true);
  });
});
