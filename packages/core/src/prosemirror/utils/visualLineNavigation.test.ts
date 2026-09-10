import { describe, expect, test } from "bun:test";

import { findPositionOnLineAtClientX } from "./visualLineNavigation";

type FakeElement = {
  closest: () => FakeElement | null;
  dataset: Record<string, string>;
  querySelector: () => FakeElement | null;
  querySelectorAll: () => FakeElement[];
};

const element = (dataset: Record<string, string> = {}): FakeElement => ({
  closest: () => null,
  dataset,
  querySelector: () => null,
  querySelectorAll: () => [],
});

describe("findPositionOnLineAtClientX", () => {
  test("uses an empty run's own position for a trailing blank line", () => {
    const paragraph = element({ pmStart: "10" });
    const emptyRun = element({ pmStart: "24", pmEnd: "25" });
    emptyRun.closest = () => paragraph;
    const line = element();
    line.querySelector = () => emptyRun;
    line.querySelectorAll = () => [emptyRun];

    expect(findPositionOnLineAtClientX(line as unknown as HTMLElement, 800)).toBe(24);
  });

  test("falls back to the paragraph content start for a structural empty line", () => {
    const paragraph = element({ pmStart: "30" });
    const emptyRun = element();
    emptyRun.closest = () => paragraph;
    const line = element();
    line.querySelector = () => emptyRun;

    expect(findPositionOnLineAtClientX(line as unknown as HTMLElement, 200)).toBe(31);
  });
});
