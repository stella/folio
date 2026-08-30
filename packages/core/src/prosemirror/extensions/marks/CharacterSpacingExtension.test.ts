import { describe, expect, test } from "bun:test";

import { CharacterSpacingExtension } from "./CharacterSpacingExtension";

const parseScale = (scale: string) => {
  const getAttrs = CharacterSpacingExtension().config.markSpec.parseDOM?.at(0)?.getAttrs;
  if (!getAttrs) {
    throw new Error("CharacterSpacingExtension must define parseDOM[0].getAttrs");
  }

  const attrs = getAttrs({ dataset: { scale } } as unknown as HTMLElement);
  if (!attrs) {
    throw new Error("CharacterSpacingExtension must return mark attributes");
  }
  return attrs["scale"];
};

describe("CharacterSpacingExtension horizontal scale parsing", () => {
  test.each([
    [" 0% ", 0],
    ["600%", 600],
  ])("preserves strict data-scale %p", (scale, expected) => {
    expect(parseScale(scale)).toBe(expected);
  });

  test.each(["", "   ", "0garbage", "1e2", "600oops", "0x10", "601%"])(
    "drops malformed data-scale %p",
    (scale) => {
      expect(parseScale(scale)).toBeUndefined();
    },
  );
});
