import { describe, expect, test } from "bun:test";

import { CharacterSpacingExtension } from "./CharacterSpacingExtension";

const parseScale = (scale: string) => {
  const getAttrs = CharacterSpacingExtension().config.markSpec.parseDOM?.at(0)?.getAttrs;
  if (!getAttrs) {
    throw new Error("CharacterSpacingExtension must define parseDOM[0].getAttrs");
  }

  // SAFETY: getAttrs reads only HTMLElement.dataset in this test fixture.
  const attrs = getAttrs({ dataset: { scale } } as unknown as HTMLElement);
  if (!attrs) {
    throw new Error("CharacterSpacingExtension must return mark attributes");
  }
  return attrs["scale"];
};

const serializeScale = (scale: number) => {
  const toDOM = CharacterSpacingExtension().config.markSpec.toDOM;
  if (!toDOM) {
    throw new Error("CharacterSpacingExtension must define toDOM");
  }
  // SAFETY: toDOM reads only the character-spacing mark's attrs in this fixture.
  const mark = {
    type: { name: "characterSpacing" },
    attrs: { scale },
  } as Parameters<typeof toDOM>[0];
  const output = toDOM(mark, true);
  if (!Array.isArray(output)) {
    throw new TypeError("Expected character-spacing DOM output to be an array");
  }
  // SAFETY: CharacterSpacingExtension always emits an attribute record at index 1.
  return output[1] as Record<string, string>;
};

describe("CharacterSpacingExtension horizontal scale parsing", () => {
  test.each([
    [" 0% ", 0],
    ["100", 100],
    ["600%", 600],
  ])("preserves strict data-scale %p", (scale, expected) => {
    expect(parseScale(scale)).toBe(expected);
  });

  test.each(["", "   ", "12.5", "0garbage", "1e2", "600oops", "0x10", "601%"])(
    "drops malformed data-scale %p",
    (scale) => {
      expect(parseScale(scale)).toBeUndefined();
    },
  );

  test("preserves an explicit 100% override through DOM serialization", () => {
    expect(serializeScale(100)["data-scale"]).toBe("100");
  });

  test("does not serialize a fractional scale outside the OOXML model domain", () => {
    expect(serializeScale(12.5)["data-scale"]).toBeUndefined();
  });
});
