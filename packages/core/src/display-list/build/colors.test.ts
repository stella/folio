/**
 * Colour parsing takes an authored string, so every lookup it does must be a
 * lookup in data rather than in a prototype chain.
 */

import { describe, expect, test } from "bun:test";

import { parseDisplayColor } from "./colors";

describe("parseDisplayColor: inherited keys are not colours", () => {
  // `Object.prototype` members: an object-keyed palette answers all three with
  // a function or an object, which the hex parser then calls `.replace` on.
  const INHERITED_KEYS = ["constructor", "toString", "__proto__", "hasOwnProperty"] as const;

  for (const key of INHERITED_KEYS) {
    test(`\`${key}\` is an unknown colour, not an inherited value`, () => {
      expect(parseDisplayColor(key)).toBeUndefined();
      expect(parseDisplayColor(key.toUpperCase())).toBeUndefined();
    });
  }

  test("a custom property named after a prototype key falls through to its fallback", () => {
    // The `var()` pattern only matches a `--`-prefixed name, so no prototype
    // key can reach the custom-property table; the fallback is what answers.
    expect(parseDisplayColor("var(--constructor)")).toBeUndefined();
    expect(parseDisplayColor("var(--constructor, #ff0000)")).toEqual({
      r: 255,
      g: 0,
      b: 0,
      a: 1,
    });
  });

  test("a real named colour still resolves", () => {
    expect(parseDisplayColor("red")).toEqual({ r: 255, g: 0, b: 0, a: 1 });
    expect(parseDisplayColor("DarkYellow")).toEqual({ r: 128, g: 128, b: 0, a: 1 });
  });
});
