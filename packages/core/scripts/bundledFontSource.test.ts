/**
 * The harness emits `@font-face` rules into a `<style>` element, and a family
 * name reaching them comes from an authored document.
 */

import { describe, expect, test } from "bun:test";

import { cssString } from "./bundledFontSource";

describe("cssString", () => {
  test("a backslash is escaped before the quote it precedes", () => {
    // The round trip that breaks when the quote is escaped first: `\"` would
    // become `\\"`, which CSS reads as a backslash and then the closing quote.
    expect(cssString('a\\"; } body { display: none } .x {')).toBe(
      '"a\\\\\\"; } body { display: none } .x {"',
    );
  });

  test("a quote alone is escaped", () => {
    expect(cssString('Fancy" Font')).toBe('"Fancy\\" Font"');
  });

  test("a line terminator becomes a hex escape, which CSS strings require", () => {
    expect(cssString("Two\nLines")).toBe('"Two\\A Lines"');
    expect(cssString("Two\r\nLines")).toBe('"Two\\D \\A Lines"');
    expect(cssString("Form\ffeed")).toBe('"Form\\C feed"');
  });

  test("an ordinary name is quoted and otherwise untouched", () => {
    expect(cssString("Times New Roman")).toBe('"Times New Roman"');
  });

  test("the escaped string closes exactly one string token", () => {
    // Everything after the opening quote up to the final one is escaped, so a
    // hostile name cannot end the token early.
    const emitted = cssString('x"; } @import url(evil.css); .y { content: "');
    expect(emitted.startsWith('"')).toBe(true);
    expect(emitted.endsWith('"')).toBe(true);
    expect(emitted.slice(1, -1)).not.toMatch(/(?<!\\)(?:\\\\)*"/u);
  });
});
