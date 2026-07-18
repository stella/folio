import { describe, expect, test } from "bun:test";

import { stripXmlDeclarations } from "./stripXmlDeclarations";

describe("stripXmlDeclarations", () => {
  test("removes complete xml declarations case-insensitively", () => {
    expect(stripXmlDeclarations('a<?xml version="1.0"?>b')).toBe("ab");
    expect(stripXmlDeclarations("a<?XML encoding='utf-8'?>b")).toBe("ab");
  });

  test("preserves unterminated openers", () => {
    expect(stripXmlDeclarations("keep<?xml dangling")).toBe("keep<?xml dangling");
  });

  test("strips many unterminated openers in linear time", () => {
    const evil = "<?xml".repeat(50_000);
    const start = performance.now();
    const out = stripXmlDeclarations(evil);
    expect(performance.now() - start).toBeLessThan(2000);
    expect(out).toBe(evil);
  });
});
