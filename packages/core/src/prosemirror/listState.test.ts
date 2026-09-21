/**
 * Which kind of list a paragraph is in is a property of the numbering
 * definitions, not of the id. The adapters used to read `numId === 1` as
 * bullets, which held only for lists Folio had created itself: an imported
 * document numbers its lists however its author did, and a bulleted import
 * whose instance happened to be 4 read as numbered in every toolbar.
 */

import { describe, expect, test } from "bun:test";

import { getCachedNumberingMap } from "../docx/numberingParser";
import type { NumberingDefinitions } from "../types/document";
import { NO_LIST_STATE, isInListState, resolveListState, sameListState } from "./listState";

const definitions = {
  abstractNums: [
    { abstractNumId: 1, levels: [{ ilvl: 0, numFmt: "bullet", lvlText: "•" }] },
    {
      abstractNumId: 2,
      levels: [
        { ilvl: 0, numFmt: "decimal", lvlText: "%1." },
        { ilvl: 1, numFmt: "bullet", lvlText: "•" },
      ],
    },
  ],
  nums: [
    { numId: 4, abstractNumId: 1 },
    { numId: 9, abstractNumId: 2 },
  ],
} as const satisfies NumberingDefinitions;

const numbering = getCachedNumberingMap(definitions);

describe("resolveListState", () => {
  test("reads a bullet list off its level's w:numFmt, whatever its id is", () => {
    expect(resolveListState(numbering, { kind: "reference", numId: 4, ilvl: 0 })).toEqual({
      type: "bullet",
      level: 0,
      numId: 4,
    });
  });

  test("reads a numbered list the same way", () => {
    expect(resolveListState(numbering, { kind: "reference", numId: 9, ilvl: 0 })).toEqual({
      type: "numbered",
      level: 0,
      numId: 9,
    });
  });

  test("resolves per level, not per instance", () => {
    expect(resolveListState(numbering, { kind: "reference", numId: 9, ilvl: 1 })).toEqual({
      type: "bullet",
      level: 1,
      numId: 9,
    });
  });

  test("an absent w:ilvl renders at level zero", () => {
    expect(resolveListState(numbering, { kind: "reference", numId: 4 })).toEqual({
      type: "bullet",
      level: 0,
      numId: 4,
    });
  });

  test("a cancellation is not a list", () => {
    expect(resolveListState(numbering, { kind: "none" })).toEqual(NO_LIST_STATE);
    expect(resolveListState(numbering, undefined)).toEqual(NO_LIST_STATE);
  });

  test("a level stated without a resolved id is not a list", () => {
    expect(resolveListState(numbering, { kind: "levelOnly", ilvl: 2 })).toEqual(NO_LIST_STATE);
  });

  test("an id no definition covers falls back to numbered", () => {
    expect(resolveListState(numbering, { kind: "reference", numId: 77, ilvl: 0 })).toEqual({
      type: "numbered",
      level: 0,
      numId: 77,
    });
    expect(resolveListState(null, { kind: "reference", numId: 4, ilvl: 0 })).toEqual({
      type: "numbered",
      level: 0,
      numId: 4,
    });
  });
});

describe("the union replaces the flag set", () => {
  test("being in a list is the discriminator, not a separate field", () => {
    expect(isInListState(NO_LIST_STATE)).toBe(false);
    expect(isInListState({ type: "bullet", level: 0, numId: 4 })).toBe(true);
    expect(isInListState(undefined)).toBe(false);
  });

  test("equality covers every arm", () => {
    expect(sameListState(NO_LIST_STATE, { type: "none" })).toBe(true);
    expect(sameListState(NO_LIST_STATE, { type: "numbered", level: 0 })).toBe(false);
    expect(
      sameListState({ type: "bullet", level: 1, numId: 4 }, { type: "bullet", level: 1, numId: 4 }),
    ).toBe(true);
    expect(
      sameListState({ type: "bullet", level: 1, numId: 4 }, { type: "bullet", level: 1, numId: 9 }),
    ).toBe(false);
    expect(sameListState(undefined, NO_LIST_STATE)).toBe(false);
  });
});
