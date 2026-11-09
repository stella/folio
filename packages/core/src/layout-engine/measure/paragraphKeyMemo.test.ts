import { describe, expect, test } from "bun:test";

import type { DocumentFontFacts } from "../../fonts/documentFontFacts";
import type { ParagraphBlock, Run } from "../types";
import { memoizedParagraphKey, noteParagraphOrigin } from "./paragraphKeyMemo";

const formatting = { bold: true };

const run = (text: string, pmStart: number): Run => ({
  kind: "text",
  text,
  fontFamily: "Body",
  fontSize: 11,
  pmStart,
  pmEnd: pmStart + text.length,
  ...formatting,
});

const paragraph = (pmStart: number): ParagraphBlock => ({
  kind: "paragraph",
  id: `p-${String(pmStart)}`,
  pmStart,
  pmEnd: pmStart + 5,
  runs: [run("Hello", pmStart)],
  attrs: {
    alignment: "left",
    listMarkerFormatting: { fontFamily: "Body", fontSize: 11 },
  },
});

/** Moves `block` the way the block cache does: own fields, shared values. */
const moved = (block: ParagraphBlock, delta: number): ParagraphBlock => ({
  ...block,
  runs: block.runs.map((each) => ({
    ...each,
    ...(each.pmStart === undefined ? {} : { pmStart: each.pmStart + delta }),
    ...(each.pmEnd === undefined ? {} : { pmEnd: each.pmEnd + delta }),
  })),
  attrs: { ...block.attrs },
});

const counting = () => {
  let calls = 0;
  return {
    calls: () => calls,
    compute: (block: ParagraphBlock) => {
      calls += 1;
      return JSON.stringify({
        attrs: block.attrs,
        runs: block.runs.map(({ pmStart: _s, pmEnd: _e, ...rest }) => rest),
      });
    },
  };
};

/** A kept block and one block moved from it, keyed once. */
const keyed = () => {
  const origin = paragraph(10);
  const first = moved(origin, 0);
  noteParagraphOrigin(first, origin);
  const counter = counting();
  const key = memoizedParagraphKey(first, 1, counter.compute);
  return { origin, first, counter, key };
};

describe("memoizedParagraphKey", () => {
  test("a block moved again takes the key without serializing", () => {
    const { origin, counter, key } = keyed();
    const next = moved(origin, 7);
    noteParagraphOrigin(next, origin);
    expect(memoizedParagraphKey(next, 1, counter.compute)).toBe(key);
    expect(counter.calls()).toBe(1);
  });

  test("a nested value copied with the same fields keeps the key", () => {
    const { origin, counter, key } = keyed();
    const next = moved(origin, 3);
    next.attrs = {
      ...next.attrs,
      listMarkerFormatting: { fontFamily: "Body", fontSize: 11 },
    };
    noteParagraphOrigin(next, origin);
    expect(memoizedParagraphKey(next, 1, counter.compute)).toBe(key);
    expect(counter.calls()).toBe(1);
  });

  test.each([
    [
      "a run's text",
      (block: ParagraphBlock) => {
        block.runs = [run("Hellx", 13)];
      },
    ],
    [
      "a paragraph attribute",
      (block: ParagraphBlock) => {
        block.attrs = { ...block.attrs, alignment: "center" };
      },
    ],
    [
      "a nested field",
      (block: ParagraphBlock) => {
        block.attrs = {
          ...block.attrs,
          listMarkerFormatting: { fontFamily: "Body", fontSize: 12 },
        };
      },
    ],
    [
      "an added field",
      (block: ParagraphBlock) => {
        const [first] = block.runs;
        if (first?.kind === "text") block.runs = [{ ...first, italic: true }];
      },
    ],
  ])("a change to %s serializes again", (_name, change) => {
    const { origin, counter, key } = keyed();
    const next = moved(origin, 3);
    change(next);
    noteParagraphOrigin(next, origin);
    const nextKey = memoizedParagraphKey(next, 1, counter.compute);
    expect(counter.calls()).toBe(2);
    expect(nextKey).toBe(counter.compute(next));
    expect(nextKey).not.toBe(key);
  });

  test("a new line break generation serializes again", () => {
    const { origin, counter } = keyed();
    const next = moved(origin, 3);
    noteParagraphOrigin(next, origin);
    memoizedParagraphKey(next, 2, counter.compute);
    expect(counter.calls()).toBe(2);
  });

  test("font facts whose identity changes on the same object serialize again", () => {
    // Not a plain record: font facts answer through methods.
    const facts = Object.assign(
      Object.create({ fontInfoOf: () => undefined }) as object,
      { cacheIdentity: "a" },
    ) as unknown as DocumentFontFacts;
    const origin = paragraph(10);
    origin.attrs = { ...origin.attrs, documentFontFacts: facts };
    const first = moved(origin, 0);
    noteParagraphOrigin(first, origin);
    const counter = counting();
    memoizedParagraphKey(first, 1, counter.compute);
    (facts as unknown as { cacheIdentity: string }).cacheIdentity = "b";
    const next = moved(origin, 3);
    noteParagraphOrigin(next, origin);
    memoizedParagraphKey(next, 1, counter.compute);
    expect(counter.calls()).toBe(2);
  });

  test("a block of unknown origin serializes", () => {
    const counter = counting();
    memoizedParagraphKey(paragraph(1), 1, counter.compute);
    memoizedParagraphKey(paragraph(1), 1, counter.compute);
    expect(counter.calls()).toBe(2);
  });
});
