import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { ParagraphContent, Run, TextFormatting } from "../model/document";
import { normalizeRunSegmentation } from "./runNormalization";

setDefaultTimeout(propertyTestTimeout(30_000));

const formattingArbitrary = fc.record(
  {
    bold: fc.boolean(),
    italic: fc.boolean(),
    color: fc.option(fc.constantFrom({ rgb: "112233" }, { rgb: "AABBCC" }), {
      nil: undefined,
    }),
  },
  { noNullPrototype: true },
);

const runArbitrary = fc
  .tuple(fc.string({ minLength: 1 }), formattingArbitrary)
  .map(([text, formatting]): Run => ({
    type: "run",
    content: [{ type: "text", text }],
    formatting,
  }));

const carrierRunArbitrary = formattingArbitrary.map((formatting): Run => ({
  type: "run",
  content: [],
  formatting,
}));

const contentArbitrary = fc.array(
  fc.oneof(
    runArbitrary,
    carrierRunArbitrary,
    fc
      .integer({ min: 1, max: 4 })
      .map((id): ParagraphContent => ({ type: "commentRangeStart", id })),
    fc.integer({ min: 1, max: 4 }).map((id): ParagraphContent => ({ type: "commentRangeEnd", id })),
    runArbitrary.map((run): ParagraphContent => ({
      type: "insertion",
      info: { id: 10, author: "Reviewer" },
      content: [run],
    })),
  ),
  { minLength: 1 },
);

const runFacts = (
  content: readonly ParagraphContent[],
): { text: string; formatting?: TextFormatting }[] =>
  content.flatMap((item) => {
    if (item.type === "run") {
      return [
        {
          text: item.content.map((part) => (part.type === "text" ? part.text : "")).join(""),
          formatting: item.formatting,
        },
      ];
    }
    if (item.type === "insertion") return runFacts(item.content);
    return [];
  });

test("canonical run segmentation is idempotent and preserves text and properties", () => {
  fc.assert(
    fc.property(contentArbitrary, (content) => {
      const once = normalizeRunSegmentation(content);
      expect(normalizeRunSegmentation(once)).toEqual(once);

      const before = runFacts(content);
      const after = runFacts(once);
      expect(after.map(({ text }) => text).join("")).toBe(before.map(({ text }) => text).join(""));
      for (const fact of before) {
        expect(
          after.some(
            ({ text, formatting }) =>
              text.includes(fact.text) &&
              JSON.stringify(formatting) === JSON.stringify(fact.formatting),
          ),
        ).toBe(true);
      }
    }),
    propertyConfig(),
  );
});

test("a comment range marker is a run boundary", () => {
  const formatting = { bold: true };
  const normalized = normalizeRunSegmentation([
    { type: "run", formatting, content: [{ type: "text", text: "left" }] },
    { type: "commentRangeStart", id: 1 },
    { type: "run", formatting, content: [{ type: "text", text: "right" }] },
  ]);

  expect(normalized.map(({ type }) => type)).toEqual(["run", "commentRangeStart", "run"]);
});

test("a payload-pending carrier is a run boundary", () => {
  const formatting = { bold: true };
  const carrier = { type: "run", formatting, content: [] } satisfies Run;
  const normalized = normalizeRunSegmentation([
    { type: "run", formatting, content: [{ type: "text", text: "left" }] },
    carrier,
    { type: "run", formatting, content: [{ type: "text", text: "right" }] },
  ]);

  expect(normalized).toEqual([
    { type: "run", formatting, content: [{ type: "text", text: "left" }] },
    carrier,
    { type: "run", formatting, content: [{ type: "text", text: "right" }] },
  ]);
});
