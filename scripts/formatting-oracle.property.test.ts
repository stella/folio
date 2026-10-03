import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  compareWithModel,
  expectOperation,
  modelOf,
  type Row,
} from "../test/consumer-scenarios/support/oracle";

test(
  "range formatting composes with restyling without permitting direct formatting leaks",
  () => {
    assertProperty(
      fc.property(
        fc.integer({ min: 2, max: 30 }),
        fc.constantFrom("bold", "italic", "underline"),
        fc.boolean(),
        fc.option(fc.boolean(), { nil: undefined }),
        fc.boolean(),
        fc.boolean(),
        (length, property, inherited, direct, styleExample, reverse) => {
          const text = "😀" + "x".repeat(length);
          const pre = {
            id: "068C5632",
            text,
            kind: "paragraph",
            styleId: "Normal",
            previewRuns: [
              {
                text,
                [property]: direct ?? false,
                ...(direct === undefined ? {} : { directFormatting: { [property]: direct } }),
              },
            ],
          } satisfies Row;
          const sample = {
            id: "078C57C5",
            text: "Heading",
            kind: "heading",
            headingLevel: 2,
            styleId: "Heading2",
            previewRuns: [{ text: "Heading", [property]: inherited }],
          } satisfies Row;
          const model = modelOf(styleExample ? [pre, sample] : [pre]);
          model.mode = "tracked-changes";
          const operations = [
            {
              type: "formatRange",
              range: {
                type: "textRange",
                story: "main",
                blockId: pre.id,
                startOffset: 0,
                endOffset: 2,
              },
              formatting: { [property]: true },
            },
            {
              type: "setBlockParagraphProperties",
              blockId: pre.id,
              properties: { styleId: "Heading2" },
            },
          ];
          for (const operation of reverse ? operations.toReversed() : operations)
            expectOperation(model, operation);
          const after = {
            id: pre.id,
            text,
            kind: "heading",
            headingLevel: 2,
            styleId: "Heading2",
            previewRuns: [
              { text: "😀", [property]: true, directFormatting: { [property]: true } },
              {
                text: text.slice(2),
                [property]: direct ?? inherited,
                ...(direct === undefined ? {} : { directFormatting: { [property]: direct } }),
              },
            ],
          } satisfies Row;
          const rows = styleExample ? [after, sample] : [after];
          expect(compareWithModel(model, rows)).toEqual([]);
          const missing = {
            ...after,
            previewRuns: [
              { text, [property]: inherited, directFormatting: { [property]: direct } },
            ],
          } satisfies Row;
          if (direct !== true)
            expect(
              compareWithModel(model, styleExample ? [missing, sample] : [missing]).length,
            ).toBeGreaterThan(0);
          const leaked = {
            ...after,
            previewRuns: [
              after.previewRuns[0],
              {
                text: text.slice(2),
                [property]: !(direct ?? inherited),
                directFormatting: { [property]: direct === undefined ? true : !direct },
              },
            ],
          } satisfies Row;
          expect(
            compareWithModel(model, styleExample ? [leaked, sample] : [leaked]).length,
          ).toBeGreaterThan(0);
        },
      ),
      { numRuns: 100 },
    );
  },
  propertyTestTimeout(5_000),
);
