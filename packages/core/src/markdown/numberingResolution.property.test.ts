import { expect, setDefaultTimeout, spyOn, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { fromMarkdown } from "./fromMarkdown";
import { toMarkdown } from "./index";
import {
  createStyleResolver,
  resolveStyleParagraphNumbering,
  StyleResolver,
} from "../prosemirror/styles/styleResolver";
import { paragraphNumberingFromSlots } from "../docx/numberingReference";

setDefaultTimeout(propertyTestTimeout(5_000));

test("Markdown numbering avoids the full run-formatting style cascade", () => {
  const fullCascade = spyOn(StyleResolver.prototype, "resolveParagraphStyle");
  try {
    expect(toMarkdown(fromMarkdown("1. First\n2. Second")).trim()).toBe("1. First\n2. Second");
    expect(fullCascade).not.toHaveBeenCalled();
  } finally {
    fullCascade.mockRestore();
  }
});

test("repeated list definitions keep independent counters and observe edits between renders", async () => {
  await assertProperty(
    fc.property(fc.integer({ min: 2, max: 8 }), fc.integer({ min: 1, max: 20 }), (count, start) => {
      const document = fromMarkdown(
        Array.from({ length: count }, (_, index) => `${index + 1}. Item`).join("\n"),
      );
      const level = document.package.numbering?.abstractNums.at(0)?.levels.at(0);
      if (!level) return panic("Missing generated list definition.");
      for (const suffix of [".", ")", "."] as const) {
        level.lvlText = `%1${suffix}`;
        level.start = start;
        expect(toMarkdown(document).trim().split("\n")).toEqual(
          Array.from({ length: count }, (_, index) => `${start + index}${suffix} Item`),
        );
      }
    }),
    { numRuns: 30 },
  );
});

test("numbering-only style resolution stays bound to the full paragraph cascade", async () => {
  const numbering = fc.option(
    fc
      .record({ numId: fc.integer({ min: 0, max: 3 }), ilvl: fc.integer({ min: 0, max: 8 }) })
      .map(paragraphNumberingFromSlots),
    { nil: undefined },
  );
  await assertProperty(
    fc.property(numbering, numbering, numbering, (defaults, normal, custom) => {
      const resolver = createStyleResolver({
        docDefaults: { pPr: { numPr: defaults } },
        styles: [
          { styleId: "Normal", type: "paragraph", default: true, pPr: { numPr: normal } },
          { styleId: "Custom", type: "paragraph", pPr: { numPr: custom } },
        ],
      });
      for (const id of [undefined, "", "Normal", "Custom", "Unknown"]) {
        expect(resolveStyleParagraphNumbering(resolver, id)).toEqual(
          resolver.resolveParagraphStyle(id).paragraphFormatting?.numPr,
        );
      }
    }),
    { numRuns: 50 },
  );
});
