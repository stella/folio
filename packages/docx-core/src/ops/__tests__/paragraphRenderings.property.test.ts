/**
 * Paragraphs whose list rendering was normalized take their paragraphs'
 * places and keep the document's standing against the seed contract.
 *
 * A list normalization replaces a paragraph by one that differs only in its
 * list rendering and the formatting fields that rendering derives. Generated
 * documents get such replacements of chosen story paragraphs: each must stand
 * where its paragraph stood, and the result must be known to meet the
 * contract exactly when the document was, with the contract checked afresh
 * agreeing. A replacement that changes any other field is refused.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import { storyParagraphs } from "../blocks";
import {
  contractViolation,
  knownToMeetContract,
  validateOpsDocument,
} from "../contract";
import { withParagraphRenderings } from "../paragraphCopies";
import { documentStories, storyBody } from "../stories";
import { OP_STORIES } from "../types";
import { documentArbitrary } from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(120_000));

const renderingArbitrary = fc.record({
  marker: fc.constantFrom("1.", "a)", "•"),
  indentLeft: fc.option(fc.integer({ min: 0, max: 2880 }), { nil: undefined }),
  hanging: fc.boolean(),
  fromStyle: fc.boolean(),
  dropFormatting: fc.boolean(),
});
type Rendering = fc.RecordValue<typeof renderingArbitrary>;

/** `paragraph` with its list rendering and the fields it derives normalized anew. */
const renderedOf = (paragraph: Paragraph, rendering: Rendering): Paragraph => {
  const formatting = { ...paragraph.formatting };
  if (rendering.indentLeft === undefined) delete formatting.indentLeft;
  else formatting.indentLeft = rendering.indentLeft;
  if (rendering.hanging) {
    formatting.hangingIndent = true;
    formatting.indentFirstLine = -360;
  }
  if (rendering.fromStyle)
    formatting.numPrFromStyle = { kind: "reference", numId: 7, ilvl: 0 };
  else delete formatting.numPrFromStyle;
  const rendered: Paragraph = {
    ...paragraph,
    listRendering: {
      numId: 7,
      level: 0,
      marker: rendering.marker,
      isBullet: rendering.marker === "•",
    },
  };
  if (
    rendering.dropFormatting &&
    Object.keys(formatting).every((field) =>
      [
        "numPrFromStyle",
        "indentLeft",
        "indentRight",
        "indentFirstLine",
        "hangingIndent",
      ].includes(field),
    )
  )
    delete rendered.formatting;
  else rendered.formatting = formatting;
  return rendered;
};

describe("normalized list renderings", () => {
  test("stand in their paragraphs' places and keep the contract", () => {
    const law = (counts: { rendered: number }) =>
      fc.property(
        documentArbitrary,
        fc.nat(),
        fc.array(fc.tuple(fc.nat(), renderingArbitrary), {
          minLength: 1,
          maxLength: 4,
        }),
        fc.boolean(),
        (generated, storyChoice, choices, checked) => {
          const document: Document = checked
            ? generated
            : { ...generated, package: { ...generated.package } };
          if (checked) expect(validateOpsDocument(document).isOk()).toBe(true);
          const stories = documentStories(document);
          const story = stories[storyChoice % stories.length] ?? stories[0];
          if (story === undefined) return;
          const located = storyParagraphs(storyBody(document, story));
          if (located.length === 0) return;
          const renderings = new Map<Paragraph, Paragraph>();
          for (const [choice, rendering] of choices) {
            const paragraph = located[choice % located.length]?.paragraph;
            if (paragraph !== undefined && !renderings.has(paragraph))
              renderings.set(paragraph, renderedOf(paragraph, rendering));
          }
          const result = withParagraphRenderings({
            document,
            story,
            renderings,
          });
          // A generated story holds each paragraph object at one place.
          if (result === undefined)
            throw new Error("A rendered paragraph stood at two places.");
          const after = storyParagraphs(storyBody(result, story));
          expect(after.length).toBe(located.length);
          for (const [index, { paragraph, list }] of located.entries()) {
            expect(after[index]?.list).toEqual(list);
            expect(after[index]?.paragraph).toBe(
              renderings.get(paragraph) ?? paragraph,
            );
          }
          expect(knownToMeetContract(result)).toBe(
            knownToMeetContract(document),
          );
          if (checked) expect(contractViolation(result)).toBeUndefined();
          counts.rendered += renderings.size;
        },
      );
    assertProperty(law({ rendered: 0 }), {
      id: "normalized list renderings keep places and the contract",
      numRuns: 300,
    });
    // The property's run count, which the threshold was set for.
    const COVERAGE_RUNS = 300;
    const counts = { rendered: 0 };
    checkCoverageSample(law(counts), COVERAGE_RUNS);
    expect(counts.rendered).toBeGreaterThan(300);
  });

  test("a replacement that changes any other field is refused", () => {
    const law = (counts: { refused: number }) =>
      fc.property(documentArbitrary, fc.nat(), (document, at) => {
        const paragraphs = storyParagraphs(
          storyBody(document, OP_STORIES.MAIN),
        );
        const paragraph =
          paragraphs[at % Math.max(1, paragraphs.length)]?.paragraph;
        if (paragraph === undefined) return;
        const refusedFor = (replacement: Paragraph) =>
          expect(() =>
            withParagraphRenderings({
              document,
              story: OP_STORIES.MAIN,
              renderings: new Map([[paragraph, replacement]]),
            }),
          ).toThrow("must keep every other field");
        refusedFor({ ...paragraph, content: [] });
        refusedFor({ ...paragraph, paraId: "7EEEEEEE" });
        refusedFor({
          ...paragraph,
          formatting: {
            ...paragraph.formatting,
            alignment:
              paragraph.formatting?.alignment === "right" ? "left" : "right",
          },
        });
        refusedFor({
          ...paragraph,
          formatting: {
            ...paragraph.formatting,
            keepNext: paragraph.formatting?.keepNext !== true,
          },
        });
        counts.refused += 1;
      });
    assertProperty(law({ refused: 0 }), {
      id: "normalized list renderings refuse other changes",
      numRuns: 100,
    });
    // The property's run count, which the threshold was set for.
    const COVERAGE_RUNS = 100;
    const counts = { refused: 0 };
    checkCoverageSample(law(counts), COVERAGE_RUNS);
    expect(counts.refused).toBeGreaterThan(50);
  });

  test("leave a paragraph that stands at two places to be replaced by place", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      paraId: "00000001",
      content: [],
    };
    const document: Document = {
      package: { document: { content: [paragraph, paragraph] } },
    };
    expect(
      withParagraphRenderings({
        document,
        story: OP_STORIES.MAIN,
        renderings: new Map([
          [
            paragraph,
            renderedOf(paragraph, {
              marker: "1.",
              indentLeft: undefined,
              hanging: false,
              fromStyle: false,
              dropFormatting: true,
            }),
          ],
        ]),
      }),
    ).toBeUndefined();
  });
});
