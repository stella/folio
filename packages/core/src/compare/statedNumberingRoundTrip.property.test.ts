import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { paragraphNumberingReference, NO_PARAGRAPH_NUMBERING } from "@stll/docx-core/model";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { expectParagraphBlock } from "../../../../test/paragraphBlock";
import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIParagraphBlock } from "../ai-edits/types";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { INHERITED_PARAGRAPH_NUMBERING } from "./content-types";
import type { FolioContentStatedNumbering } from "./content-types";
import { compareDocx } from "./compare";

setDefaultTimeout(propertyTestTimeout(120_000));

const PARAGRAPH_TEXT = "The numbered clause remains intact.";
const NUMBER_IDS = [5, 6, 7, 8] as const;
const NUMBER_LEVELS = [0, 1] as const;
const COMPARE_OPTIONS = {
  author: "compare",
  timestamp: "2026-09-13T00:00:00.000Z",
} as const;

const numbering = {
  abstractNums: NUMBER_IDS.map((abstractNumId) => ({
    abstractNumId,
    levels: NUMBER_LEVELS.map((ilvl) => ({
      ilvl,
      numFmt: "decimal" as const,
      lvlText: ilvl === 0 ? "%1." : "%1.%2.",
    })),
  })),
  nums: NUMBER_IDS.map((numId) => ({ numId, abstractNumId: numId })),
};

const buildDocument = async (
  styleId: "BaseNumbered" | "Normal" | "StyleNumbered",
  statedNumbering: FolioContentStatedNumbering,
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "1234ABCD",
      textId: "1234ABCD",
      formatting: {
        styleId,
        ...(statedNumbering.kind !== "inherit" && { numPr: statedNumbering }),
      },
      content: [
        {
          type: "run",
          content: [{ type: "text", text: PARAGRAPH_TEXT }],
        },
      ],
    },
  ];
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      {
        type: "paragraph",
        styleId: "BaseNumbered",
        name: "Base Numbered",
        pPr: { numPr: paragraphNumberingReference({ numId: 5 }) },
      },
      {
        type: "paragraph",
        styleId: "StyleNumbered",
        name: "Style Numbered",
        pPr: { numPr: paragraphNumberingReference({ numId: 5 }) },
      },
    ],
  };
  document.package.numbering = numbering;
  return await createDocx(document);
};

const TARGET_STATED_NUMBERINGS: readonly FolioContentStatedNumbering[] = [
  INHERITED_PARAGRAPH_NUMBERING,
  NO_PARAGRAPH_NUMBERING,
  { kind: "levelOnly", ilvl: 0 },
  { kind: "levelOnly", ilvl: 1 },
  paragraphNumberingReference({ numId: 5 }),
  paragraphNumberingReference({ numId: 6, ilvl: 0 }),
  paragraphNumberingReference({ numId: 7, ilvl: 1 }),
  paragraphNumberingReference({ numId: 8 }),
];

const ROUND_TRIP_CASES = (["Normal", "StyleNumbered"] as const).flatMap((targetStyle) =>
  TARGET_STATED_NUMBERINGS.map((targetNumbering) => ({
    targetStyle,
    targetNumbering,
  })),
);

const paragraphProjection = async (buffer: ArrayBuffer) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const paragraph = expectParagraphBlock(
    reviewer.snapshot().blocks.find(({ text }) => text === PARAGRAPH_TEXT),
  );
  return {
    kind: paragraph.kind,
    text: paragraph.text,
    ...(paragraph.styleId !== undefined && { styleId: paragraph.styleId }),
    ...(paragraph.displayLabel !== undefined && { displayLabel: paragraph.displayLabel }),
    ...(paragraph.listReference !== undefined && { listReference: paragraph.listReference }),
    statedNumbering: paragraph.statedNumbering,
  } satisfies Pick<
    FolioAIParagraphBlock,
    "kind" | "text" | "styleId" | "displayLabel" | "listReference" | "statedNumbering"
  >;
};

test("accept and reject preserve authored numbering independently from rendered membership", async () => {
  await assertProperty(
    fc.asyncProperty(fc.constant(ROUND_TRIP_CASES), async (cases) => {
      const baseBuffer = await buildDocument(
        "BaseNumbered",
        paragraphNumberingReference({ numId: 5, ilvl: 0 }),
      );
      const baseProjection = await paragraphProjection(baseBuffer);
      for (const { targetStyle, targetNumbering } of cases) {
        // The base's direct reference matches its style's reference. Every
        // target state therefore exercises an authored transition while the
        // style dimension independently changes the inherited tier. The base
        // reference also pins explicit-reference-to-inherit and same-style-id
        // cases, where source provenance differs while the visible list can match.
        const targetBuffer = await buildDocument(targetStyle, targetNumbering);
        const targetProjection = await paragraphProjection(targetBuffer);
        const compared = await compareDocx(baseBuffer, targetBuffer, COMPARE_OPTIONS);
        if (compared.isErr()) throw compared.error;
        expect(compared.value.verification).toEqual({ status: "verified" });

        const accepting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
        expect(accepting.acceptAll()).toBeGreaterThan(0);
        const acceptedProjection = await paragraphProjection(await accepting.toBuffer());
        expect(acceptedProjection).toEqual(targetProjection);

        const rejecting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
        expect(rejecting.rejectAll()).toBeGreaterThan(0);
        const rejectedProjection = await paragraphProjection(await rejecting.toBuffer());
        expect(rejectedProjection).toEqual(baseProjection);
      }
    }),
    { numRuns: 1 },
  );
});
