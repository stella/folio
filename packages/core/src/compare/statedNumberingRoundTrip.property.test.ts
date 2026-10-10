import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { paragraphNumberingReference, NO_PARAGRAPH_NUMBERING } from "@stll/docx-core/model";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { expectParagraphBlock } from "../../../../test/paragraphBlock";
import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIParagraphBlock } from "../ai-edits/types";
import { createDocx } from "../docx/rezip";
import { generateRedlineDocx } from "../redline";
import { createEmptyDocument } from "../utils/createDocument";
import { INHERITED_PARAGRAPH_NUMBERING } from "./content-types";
import type { FolioContentStatedNumbering } from "./content-types";
import { compareDocx } from "./compare";

setDefaultTimeout(propertyTestTimeout(120_000));

const PARAGRAPH_TEXT = "The numbered clause remains intact.";
const ANCHOR_TEXT = "The stable anchor remains intact.";
const TRAILING_ANCHOR_TEXT = "The trailing anchor remains intact.";
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

const buildInheritedCollisionDocument = async ({
  statedNumbering,
  numFmt,
  styleId,
  includeAnchor = false,
  includeStyledParagraph = true,
  omitParagraphIds = false,
  duplicateParagraphIds = false,
  includeTrailingAnchor = false,
}: {
  statedNumbering: FolioContentStatedNumbering;
  numFmt: "decimal" | "lowerRoman" | "bullet";
  styleId: "BaseNumbered" | "StyleNumbered";
  includeAnchor?: boolean;
  includeStyledParagraph?: boolean;
  omitParagraphIds?: boolean;
  duplicateParagraphIds?: boolean;
  includeTrailingAnchor?: boolean;
}): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    ...(includeAnchor
      ? [
          {
            type: "paragraph" as const,
            ...(!omitParagraphIds && { paraId: "ABCD1234", textId: "ABCD1234" }),
            formatting: { styleId: "Normal" },
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: ANCHOR_TEXT }],
              },
            ],
          },
        ]
      : []),
    ...(includeStyledParagraph
      ? [
          {
            type: "paragraph" as const,
            ...(!omitParagraphIds && {
              paraId: duplicateParagraphIds ? "ABCD1234" : "1234ABCD",
              textId: "1234ABCD",
            }),
            formatting: {
              styleId,
              ...(statedNumbering.kind !== "inherit" && { numPr: statedNumbering }),
            },
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: PARAGRAPH_TEXT }],
              },
            ],
          },
        ]
      : []),
    ...(includeTrailingAnchor
      ? [
          {
            type: "paragraph" as const,
            ...(!omitParagraphIds && { paraId: "DCBA4321", textId: "DCBA4321" }),
            formatting: { styleId: "Normal" },
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: TRAILING_ANCHOR_TEXT }],
              },
            ],
          },
        ]
      : []),
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
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        levels: [
          {
            ilvl: 0,
            numFmt,
            lvlText: numFmt === "bullet" ? "•" : "%1.",
          },
        ],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  const buffer = await createDocx(document);
  if (!duplicateParagraphIds || !includeStyledParagraph) return buffer;
  // Serialization normalizes duplicate IDs. Restore the malformed producer
  // shape in the actual package so projection sees the duplicate identities.
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")?.async("text");
  expect(xml).toBeDefined();
  const firstId = xml?.match(/w14:paraId="([^"]+)"/u)?.at(1);
  expect(firstId).toBeDefined();
  let paragraphIndex = 0;
  zip.file(
    "word/document.xml",
    xml?.replace(/w14:paraId="[^"]+"/gu, (attribute) => {
      paragraphIndex += 1;
      return paragraphIndex === 2 ? `w14:paraId="${firstId}"` : attribute;
    }) ?? "",
  );
  return await zip.generateAsync({ type: "arraybuffer" });
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

const paragraphProjection = async (buffer: ArrayBuffer, text = PARAGRAPH_TEXT) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  const paragraph = expectParagraphBlock(
    reviewer.snapshot().blocks.find((block) => block.text === text),
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

const INHERITED_COLLISION_CASES = [
  {
    statedNumbering: INHERITED_PARAGRAPH_NUMBERING,
    targetMarker: "i.",
    targetNumbering: "lowerRoman",
  },
  {
    statedNumbering: { kind: "levelOnly", ilvl: 0 },
    targetMarker: "i.",
    targetNumbering: "lowerRoman",
  },
] as const;

const INSERTED_NUMBERING_CASES = [
  {
    statedNumbering: INHERITED_PARAGRAPH_NUMBERING,
    numFmt: "lowerRoman",
    targetMarker: "i.",
  },
  {
    statedNumbering: INHERITED_PARAGRAPH_NUMBERING,
    numFmt: "bullet",
    targetMarker: "•",
  },
  {
    statedNumbering: { kind: "levelOnly", ilvl: 0 },
    numFmt: "lowerRoman",
    targetMarker: "i.",
  },
  {
    statedNumbering: { kind: "levelOnly", ilvl: 0 },
    numFmt: "bullet",
    targetMarker: "•",
  },
] as const;

const INSERTED_COLLISION_CASES = [false, true].flatMap((omitParagraphIds) =>
  INSERTED_NUMBERING_CASES.map(({ statedNumbering, numFmt, targetMarker }) => ({
    statedNumbering,
    numFmt,
    targetMarker,
    omitParagraphIds,
  })),
);

test("accept and reject preserve style-only numbering across colliding definitions", async () => {
  await assertProperty(
    fc.asyncProperty(fc.constant(INHERITED_COLLISION_CASES), async (cases) => {
      const baseBuffer = await buildInheritedCollisionDocument({
        statedNumbering: INHERITED_PARAGRAPH_NUMBERING,
        numFmt: "decimal",
        styleId: "BaseNumbered",
      });
      const baseProjection = await paragraphProjection(baseBuffer);
      expect(baseProjection.statedNumbering).toEqual(INHERITED_PARAGRAPH_NUMBERING);
      expect(baseProjection.displayLabel).toBe("1.");

      for (const { statedNumbering, targetMarker, targetNumbering } of cases) {
        const targetBuffer = await buildInheritedCollisionDocument({
          statedNumbering,
          numFmt: targetNumbering,
          styleId: "StyleNumbered",
        });
        const targetProjection = await paragraphProjection(targetBuffer);
        expect(targetProjection.statedNumbering).toEqual(statedNumbering);
        expect(targetProjection.displayLabel).toBe(targetMarker);

        const compared = await compareDocx(baseBuffer, targetBuffer, COMPARE_OPTIONS);
        if (compared.isErr()) throw compared.error;
        expect(compared.value.verification).toEqual({ status: "verified" });

        const accepting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
        expect(accepting.acceptAll()).toBeGreaterThan(0);
        const acceptedBuffer = await accepting.toBuffer();
        const accepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
        const acceptedProjection = await paragraphProjection(acceptedBuffer);
        expect(acceptedProjection.statedNumbering).toEqual(statedNumbering);
        expect(acceptedProjection.styleId).not.toBe("BaseNumbered");
        expect(acceptedProjection.displayLabel).toBe(targetMarker);
        expect(acceptedProjection.listReference).toBeDefined();
        expect(accepted.readNumberingDefinitions()).toContainEqual(
          expect.objectContaining({
            numId: acceptedProjection.listReference?.numId,
            format: targetNumbering,
          }),
        );

        const rejecting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
        expect(rejecting.rejectAll()).toBeGreaterThan(0);
        const rejectedBuffer = await rejecting.toBuffer();
        const rejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
        const rejectedProjection = await paragraphProjection(rejectedBuffer);
        expect(rejectedProjection.statedNumbering).toEqual(INHERITED_PARAGRAPH_NUMBERING);
        expect(rejectedProjection.styleId).toBe("BaseNumbered");
        expect(rejectedProjection.displayLabel).toBe("1.");
        expect(rejectedProjection.listReference).toBeDefined();
        expect(rejected.readNumberingDefinitions()).toContainEqual(
          expect.objectContaining({
            numId: rejectedProjection.listReference?.numId,
            format: "decimal",
          }),
        );
      }
    }),
    { numRuns: 1 },
  );
});

test("compare and redline preserve inserted style-only numbering across collisions", async () => {
  await assertProperty(
    fc.asyncProperty(fc.constant(INSERTED_COLLISION_CASES), async (cases) => {
      const baseBuffer = await buildInheritedCollisionDocument({
        statedNumbering: INHERITED_PARAGRAPH_NUMBERING,
        numFmt: "decimal",
        styleId: "BaseNumbered",
        includeAnchor: true,
        includeStyledParagraph: false,
      });
      const baseProjection = await paragraphProjection(baseBuffer, ANCHOR_TEXT);
      expect(baseProjection.text).toBe(ANCHOR_TEXT);

      for (const { statedNumbering, numFmt, targetMarker, omitParagraphIds } of cases) {
        const targetBuffer = await buildInheritedCollisionDocument({
          statedNumbering,
          numFmt,
          styleId: "StyleNumbered",
          includeAnchor: true,
          omitParagraphIds,
        });
        const caseBaseBuffer = omitParagraphIds
          ? await buildInheritedCollisionDocument({
              statedNumbering: INHERITED_PARAGRAPH_NUMBERING,
              numFmt: "decimal",
              styleId: "BaseNumbered",
              includeAnchor: true,
              includeStyledParagraph: false,
              omitParagraphIds: true,
            })
          : baseBuffer;
        if (omitParagraphIds) {
          for (const fixtureBuffer of [caseBaseBuffer, targetBuffer]) {
            const zip = await JSZip.loadAsync(fixtureBuffer);
            const documentXml = await zip.file("word/document.xml")?.async("text");
            expect(documentXml).toBeDefined();
            expect(documentXml).not.toMatch(/w14:(?:paraId|textId)="/u);
          }
        }
        const targetProjection = await paragraphProjection(targetBuffer);
        expect(targetProjection.statedNumbering).toEqual(statedNumbering);
        expect(targetProjection.displayLabel).toBe(targetMarker);

        const generated = await Promise.all([
          compareDocx(caseBaseBuffer, targetBuffer, COMPARE_OPTIONS),
          generateRedlineDocx(caseBaseBuffer, targetBuffer),
        ]);
        const [compared, redline] = generated;
        if (compared.isErr()) throw compared.error;
        expect(compared.value.verification).toEqual({ status: "verified" });
        expect(redline.skipped).toEqual([]);

        for (const buffer of [compared.value.buffer, redline.buffer]) {
          const accepting = await FolioDocxReviewer.fromBuffer(buffer);
          expect(accepting.acceptAll()).toBeGreaterThan(0);
          const acceptedBuffer = await accepting.toBuffer();
          const accepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
          expect(accepted.snapshot().blocks.map((block) => block.text)).toEqual([
            ANCHOR_TEXT,
            PARAGRAPH_TEXT,
          ]);
          const acceptedProjection = await paragraphProjection(acceptedBuffer);
          expect(acceptedProjection.statedNumbering).toEqual(statedNumbering);
          expect(acceptedProjection.displayLabel).toBe(targetMarker);
          expect(acceptedProjection.listReference).toBeDefined();
          expect(accepted.readNumberingDefinitions()).toContainEqual(
            expect.objectContaining({
              numId: acceptedProjection.listReference?.numId,
              format: numFmt,
            }),
          );

          const rejecting = await FolioDocxReviewer.fromBuffer(buffer);
          expect(rejecting.rejectAll()).toBeGreaterThan(0);
          const rejectedBuffer = await rejecting.toBuffer();
          const rejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
          expect(rejected.snapshot().blocks.map((block) => block.text)).toEqual([ANCHOR_TEXT]);
          expect(rejected.readNumberingDefinitions()).toContainEqual(
            expect.objectContaining({ format: "decimal" }),
          );
        }
      }
    }),
    { numRuns: 1 },
  );
});

test("redline preserves insertion resources between unchanged anchors with duplicate or absent paragraph IDs", async () => {
  for (const identity of ["duplicate", "absent"] as const) {
    const identityOptions = {
      omitParagraphIds: identity === "absent",
      duplicateParagraphIds: identity === "duplicate",
      includeAnchor: true,
      includeTrailingAnchor: true,
    };
    const baseBuffer = await buildInheritedCollisionDocument({
      ...identityOptions,
      statedNumbering: INHERITED_PARAGRAPH_NUMBERING,
      numFmt: "decimal",
      styleId: "BaseNumbered",
      includeStyledParagraph: false,
    });
    const anchorProjections = await Promise.all(
      [ANCHOR_TEXT, TRAILING_ANCHOR_TEXT].map((text) => paragraphProjection(baseBuffer, text)),
    );
    for (const { statedNumbering, numFmt, targetMarker } of INSERTED_NUMBERING_CASES) {
      const targetBuffer = await buildInheritedCollisionDocument({
        ...identityOptions,
        statedNumbering,
        numFmt,
        styleId: "StyleNumbered",
      });
      const zip = await JSZip.loadAsync(targetBuffer);
      const xml = await zip.file("word/document.xml")?.async("text");
      expect(xml).toBeDefined();
      if (identity === "duplicate") {
        const ids = [...(xml?.matchAll(/w14:paraId="([^"]+)"/gu) ?? [])].map((match) =>
          match.at(1),
        );
        expect(ids).toHaveLength(3);
        expect(ids.at(0)).toBe(ids.at(1));
        expect(ids.at(2)).not.toBe(ids.at(0));
      } else {
        expect(xml).not.toMatch(/w14:(?:paraId|textId)="/u);
      }
      const revised = await FolioDocxReviewer.fromBuffer(targetBuffer);
      const inserted = expectParagraphBlock(
        revised.snapshot().blocks.find((block) => block.text === PARAGRAPH_TEXT),
      );
      expect(inserted.displayLabel).toBe(targetMarker);
      expect(inserted.idStability).toBe("positional");
      const redline = await generateRedlineDocx(baseBuffer, targetBuffer);
      expect(redline.skipped).toEqual([]);
      const accepting = await FolioDocxReviewer.fromBuffer(redline.buffer);
      expect(accepting.acceptAll()).toBeGreaterThan(0);
      const acceptedBuffer = await accepting.toBuffer();
      const accepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
      expect(accepted.snapshot().blocks.map((block) => block.text)).toEqual([
        ANCHOR_TEXT,
        PARAGRAPH_TEXT,
        TRAILING_ANCHOR_TEXT,
      ]);
      const projection = await paragraphProjection(acceptedBuffer);
      expect(projection.statedNumbering).toEqual(statedNumbering);
      expect(projection.displayLabel).toBe(targetMarker);
      expect(projection.styleId).not.toBe(inserted.styleId);
      expect(projection.listReference?.numId).not.toBe(inserted.listReference?.numId);
      expect(accepted.readNumberingDefinitions()).toContainEqual(
        expect.objectContaining({ numId: projection.listReference?.numId, format: numFmt }),
      );
      expect(
        await Promise.all(
          [ANCHOR_TEXT, TRAILING_ANCHOR_TEXT].map((text) =>
            paragraphProjection(acceptedBuffer, text),
          ),
        ),
      ).toEqual(anchorProjections);
      const rejecting = await FolioDocxReviewer.fromBuffer(redline.buffer);
      expect(rejecting.rejectAll()).toBeGreaterThan(0);
      const rejectedBuffer = await rejecting.toBuffer();
      const rejected = await FolioDocxReviewer.fromBuffer(rejectedBuffer);
      expect(rejected.snapshot().blocks.map((block) => block.text)).toEqual([
        ANCHOR_TEXT,
        TRAILING_ANCHOR_TEXT,
      ]);
      expect(
        await Promise.all(
          [ANCHOR_TEXT, TRAILING_ANCHOR_TEXT].map((text) =>
            paragraphProjection(rejectedBuffer, text),
          ),
        ),
      ).toEqual(anchorProjections);
      expect(rejected.readNumberingDefinitions()).toContainEqual(
        expect.objectContaining({ numId: 5, format: "decimal" }),
      );
    }
  }
});
