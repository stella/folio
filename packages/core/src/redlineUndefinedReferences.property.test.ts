import { expect, setDefaultTimeout, test } from "bun:test";
import { paragraphNumberingReference, NO_PARAGRAPH_NUMBERING } from "@stll/docx-core/model";
import { panic } from "better-result";
import fc from "fast-check";
import JSZip from "jszip";

import { assertProperty, propertyTestTimeout } from "../../../test/property-testing";
import { expectParagraphBlock } from "../../../test/paragraphBlock";
import { FolioDocxReviewer } from "./ai-edits/headless";
import { createDocx } from "./docx/rezip";
import { generateRedlineDocx } from "./redline";
import type { Paragraph } from "./types/document";
import { createEmptyDocument } from "./utils/createDocument";
import { alignFolioBlocks } from "./version-comparison";

const PROPERTY_BUDGET = 30_000;
const UNKNOWN_STYLE = "UndefinedInsertionStyle";
const UNKNOWN_NUM_ID = 5;
const INSERTED_ID = "12000004";
const MAIN_STORY = { type: "main" } as const;

setDefaultTimeout(propertyTestTimeout(PROPERTY_BUDGET));

type UndefinedReferenceFixtureOptions = {
  side: "base" | "revised";
  resource: "style" | "numbering";
  definition: "globallyUndefined" | "baseDefined";
  independent: boolean;
  suffix: string;
};

const paragraph = (paraId: string, text: string) =>
  ({
    type: "paragraph",
    paraId,
    textId: paraId,
    content: [{ type: "run", content: [{ type: "text", text }] }],
  }) as const satisfies Paragraph;

const undefinedReferenceFixture = ({
  side,
  resource,
  definition,
  independent,
  suffix,
}: UndefinedReferenceFixtureOptions) => {
  const document = createEmptyDocument();
  const defineCollision = side === "base" && definition === "baseDefined";
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      ...(independent
        ? [
            {
              type: "paragraph" as const,
              styleId: "IndependentCollision",
              name: "Independent collision",
              rPr: side === "base" ? { bold: true } : { italic: true },
            },
          ]
        : []),
      ...(defineCollision && resource === "style"
        ? [
            {
              type: "paragraph" as const,
              styleId: UNKNOWN_STYLE,
              name: "Base collision style",
              rPr: { bold: true },
            },
          ]
        : []),
    ],
  };
  if (defineCollision && resource === "numbering") {
    document.package.numbering = {
      abstractNums: [
        { abstractNumId: UNKNOWN_NUM_ID, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] },
      ],
      nums: [{ numId: UNKNOWN_NUM_ID, abstractNumId: UNKNOWN_NUM_ID }],
    };
  }
  document.package.document.content = [
    paragraph("11000001", `Leading anchor ${suffix}`),
    ...(side === "base" ? [paragraph("11000002", `Removed original ${suffix}`)] : []),
    paragraph("11000003", `Middle anchor ${suffix}`),
    ...(side === "revised"
      ? [
          {
            ...paragraph(INSERTED_ID, `Replacement ${suffix}`),
            formatting:
              resource === "style"
                ? { styleId: UNKNOWN_STYLE }
                : { numPr: paragraphNumberingReference({ numId: UNKNOWN_NUM_ID, ilvl: 0 }) },
          },
        ]
      : []),
    paragraph("11000005", `Trailing anchor ${suffix}`),
    ...(side === "revised" && independent
      ? [
          {
            ...paragraph("12000006", `Independent insertion ${suffix}`),
            formatting: { styleId: "IndependentCollision" },
          },
        ]
      : []),
  ];
  return createDocx(document);
};

test("undefined inserted references stay dangling or are cleared before binding base resources", async () => {
  await assertProperty(
    fc.asyncProperty(fc.stringMatching(/^[a-z]{1,20}$/u), async (suffix) => {
      for (const resource of ["style", "numbering"] as const) {
        for (const definition of ["globallyUndefined", "baseDefined"] as const) {
          for (const independent of [false, true]) {
            const fixture = { resource, definition, independent, suffix };
            const base = await undefinedReferenceFixture({ ...fixture, side: "base" });
            const revised = await undefinedReferenceFixture({ ...fixture, side: "revised" });
            const baseBytes = new Uint8Array(base).slice();
            const revisedBytes = new Uint8Array(revised).slice();
            const before = await FolioDocxReviewer.fromBuffer(base);
            const target = await FolioDocxReviewer.fromBuffer(revised);
            const targetSnapshot = target.snapshot();
            const insertedSource = expectParagraphBlock(
              targetSnapshot.blocks.find(({ text }) => text === `Replacement ${suffix}`),
            );
            const insertionAnchor = targetSnapshot.anchors[insertedSource.id];
            if (!insertionAnchor) panic("Replacement fixture lost its source anchor");
            const events = alignFolioBlocks(before.snapshot().blocks, targetSnapshot.blocks);
            expect(events.filter(({ type }) => type === "baseOnly")).toHaveLength(1);
            expect(events.filter(({ type }) => type === "revisedOnly")).toHaveLength(
              independent ? 2 : 1,
            );
            const result = await generateRedlineDocx(base, revised);
            expect(result.skipped).toEqual([]);
            expect(result.unprocessedStories).toEqual([]);
            // Derive warnings from authored source references and original
            // full-story coordinates, independently of the importer result.
            const warningLocation = { paragraphPosition: insertionAnchor.from, story: MAIN_STORY };
            const expectedWarning = (
              resource === "style"
                ? { ...warningLocation, kind: "style" as const, id: UNKNOWN_STYLE }
                : { ...warningLocation, kind: "numbering" as const, id: UNKNOWN_NUM_ID }
            ) satisfies (typeof result.referenceWarnings)[number];
            expect(result.referenceWarnings).toEqual(
              definition === "baseDefined" ? [expectedWarning] : [],
            );
            expect(new Uint8Array(base)).toEqual(baseBytes);
            expect(new Uint8Array(revised)).toEqual(revisedBytes);
            const accepting = await FolioDocxReviewer.fromBuffer(result.buffer);
            expect(accepting.acceptAll()).toBeGreaterThan(0);
            const acceptedBuffer = await accepting.toBuffer();
            const accepted = await FolioDocxReviewer.fromBuffer(acceptedBuffer);
            expect(accepted.snapshot().blocks.map(({ text }) => text)).toEqual(
              targetSnapshot.blocks.map(({ text }) => text),
            );
            const inserted = expectParagraphBlock(
              accepted.snapshot().blocks.find(({ text }) => text === insertedSource.text),
            );
            expect(inserted.displayLabel).toBeUndefined();
            expect(inserted.listReference).toEqual(
              resource === "numbering" && definition === "globallyUndefined"
                ? { numId: UNKNOWN_NUM_ID, level: 0 }
                : undefined,
            );
            expect(inserted.previewRuns?.some(({ bold }) => bold === true) ?? false).toBe(false);
            if (resource === "style") {
              const zip = await JSZip.loadAsync(acceptedBuffer);
              const xml = await zip.file("word/document.xml")?.async("text");
              expect(xml).toBeDefined();
              const authoredStyle = new RegExp(`<w:pStyle\\b[^>]*\\bw:val="${UNKNOWN_STYLE}"`, "u");
              if (definition === "globallyUndefined") {
                expect(xml).toMatch(authoredStyle);
                expect(inserted.styleId).toBe(UNKNOWN_STYLE);
              } else {
                expect(xml).not.toMatch(authoredStyle);
                expect(inserted.styleId).not.toBe(UNKNOWN_STYLE);
              }
            } else {
              expect(inserted.statedNumbering).toEqual(
                definition === "globallyUndefined"
                  ? paragraphNumberingReference({ numId: UNKNOWN_NUM_ID, ilvl: 0 })
                  : NO_PARAGRAPH_NUMBERING,
              );
            }
            const rejecting = await FolioDocxReviewer.fromBuffer(result.buffer);
            expect(rejecting.rejectAll()).toBeGreaterThan(0);
            const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
            expect(rejected.snapshot().blocks.map(({ text }) => text)).toEqual(
              before.snapshot().blocks.map(({ text }) => text),
            );
          }
        }
      }
    }),
    // Every generated text exercises both resource types and definition states,
    // with and without an unrelated insertion through real DOCX round trips.
    { numRuns: 4 },
  );
});
