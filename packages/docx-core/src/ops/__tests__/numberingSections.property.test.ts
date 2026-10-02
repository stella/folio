import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { serializeDocumentToDocx } from "../../serialize/docx";
import { paragraphNumberingReference } from "../../model/paragraphNumbering";
import { DOCUMENT_OP_TYPES } from "../types";
import type { DocumentOp } from "../types";
import { propertyConfig } from "../../../../../test/property-testing";
import { documentArbitrary } from "./documentArbitraries";

const paragraph: Paragraph = {
  type: "paragraph",
  paraId: "00000001",
  content: [{ type: "run", content: [{ type: "text", text: "item" }] }],
};

const base: Document = { package: { document: { content: [paragraph] } } };

test("numbering instance creation and its inverse preserve absent numbering and exact ids", () => {
  fc.assert(
    fc.property(
      fc.constantFrom(
        base,
        { ...base, package: { ...base.package, numbering: undefined } },
        {
          ...base,
          package: {
            ...base.package,
            numbering: {
              abstractNums: [
                {
                  abstractNumId: 7,
                  levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
                },
              ],
              nums: [],
            },
          },
        },
      ),
      fc.integer({ min: 1, max: 2_000_000_000 }),
      fc.integer({ min: 1, max: 2_000_000_000 }),
      (document, abstractNumId, numId) => {
        fc.pre(abstractNumId !== numId);
        const existingAbstract = document.package.numbering?.abstractNums.at(0);
        const abstractNumIdToUse = existingAbstract?.abstractNumId ?? abstractNumId;
        fc.pre(abstractNumIdToUse !== numId);
        const abstractNum =
          existingAbstract === undefined
            ? {
                abstractNumId: abstractNumIdToUse,
                levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
              }
            : undefined;
        const num = { numId, abstractNumId: abstractNumIdToUse };
        const forward = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
          ...(abstractNum === undefined ? {} : { abstractNum }),
          num,
        });
        expect(forward.isOk()).toBe(true);
        if (forward.isErr()) return;
        expect(forward.value.document.package.numbering?.nums).toEqual([num]);
        const undo = applyDocumentOps(forward.value.document, forward.value.inverse);
        expect(undo.isOk()).toBe(true);
        if (undo.isErr()) return;
        expect(undo.value.document).toStrictEqual(document);
        const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
        expect(redo.isOk()).toBe(true);
        if (redo.isErr()) return;
        expect(redo.value.document).toEqual(forward.value.document);
      },
    ),
    propertyConfig(),
  );
});

test("generated section endpoint edits restore paragraph, final, and section-view state", () => {
  fc.assert(
    fc.property(documentArbitrary, fc.integer({ min: 10_000, max: 20_000 }), (document, width) => {
      const topLevel = document.package.document.content.filter(
        (block): block is Paragraph => block.type === "paragraph",
      );
      const target = topLevel.at(0);
      if (target?.paraId === undefined) return;
      const endpoint = { type: "paragraph" as const, blockId: target.paraId };
      const current = target.sectionProperties;
      const properties = { ...current, pageWidth: width };
      const original = structuredClone(document);
      const paragraphEdit = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
        endpoint,
        expected:
          current === undefined
            ? { type: "absent" as const }
            : { type: "present" as const, value: current },
        properties,
      });
      expect(paragraphEdit.isOk()).toBe(true);
      if (paragraphEdit.isErr()) return;
      const paragraphUndo = applyDocumentOps(
        paragraphEdit.value.document,
        paragraphEdit.value.inverse,
      );
      expect(paragraphUndo.isOk()).toBe(true);
      if (paragraphUndo.isErr()) return;
      expect(paragraphUndo.value.document).toStrictEqual(original);
      const paragraphRedo = applyDocumentOps(
        paragraphUndo.value.document,
        paragraphUndo.value.inverse,
      );
      expect(paragraphRedo.isOk()).toBe(true);
      if (paragraphRedo.isErr()) return;
      expect(paragraphRedo.value.document).toStrictEqual(paragraphEdit.value.document);

      const finalProperties = document.package.document.finalSectionProperties;
      const finalEdit = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
        endpoint: { type: "final" },
        expected:
          finalProperties === undefined
            ? { type: "absent" as const }
            : { type: "present" as const, value: finalProperties },
        properties: { ...finalProperties, pageHeight: width + 1 },
      });
      expect(finalEdit.isOk()).toBe(true);
      if (finalEdit.isErr()) return;
      const finalUndo = applyDocumentOps(finalEdit.value.document, finalEdit.value.inverse);
      expect(finalUndo.isOk()).toBe(true);
      if (finalUndo.isErr()) return;
      expect(finalUndo.value.document).toStrictEqual(original);
      const finalRedo = applyDocumentOps(finalUndo.value.document, finalUndo.value.inverse);
      expect(finalRedo.isOk()).toBe(true);
      if (finalRedo.isErr()) return;
      expect(finalRedo.value.document).toStrictEqual(finalEdit.value.document);
    }),
    propertyConfig({ numRuns: 100 }),
  );
});

test("section endpoint edits preserve exact inverse for explicit and absent properties", () => {
  for (const endpoint of [
    { type: "paragraph" as const, blockId: "00000001" },
    { type: "final" as const },
  ]) {
    const forward = applyDocumentOp(base, {
      type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
      endpoint,
      expected: { type: "absent" },
      properties: { pageWidth: 12_240, marginTop: 720 },
    });
    expect(forward.isOk()).toBe(true);
    if (forward.isErr()) continue;
    const undo = applyDocumentOps(forward.value.document, forward.value.inverse);
    expect(undo.isOk()).toBe(true);
    if (undo.isErr()) continue;
    expect(undo.value.document).toEqual(base);
    const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
    expect(redo.isOk()).toBe(true);
    if (redo.isErr()) continue;
    expect(redo.value.document).toEqual(forward.value.document);
  }
});

test("numbering deletion refuses live paragraph and shared abstract references", () => {
  const abstractNum = {
    abstractNumId: 1,
    levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
  };
  const num = { numId: 1, abstractNumId: 1 };
  const referenced = {
    ...base,
    package: {
      ...base.package,
      document: {
        content: [
          { ...paragraph, formatting: { numPr: paragraphNumberingReference({ numId: 1 }) } },
        ],
      },
      numbering: { abstractNums: [abstractNum], nums: [num] },
    },
  };
  const liveReference = applyDocumentOp(referenced, {
    type: DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE,
    num,
    abstractNum,
  });
  expect(liveReference.isErr()).toBe(true);

  const sharedAbstract = {
    ...base,
    package: {
      ...base.package,
      numbering: { abstractNums: [abstractNum], nums: [num, { numId: 2, abstractNumId: 1 }] },
    },
  };
  const stillUsed = applyDocumentOp(sharedAbstract, {
    type: DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE,
    num,
    abstractNum,
  });
  expect(stillUsed.isErr()).toBe(true);
});

test("numbering and section inverses restore byte-equivalent DOCX parts", async () => {
  const originalZip = await JSZip.loadAsync(await serializeDocumentToDocx(base));
  const originalDocumentXml = await originalZip.file("word/document.xml")?.async("uint8array");
  const operations = [
    {
      type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
      num: { numId: 8, abstractNumId: 8 },
      abstractNum: {
        abstractNumId: 8,
        levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
      },
    },
    {
      type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
      endpoint: { type: "paragraph" as const, blockId: "00000001" },
      properties: { pageWidth: 12_240, marginTop: 720 },
    },
  ] as const;
  let current = base;
  const inverses: DocumentOp[][] = [];
  for (const operation of operations) {
    const applied = applyDocumentOp(current, operation);
    expect(applied.isOk()).toBe(true);
    if (applied.isErr()) return;
    current = applied.value.document;
    inverses.push([...applied.value.inverse]);
  }
  const restored = applyDocumentOps(current, inverses.toReversed().flat());
  expect(restored.isOk()).toBe(true);
  if (restored.isErr()) return;
  const restoredZip = await JSZip.loadAsync(await serializeDocumentToDocx(restored.value.document));
  expect(await restoredZip.file("word/document.xml")?.async("uint8array")).toEqual(
    originalDocumentXml,
  );
  expect(restoredZip.file("word/numbering.xml")).toBeNull();
});
