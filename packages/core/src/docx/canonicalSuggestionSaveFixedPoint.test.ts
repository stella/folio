import { expect, test } from "bun:test";
import { panic } from "better-result";
import JSZip from "jszip";
import {
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  OP_STORIES,
  REVISION_DECISIONS,
  type EditorIntent,
} from "@stll/docx-core/ops";

import { PARAGRAPH_MARK_CHANGE_KINDS } from "@stll/docx-core/model";
import { parseDocumentBody } from "./documentParser";
import { identityKeysIn, IDENTITY_SPACES } from "../../../docx-core/src/ops/ids";

import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import { serializeDocument } from "./serializer/documentSerializer";
import { buildStructuralDocumentPatch } from "./structuralXmlPatch";

const apply = (document: Document, intent: EditorIntent) => {
  const ids = createEditorIntentIdAllocator()(document, intent);
  const planned = compileEditorIntent(document, {
    intent,
    mode: {
      type: "suggesting",
      revision: { id: ids.revisionId, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
      newIds: ids.newIds,
    },
  });
  if (planned.isErr()) throw planned.error;
  const applied = applyDocumentOps(document, planned.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

// Exercise edit sequences and both save paths; a static serializer fixture misses
// metadata synthesized only when paragraph marks and structural splices are added.
test.each(["absent", "empty", "bold"] as const)(
  "pending edits retain exact content through full and structural saves (%s mark properties)",
  async (markProperties) => {
    const seed = createEmptyDocument({ initialText: "A😀B" });
    seed.package.document.content.push(
      ...createEmptyDocument({ initialText: "C" }).package.document.content,
    );
    for (const [index, paragraph] of seed.package.document.content.entries()) {
      if (paragraph.type !== "paragraph") panic("Expected paragraphs");
      paragraph.paraId = (index + 17).toString(16).padStart(8, "0").toUpperCase();
    }
    if (markProperties !== "absent") {
      for (const paragraph of seed.package.document.content) {
        if (paragraph.type !== "paragraph") panic("Expected paragraphs");
        paragraph.formatting = {
          ...paragraph.formatting,
          runProperties: markProperties === "empty" ? {} : { bold: true },
        };
      }
    }
    const source = await createDocx(seed);
    const original = await parseDocx(source, { preloadFonts: false, detectVariables: false });
    const first = original.package.document.content.at(0);
    const second = original.package.document.content.at(1);
    if (
      first?.type !== "paragraph" ||
      second?.type !== "paragraph" ||
      !first.paraId ||
      !second.paraId
    )
      panic("Expected identified paragraphs");
    const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });
    const inserted = apply(original, {
      type: "replaceText",
      from: at(first.paraId, 4),
      to: at(first.paraId, 4),
      text: "x",
    });
    const deleted = apply(inserted, {
      type: "replaceText",
      from: at(first.paraId, 1),
      to: at(first.paraId, 3),
      text: "",
    });
    const splitIds = createEditorIntentIdAllocator()(deleted, {
      type: "splitParagraph",
      at: at(first.paraId, 1),
    });
    const split = apply(deleted, {
      type: "splitParagraph",
      at: at(first.paraId, 1),
      newBlockId: splitIds.newBlockId,
    });
    const suggested = apply(split, {
      type: "joinParagraphs",
      story: OP_STORIES.MAIN,
      blockId: first.paraId,
      nextBlockId: second.paraId,
    });
    const fullSave = await createDocx(suggested);
    const zip = await JSZip.loadAsync(source);
    const originalXml = await zip.file("word/document.xml")?.async("text");
    if (!originalXml) panic("Expected document part");
    const patched = buildStructuralDocumentPatch({
      originalXml,
      serializedXml: serializeDocument(suggested),
      changedIds: new Set([first.paraId, second.paraId, splitIds.newBlockId]),
    });
    if (patched === null) panic("Expected structural save to handle the edit sequence");
    zip.file("word/document.xml", patched);
    const structuralSave = await zip.generateAsync({ type: "arraybuffer" });
    for (const saved of [fullSave, structuralSave]) {
      const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
      expect(reopened.package.document.content).toStrictEqual(suggested.package.document.content);
      for (const decision of Object.values(REVISION_DECISIONS)) {
        const resolution = {
          type: "resolveRevision",
          story: OP_STORIES.MAIN,
          decision,
          revisionIds: identityKeysIn(suggested.package.document.content).flatMap((key) =>
            key.startsWith(`${IDENTITY_SPACES.REVISION}:`)
              ? [Number(key.slice(IDENTITY_SPACES.REVISION.length + 1))]
              : [],
          ),
        } as const;
        const before = applyDocumentOps(suggested, [resolution]);
        const after = applyDocumentOps(reopened, [resolution]);
        if (before.isErr()) throw before.error;
        if (after.isErr()) throw after.error;
        expect(after.value.document.package.document.content).toStrictEqual(
          before.value.document.package.document.content,
        );
      }
    }
  },
);

test("every paragraph mark kind retains absent, empty and populated formatting", () => {
  for (const kind of [undefined, ...PARAGRAPH_MARK_CHANGE_KINDS]) {
    for (const properties of [undefined, {}, { bold: true }]) {
      for (const runInWithNext of [undefined, false, true]) {
        const original = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000011",
                  formatting: {
                    styleId: "Normal",
                    ...(runInWithNext === undefined ? {} : { runInWithNext }),
                    ...(properties === undefined ? {} : { runProperties: properties }),
                  },
                  ...(kind === undefined
                    ? {}
                    : { pPrMark: { kind, info: { id: 1, author: "Reviewer" } } }),
                  content: [
                    { type: "run", formatting: {}, content: [{ type: "text", text: "é😀" }] },
                  ],
                },
              ],
            },
          },
        } satisfies Document;
        const serialized = serializeDocument(original);
        for (const xml of [
          serialized,
          serialized
            .replaceAll("folio:emptyMarkProperties", "f:emptyMarkProperties")
            .replaceAll("xmlns:folio=", "xmlns:f="),
        ]) {
          const reopened = parseDocumentBody(xml);
          expect(reopened.content).toStrictEqual(original.package.document.content);
          expect(
            parseDocumentBody(serializeDocument({ package: { document: reopened } })).content,
          ).toStrictEqual(reopened.content);
        }
      }
    }
  }
});
