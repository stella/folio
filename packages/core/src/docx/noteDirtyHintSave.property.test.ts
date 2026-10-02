/**
 * Dirty note paragraph ids are hints, not proof of a semantic change. The note
 * save oracle must distinguish stale hints from edits while preserving source
 * note content and reaching a save fixed point. The previous save tests only
 * hinted real edits and never exercised unchanged source-less paragraph hints.
 */
import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx, repackDocx, repackDocxFromRaw } from "./rezip";
import { unzipDocx } from "./unzip";

const NOTE_KINDS = ["footnote", "endnote"] as const;
const SAVE_PATHS = ["repack", "raw"] as const;
const NOTE_PARA_IDS = ["71000001", "71000002", "71000003"] as const;
const STALE_PARA_ID = "71000004";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const setText = (block: Paragraph, text: string): void => {
  const content = block.content
    .flatMap((item) => (item.type === "run" ? item.content : []))
    .find((item) => item.type === "text");
  if (content?.type !== "text") panic("Expected fixture text");
  content.text = text;
};

const readPart = async (buffer: ArrayBuffer, path: string): Promise<string> => {
  const entry = (await JSZip.loadAsync(buffer)).file(path);
  if (!entry) panic(`Expected fixture part ${path}`);
  return entry.async("text");
};

const createFixture = async (kind: (typeof NOTE_KINDS)[number]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument({ initialText: "Body text" });
  const body = document.package.document.content.at(0);
  if (body?.type !== "paragraph") panic("Expected fixture body paragraph");
  body.content.push({
    type: "run",
    content: [{ type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: 1 }],
  });
  const content = NOTE_PARA_IDS.map((id, index) => paragraph(id, `Note paragraph ${index}`));
  if (kind === "footnote") {
    document.package.footnotes = [{ type: "footnote", id: 1, content }];
  } else {
    document.package.endnotes = [{ type: "endnote", id: 1, content }];
  }
  return createDocx(document);
};

describe("note dirty hints preserve semantic edits and unchanged parts", () => {
  for (const kind of NOTE_KINDS) {
    for (const savePath of SAVE_PATHS) {
      test(
        `${kind}: ${savePath} ignores stale dirty subsets without losing body edits`,
        async () => {
          const originalWithIds = await createFixture(kind);
          const partPath = `word/${kind}s.xml`;
          const sourceZip = await JSZip.loadAsync(originalWithIds);
          sourceZip.file(
            partPath,
            (await readPart(originalWithIds, partPath)).replace(/ w14:paraId="[^"]+"/gu, ""),
          );
          const originalWithoutIds = await sourceZip.generateAsync({ type: "arraybuffer" });
          await fc.assert(
            fc.asyncProperty(
              fc.subarray(NOTE_PARA_IDS),
              fc.boolean(),
              fc.boolean(),
              fc.boolean(),
              fc.option(fc.integer({ min: 0, max: NOTE_PARA_IDS.length - 1 }), { nil: null }),
              async (dirtyIds, includeMissingId, editBody, sourceHasIds, editedNoteIndex) => {
                const original = sourceHasIds ? originalWithIds : originalWithoutIds;
                const originalPart = await readPart(original, partPath);
                const document = await parseDocx(original, { preloadFonts: false });
                const notes =
                  kind === "footnote" ? document.package.footnotes : document.package.endnotes;
                expect(notes?.at(0)?.content).toHaveLength(NOTE_PARA_IDS.length);
                if (!sourceHasIds) {
                  for (const [index, block] of (notes?.at(0)?.content ?? []).entries()) {
                    if (block.type !== "paragraph") panic("Expected note fixture paragraph");
                    block.paraId = NOTE_PARA_IDS.at(index);
                  }
                }
                const body = document.package.document.content.at(0);
                if (body?.type !== "paragraph") panic("Expected parsed fixture body");
                if (editBody) setText(body, "Changed body text");
                const changedNoteParaIds = new Set<string>(dirtyIds);
                if (includeMissingId) changedNoteParaIds.add(STALE_PARA_ID);
                if (editedNoteIndex !== null) {
                  const noteParagraph = notes?.at(0)?.content.at(editedNoteIndex);
                  const editedParaId = NOTE_PARA_IDS.at(editedNoteIndex);
                  if (noteParagraph?.type !== "paragraph" || !editedParaId)
                    panic("Expected note edit target");
                  setText(noteParagraph, "Changed note text");
                  changedNoteParaIds.add(editedParaId);
                }
                const options = { changedNoteParaIds, updateModifiedDate: false };
                const saved =
                  savePath === "repack"
                    ? await repackDocx(document, options)
                    : await repackDocxFromRaw(document, await unzipDocx(original), options);
                const savedPart = await readPart(saved, partPath);
                if (sourceHasIds && editedNoteIndex === null && dirtyIds.length === 0) {
                  expect(savedPart).toBe(originalPart);
                }
                const reparsed = await parseDocx(saved, { preloadFonts: false });
                const savedNotes =
                  kind === "footnote" ? reparsed.package.footnotes : reparsed.package.endnotes;
                expect(
                  savedNotes?.at(0)?.content.map((block) =>
                    (block.type === "paragraph" ? block.content : [])
                      .flatMap((item) => (item.type === "run" ? item.content : []))
                      .map((item) => (item.type === "text" ? item.text : ""))
                      .join(""),
                  ),
                ).toEqual(
                  NOTE_PARA_IDS.map((_, index) =>
                    index === editedNoteIndex ? "Changed note text" : `Note paragraph ${index}`,
                  ),
                );
                const savedAgain =
                  savePath === "repack"
                    ? await repackDocx(reparsed, options)
                    : await repackDocxFromRaw(reparsed, await unzipDocx(saved), options);
                expect(await readPart(savedAgain, partPath)).toBe(savedPart);
                if (editBody) {
                  const savedBody = reparsed.package.document.content.at(0);
                  const savedText = (savedBody?.type === "paragraph" ? savedBody.content : [])
                    .flatMap((item) => (item.type === "run" ? item.content : []))
                    .map((item) => (item.type === "text" ? item.text : ""))
                    .join("");
                  expect(savedText).toBe("Changed body text");
                }
              },
            ),
            propertyConfig({
              numRuns: 20,
              examples: [
                [["71000001"], false, false, false, null],
                [["71000001", "71000002"], true, true, false, 2],
                [[], false, true, true, null],
              ],
            }),
          );
        },
        propertyTestTimeout(30_000),
      );
    }
  }
});
