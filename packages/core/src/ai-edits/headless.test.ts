/**
 * Headless `.docx` review round-trip: parse a fixture, apply body-block edits
 * with no editor view, serialise back, and re-parse to assert the edits landed.
 *
 * Covers `replaceInBlock` and `insertAfterBlock` in both `direct` and
 * `tracked-changes` modes, that tracked mode emits ins/del marks, and that
 * unedited parts survive byte-exact (selective save for the non-structural
 * case; the copied-through package parts for the structural full-repack case).
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import { readFileSync } from "node:fs";
import path from "node:path";

import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ensureParaIdsInState } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { schema, singletonManager } from "../prosemirror/schema";
import { FolioDocxReviewer, applyFolioAIEditsToBuffer } from "./headless";
import type { FolioAIBlock } from "./types";

const FIXTURE = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus/authored-empty-paragraph.docx",
);

const readFixture = (): ArrayBuffer => {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

/**
 * Corpus fixtures ship without `w14:paraId`s. The editor allocates them on load
 * and its first save is a full repack that writes them out; only subsequent
 * saves take the selective path. Reproduce that here so the reviewer under test
 * sees a paraId-bearing baseline and can exercise selective save.
 */
const makeParaIdBaseline = async (source: ArrayBuffer): Promise<ArrayBuffer> => {
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  const state = ensureParaIdsInState(
    EditorState.create({
      schema,
      doc: toProseDoc(document),
      plugins: singletonManager.getPlugins(),
    }),
  );
  return repackDocx({ ...updateDocumentContent(document, state.doc), originalBuffer: source });
};

const partText = async (buffer: ArrayBuffer, part: string): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(part);
  if (!file) {
    throw new Error(`missing part: ${part}`);
  }
  return file.async("text");
};

const partBytes = async (buffer: ArrayBuffer, part: string): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(part);
  if (!file) {
    throw new Error(`missing part: ${part}`);
  }
  return file.async("uint8array");
};

/** Extract the `<w:p …>…</w:p>` element that contains `needle`. */
const paragraphContaining = (documentXml: string, needle: string): string => {
  const at = documentXml.indexOf(needle);
  const start = documentXml.lastIndexOf("<w:p ", at);
  const startAlt = documentXml.lastIndexOf("<w:p>", at);
  const from = Math.max(start, startAlt);
  const end = documentXml.indexOf("</w:p>", at);
  if (from === -1 || end === -1) {
    throw new Error(`no <w:p> around ${JSON.stringify(needle)}`);
  }
  return documentXml.slice(from, end + "</w:p>".length);
};

const blockText = (block: FolioAIBlock): string => block.text;

const findBlock = (blocks: FolioAIBlock[], needle: string): FolioAIBlock => {
  const block = blocks.find((b) => b.text.includes(needle));
  if (!block) {
    throw new Error(`no block containing ${JSON.stringify(needle)}`);
  }
  return block;
};

describe("headless docx review round-trip", () => {
  test("replaceInBlock (direct) rewrites the text with no tracked marks", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");

    const result = reviewer.applyOperations(
      [{ id: "r1", type: "replaceInBlock", blockId: target.id, find: "Heading", replace: "Intro" }],
      { mode: "direct" },
    );
    expect(result.applied.map((a) => a.id)).toEqual(["r1"]);
    expect(result.skipped).toEqual([]);

    const out = await reviewer.toBuffer();
    const outXml = await partText(out, "word/document.xml");
    expect(outXml).toContain("Intro paragraph.");
    expect(outXml).not.toContain("Heading paragraph.");
    // Direct mode is an in-place edit: no redline marks introduced.
    expect(outXml).not.toContain("<w:ins ");
    expect(outXml).not.toContain("<w:del ");

    // Selective save leaves the untouched paragraph byte-for-byte identical.
    const baseXml = await partText(baseline, "word/document.xml");
    expect(outXml).toContain(paragraphContaining(baseXml, "Trailing paragraph."));
  });

  test("replaceInBlock (tracked-changes) redlines old vs new and preserves untouched parts", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");

    const result = reviewer.applyOperations([
      { id: "t1", type: "replaceInBlock", blockId: target.id, find: "Heading", replace: "Intro" },
    ]);
    expect(result.applied).toHaveLength(1);
    expect(result.applied[0]?.revisionIds?.length).toBe(2);

    const out = await reviewer.toBuffer();
    const outXml = await partText(out, "word/document.xml");
    // Tracked change: the deletion carries the old word, the insertion the new.
    expect(outXml).toContain("<w:ins ");
    expect(outXml).toContain("<w:del ");
    expect(outXml).toContain('w:author="AI Reviewer"');
    const editedPara = paragraphContaining(outXml, "Intro");
    expect(editedPara).toContain("<w:ins ");
    expect(editedPara).toContain("<w:del ");
    // Re-parse round-trips to a valid package.
    const reparsed = await parseDocx(out);
    expect(reparsed.package.document.content.length).toBeGreaterThan(0);

    // Byte-exact: the untouched paragraph and unrelated parts survive intact.
    const baseXml = await partText(baseline, "word/document.xml");
    expect(outXml).toContain(paragraphContaining(baseXml, "Trailing paragraph."));
    expect(await partBytes(out, "word/styles.xml")).toEqual(
      await partBytes(baseline, "word/styles.xml"),
    );
  });

  test("insertAfterBlock (direct) adds a sibling paragraph, no tracked marks", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const target = findBlock(reviewer.snapshot().blocks, "Heading");

    const result = reviewer.applyOperations(
      [
        {
          id: "i1",
          type: "insertAfterBlock",
          blockId: target.id,
          text: "Newly inserted clause.",
        },
      ],
      { mode: "direct" },
    );
    expect(result.applied.map((a) => a.id)).toEqual(["i1"]);

    const out = await reviewer.toBuffer();
    const outXml = await partText(out, "word/document.xml");
    expect(outXml).toContain("Newly inserted clause.");
    expect(outXml).not.toContain("<w:ins ");
    expect(outXml).not.toContain("<w:del ");

    const reparsed = await parseDocx(out);
    const blocks = reparsed.package.document.content.filter((b) => b.type === "paragraph");
    const joined = blocks.map((b) => (b.type === "paragraph" ? JSON.stringify(b) : "")).join(" ");
    expect(joined).toContain("Newly inserted clause.");
    expect(joined).toContain("Heading paragraph.");
    expect(joined).toContain("Trailing paragraph.");
  });

  test("insertAfterBlock (tracked-changes) marks the new paragraph as an insertion", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");

    const result = reviewer.applyOperations([
      {
        id: "i2",
        type: "insertAfterBlock",
        blockId: target.id,
        text: "Newly inserted clause.",
      },
    ]);
    expect(result.applied.map((a) => a.id)).toEqual(["i2"]);
    expect(result.applied[0]?.revisionId).toBeDefined();

    const out = await reviewer.toBuffer();
    const outXml = await partText(out, "word/document.xml");
    expect(outXml).toContain("Newly inserted clause.");
    expect(outXml).toContain("<w:ins ");
    // The inserted run sits inside a tracked insertion.
    expect(paragraphContaining(outXml, "Newly inserted clause.")).toContain("<w:ins ");

    // A structural edit falls back to full repack, which copies unrelated parts
    // through untouched.
    expect(await partBytes(out, "word/styles.xml")).toEqual(
      await partBytes(baseline, "word/styles.xml"),
    );
    const reparsed = await parseDocx(out);
    expect(reparsed.package.document.content.length).toBeGreaterThan(0);
  });

  test("applyFolioAIEditsToBuffer one-shot convenience applies against a fresh snapshot", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    // Derive the operation's blockId from a snapshot of the same buffer.
    const snapshot = (await FolioDocxReviewer.fromBuffer(baseline)).snapshot();
    const target = findBlock(snapshot.blocks, "Heading");

    const { buffer, applied, skipped } = await applyFolioAIEditsToBuffer(
      baseline,
      [{ id: "o1", type: "replaceInBlock", blockId: target.id, find: "Heading", replace: "Intro" }],
      { mode: "direct", snapshot },
    );
    expect(applied.map((a) => a.id)).toEqual(["o1"]);
    expect(skipped).toEqual([]);
    expect(blockText(target)).toContain("Heading");
    expect(await partText(buffer, "word/document.xml")).toContain("Intro paragraph.");
  });
});
