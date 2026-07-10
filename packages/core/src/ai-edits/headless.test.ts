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
import type { FolioReviewChange } from "./headless";
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

  test("applyDocumentOperations executes a versioned tracked-change batch", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");

    const result = reviewer.applyDocumentOperations({
      version: 1,
      operations: [
        {
          id: "v1",
          type: "replaceInBlock",
          blockId: target.id,
          find: "Heading",
          replace: "Intro",
        },
      ],
      mode: "tracked-changes",
    });

    expect(result.version).toBe(1);
    expect(result.status).toBe("committed");
    expect(result.applied.map(({ id }) => id)).toEqual(["v1"]);
    expect(result.skipped).toEqual([]);
    expect(await partText(await reviewer.toBuffer(), "word/document.xml")).toContain("<w:ins ");
  });

  test("atomic batches reject every operation without document or comment changes", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");
    const contentBefore = reviewer.getContentAsText();

    const result = reviewer.applyDocumentOperations({
      version: 1,
      atomic: true,
      mode: "direct",
      operations: [
        {
          id: "valid",
          type: "replaceInBlock",
          blockId: target.id,
          find: "Heading",
          replace: "Intro",
          comment: { text: "Review this change." },
        },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });

    expect(result).toEqual({
      version: 1,
      status: "rejected",
      applied: [],
      skipped: [
        { id: "valid", reason: "atomicBatchRejected" },
        { id: "missing", reason: "missingBlock" },
      ],
    });
    expect(reviewer.getContentAsText()).toBe(contentBefore);
    expect(reviewer.getComments()).toEqual([]);
  });

  test("atomic batches commit all operations after a successful preflight", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");

    const result = reviewer.applyDocumentOperations({
      version: 1,
      atomic: true,
      mode: "direct",
      operations: [
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: target.id,
          find: "Heading",
          replace: "Intro",
          comment: { text: "Review this change." },
        },
        {
          id: "insert",
          type: "insertAfterBlock",
          blockId: target.id,
          text: "New clause.",
        },
      ],
    });

    expect(result.status).toBe("committed");
    expect(result.applied.map(({ id }) => id).toSorted()).toEqual(["insert", "replace"]);
    expect(result.skipped).toEqual([]);
    expect(reviewer.getContentAsText()).toContain("Intro paragraph.");
    expect(reviewer.getContentAsText()).toContain("New clause.");
    expect(reviewer.getComments()).toHaveLength(1);
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

const insertionChange = (reviewer: FolioDocxReviewer): FolioReviewChange => {
  const change = reviewer.getChanges().find((c) => c.type === "insertion");
  if (!change) {
    throw new Error("expected an insertion change");
  }
  return change;
};

describe("headless docx review discovery + resolve", () => {
  test("getContentAsText labels every block with its stable id", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const blocks = reviewer.getContent();
    const heading = findBlock(blocks, "Heading");

    const text = reviewer.getContentAsText();
    expect(text).toContain(`[${heading.id}] `);
    expect(text).toContain("Heading paragraph.");
    for (const block of blocks) {
      expect(text).toContain(`[${block.id}]`);
    }
  });

  test("getChanges surfaces the deletion and insertion of a tracked replace", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");
    reviewer.applyOperations([
      { id: "t1", type: "replaceInBlock", blockId: target.id, find: "Heading", replace: "Intro" },
    ]);

    const changes = reviewer.getChanges();
    const insertion = changes.find((c) => c.type === "insertion");
    const deletion = changes.find((c) => c.type === "deletion");
    expect(insertion?.text).toContain("Intro");
    expect(deletion?.text).toContain("Heading");
    expect(insertion?.author).toBe("AI Reviewer");
    expect(insertion?.blockId).toBe(target.id);
    // The two sides carry distinct revision ids.
    expect(insertion?.id).not.toBe(deletion?.id);
    // Filtering narrows to one kind.
    const onlyInsertions = reviewer.getChanges({ type: "insertion" });
    expect(onlyInsertions.every((c) => c.type === "insertion")).toBe(true);
    expect(reviewer.getChanges({ author: "Nobody" })).toEqual([]);
  });

  test("acceptChange keeps the insertion as plain text; rejectChange drops it", async () => {
    const baseline = await makeParaIdBaseline(readFixture());

    const accepting = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI" });
    const acceptTarget = findBlock(accepting.snapshot().blocks, "Heading");
    accepting.applyOperations([
      {
        id: "a1",
        type: "replaceInBlock",
        blockId: acceptTarget.id,
        find: "Heading",
        replace: "Intro",
      },
    ]);
    expect(accepting.acceptChange(insertionChange(accepting))).toBe(true);
    const acceptedXml = await partText(await accepting.toBuffer(), "word/document.xml");
    expect(acceptedXml).toContain("Intro");
    expect(acceptedXml).not.toContain("<w:ins ");
    // The accepted change no longer shows up in discovery.
    expect(accepting.getChanges().some((c) => c.type === "insertion")).toBe(false);

    const rejecting = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI" });
    const rejectTarget = findBlock(rejecting.snapshot().blocks, "Heading");
    rejecting.applyOperations([
      {
        id: "r1",
        type: "replaceInBlock",
        blockId: rejectTarget.id,
        find: "Heading",
        replace: "Intro",
      },
    ]);
    expect(rejecting.rejectChange(insertionChange(rejecting))).toBe(true);
    const rejectedXml = await partText(await rejecting.toBuffer(), "word/document.xml");
    expect(rejectedXml).not.toContain("Intro");
    expect(rejectedXml).not.toContain("<w:ins ");
  });

  test("acceptAll and rejectAll resolve every tracked change", async () => {
    const baseline = await makeParaIdBaseline(readFixture());

    const accepting = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI" });
    const acceptBlocks = accepting.snapshot().blocks;
    accepting.applyOperations([
      {
        id: "m1",
        type: "replaceInBlock",
        blockId: findBlock(acceptBlocks, "Heading").id,
        find: "Heading",
        replace: "Intro",
      },
      {
        id: "m2",
        type: "replaceInBlock",
        blockId: findBlock(acceptBlocks, "Trailing").id,
        find: "Trailing",
        replace: "Closing",
      },
    ]);
    expect(accepting.getChanges().length).toBeGreaterThanOrEqual(4);
    expect(accepting.acceptAll()).toBeGreaterThanOrEqual(4);
    const acceptedXml = await partText(await accepting.toBuffer(), "word/document.xml");
    expect(acceptedXml).toContain("Intro paragraph.");
    expect(acceptedXml).toContain("Closing paragraph.");
    expect(acceptedXml).not.toContain("<w:ins ");
    expect(acceptedXml).not.toContain("<w:del ");
    expect(accepting.getChanges()).toEqual([]);

    const rejecting = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI" });
    const rejectBlocks = rejecting.snapshot().blocks;
    rejecting.applyOperations([
      {
        id: "m3",
        type: "replaceInBlock",
        blockId: findBlock(rejectBlocks, "Heading").id,
        find: "Heading",
        replace: "Intro",
      },
      {
        id: "m4",
        type: "replaceInBlock",
        blockId: findBlock(rejectBlocks, "Trailing").id,
        find: "Trailing",
        replace: "Closing",
      },
    ]);
    expect(rejecting.rejectAll()).toBeGreaterThanOrEqual(4);
    const rejectedXml = await partText(await rejecting.toBuffer(), "word/document.xml");
    expect(rejectedXml).toContain("Heading paragraph.");
    expect(rejectedXml).toContain("Trailing paragraph.");
    expect(rejectedXml).not.toContain("Intro");
    expect(rejectedXml).not.toContain("Closing");
    expect(rejectedXml).not.toContain("<w:ins ");
    expect(rejectedXml).not.toContain("<w:del ");
  });

  test("getComments returns authored comments with their anchor", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");
    reviewer.applyOperations([
      {
        id: "c1",
        type: "commentOnBlock",
        blockId: target.id,
        comment: { text: "Clarify this clause." },
      },
    ]);

    const comments = reviewer.getComments();
    expect(comments).toHaveLength(1);
    expect(comments[0]?.text).toBe("Clarify this clause.");
    expect(comments[0]?.author).toBe("AI Reviewer");
    expect(comments[0]?.blockId).toBe(target.id);
    expect(comments[0]?.anchoredText).toContain("Heading paragraph.");
    expect(reviewer.getComments({ author: "Nobody" })).toEqual([]);
  });

  test("replyTo threads a reply under a comment and round-trips it", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");
    reviewer.applyOperations([
      { id: "c1", type: "commentOnBlock", blockId: target.id, comment: { text: "Parent note." } },
    ]);
    const parent = reviewer.getComments()[0];
    expect(parent).toBeDefined();
    if (!parent) {
      return;
    }

    const reply = reviewer.replyTo(parent, { text: "Threaded response.", author: "Second" });
    expect(reply).not.toBeNull();

    // The thread now carries the reply.
    const thread = reviewer.getComments()[0];
    expect(thread?.replies).toHaveLength(1);
    expect(thread?.replies[0]?.text).toBe("Threaded response.");

    // Round-trips as a real reply: linked in commentsExtended.xml and anchored.
    const saved = await reviewer.toBuffer();
    const reparsed = await parseDocx(saved, { preloadFonts: false });
    const savedComments = reparsed.package.document.comments ?? [];
    const savedReply = savedComments.find((c) => c.parentId !== undefined);
    expect(savedReply?.parentId).toBe(parent.id);

    const extended = await partText(saved, "word/commentsExtended.xml");
    expect(extended).toContain("w15:paraIdParent");
    const documentXml = await partText(saved, "word/document.xml");
    expect(documentXml).toContain(`<w:commentRangeStart w:id="${savedReply?.id}"/>`);
    expect(documentXml).toContain(`<w:commentReference w:id="${savedReply?.id}"/>`);

    // Adding comments.xml AND commentsExtended.xml in one repack must wire both
    // parts' content-type overrides and relationships (they edit the same two
    // packaging files, so a concurrent write would drop one).
    const contentTypes = await partText(saved, "[Content_Types].xml");
    expect(contentTypes).toContain("/word/comments.xml");
    expect(contentTypes).toContain("/word/commentsExtended.xml");
    const rels = await partText(saved, "word/_rels/document.xml.rels");
    expect(rels.toLowerCase()).toContain('target="comments.xml"');
    expect(rels.toLowerCase()).toContain("commentsextended.xml");

    // A reply to an unknown comment is refused.
    expect(reviewer.replyTo(9999, { text: "orphan" })).toBeNull();
  });

  test("resolveComment marks a thread resolved, round-trips, and can reopen it", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");
    reviewer.applyOperations([
      { id: "c1", type: "commentOnBlock", blockId: target.id, comment: { text: "Please check." } },
    ]);
    const comment = reviewer.getComments()[0];
    expect(comment).toBeDefined();
    if (!comment) {
      return;
    }
    expect(comment.done).toBe(false);

    // Resolving an unknown id is refused.
    expect(reviewer.resolveComment("9999")).toBe(false);

    expect(reviewer.resolveComment(String(comment.id))).toBe(true);
    expect(reviewer.getComments()[0]?.done).toBe(true);

    // Survives a save + re-parse via commentsExtended.xml's `w15:done`.
    const saved = await reviewer.toBuffer();
    const extended = await partText(saved, "word/commentsExtended.xml");
    expect(extended).toContain('w15:done="1"');
    const reparsed = await parseDocx(saved, { preloadFonts: false });
    const savedComment = (reparsed.package.document.comments ?? []).find(
      (c) => c.id === comment.id,
    );
    expect(savedComment?.done).toBe(true);

    const reopenedReviewer = await FolioDocxReviewer.fromBuffer(saved, { author: "AI Reviewer" });
    expect(reopenedReviewer.getComments()[0]?.done).toBe(true);
    expect(reopenedReviewer.resolveComment(String(comment.id), { resolved: false })).toBe(true);
    expect(reopenedReviewer.getComments()[0]?.done).toBe(false);

    const reopenedSaved = await reopenedReviewer.toBuffer();
    const reopenedExtended = await partText(reopenedSaved, "word/commentsExtended.xml");
    expect(reopenedExtended).toContain('w15:done="0"');
    const reopenedReparsed = await parseDocx(reopenedSaved, { preloadFonts: false });
    const reopenedSavedComment = (reopenedReparsed.package.document.comments ?? []).find(
      (c) => c.id === comment.id,
    );
    expect(reopenedSavedComment?.done).toBe(false);
  });

  test("ops from one parse's snapshot resolve on a separate parse of a paraId-less doc", async () => {
    // The raw corpus fixture ships without `w14:paraId`s. Two independent
    // parses must mint identical block ids so ops built against the first
    // snapshot resolve against the second rather than skipping as missingBlock.
    const first = await FolioDocxReviewer.fromBuffer(readFixture());
    const snapshot = first.snapshot();
    const target = findBlock(snapshot.blocks, "Heading");

    const second = await FolioDocxReviewer.fromBuffer(readFixture());
    expect(second.getContent().map((b) => b.id)).toEqual(first.getContent().map((b) => b.id));

    const result = second.applyOperations(
      [{ id: "d1", type: "replaceInBlock", blockId: target.id, find: "Heading", replace: "Intro" }],
      { mode: "direct" },
    );
    expect(result.skipped).toEqual([]);
    expect(result.applied.map((a) => a.id)).toEqual(["d1"]);

    const outXml = await partText(await second.toBuffer(), "word/document.xml");
    expect(outXml).toContain("Intro paragraph.");
  });
});

describe("headless docx review annotated read surface", () => {
  test("getContentAsText({ annotated }) inlines redline + comment tags; default stays clean", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const blocks = reviewer.snapshot().blocks;
    reviewer.applyOperations([
      {
        id: "t1",
        type: "replaceInBlock",
        blockId: findBlock(blocks, "Heading").id,
        find: "Heading",
        replace: "Intro",
      },
      {
        id: "c1",
        type: "commentOnBlock",
        blockId: findBlock(blocks, "Trailing").id,
        comment: { text: "Check this clause." },
      },
    ]);

    const annotated = reviewer.getContentAsText({ annotated: true });
    // Tracked replace surfaces as a deletion of the old word and an insertion
    // of the new one, both attributed to the reviewer's author.
    expect(annotated).toContain('<ins author="AI Reviewer">Intro</ins>');
    expect(annotated).toContain('<del author="AI Reviewer">Heading</del>');
    // The comment anchor wraps its block text with its allocated id.
    expect(annotated).toMatch(/<comment id="\d+">[^<]*Trailing paragraph\.[^<]*<\/comment>/u);

    // Default (clean) output flattens tracked changes and carries no tags.
    const clean = reviewer.getContentAsText();
    expect(clean).not.toContain("<ins ");
    expect(clean).not.toContain("<del ");
    expect(clean).not.toContain("<comment ");
    expect(clean).toContain("Intro paragraph.");
  });
});

const NOTES_FIXTURE = path.join(
  import.meta.dir,
  "../../../../tests/visual/fixtures/docx-editor-demo.docx",
);

/**
 * The demo fixture already carries a header + footer with real text, but its
 * `footnotes.xml` holds only separator notes. Splice in one normal footnote so
 * the notes read surface has a header, a footer, and a footnote to report.
 */
const readNotesFixture = async (): Promise<ArrayBuffer> => {
  const bytes = readFileSync(NOTES_FIXTURE);
  const src = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const zip = await JSZip.loadAsync(src);
  const footnotesFile = zip.file("word/footnotes.xml");
  if (!footnotesFile) {
    throw new Error("fixture missing word/footnotes.xml");
  }
  const footnotesXml = await footnotesFile.async("text");
  const injected = footnotesXml.replace(
    "</w:footnotes>",
    '<w:footnote w:id="2"><w:p><w:r><w:t>Injected footnote body text.</w:t></w:r></w:p></w:footnote></w:footnotes>',
  );
  zip.file("word/footnotes.xml", injected);
  return zip.generateAsync({ type: "arraybuffer" });
};

describe("headless docx review notes read surface", () => {
  test("getNotesAsText surfaces header, footer, and footnote text with labels", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await readNotesFixture());
    const notes = reviewer.getNotesAsText();

    expect(notes).toContain(
      "[header default] Docx Editor: Project Charter & Contributor Agreement",
    );
    expect(notes).toContain("[footer default]");
    expect(notes).toContain("[footnote #2] Injected footnote body text.");
    // Body content is not folded into the notes surface.
    expect(notes).not.toContain("[endnote");
  });

  test("getNotesAsText is empty for a document without headers/footers or notes", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    expect(reviewer.getNotesAsText()).toBe("");
  });
});
