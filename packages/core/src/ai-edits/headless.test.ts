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

import { buildTextBoxTableDocument, findTextBoxShape } from "../__tests__/textBoxTableDocument";
import { parseDocx } from "../docx/parser";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createEmptyDocx, repackDocx } from "../docx/rezip";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ensureParaIdsInState } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { schema, singletonManager } from "../prosemirror/schema";
import type { HeaderFooter } from "../types/document";
import {
  FolioDocumentStoryNotFoundError,
  FolioDocxReviewer,
  UnsupportedFolioReviewedViewError,
  applyFolioAIEditsToBuffer,
} from "./headless";
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

const HEADER_RELATIONSHIP_ID = "rId_header_story";
const FOOTER_RELATIONSHIP_ID = "rId_footer_story";

const textStory = (type: "header" | "footer", text: string, paraId: string): HeaderFooter => ({
  type,
  hdrFtrType: "default",
  content: [
    {
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

const makeHeaderFooterBaseline = async (): Promise<ArrayBuffer> => {
  const source = await createEmptyDocx();
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  document.package.headers = new Map([
    [HEADER_RELATIONSHIP_ID, textStory("header", "Header text", "A1000001")],
  ]);
  document.package.footers = new Map([
    [FOOTER_RELATIONSHIP_ID, textStory("footer", "Footer text", "A1000002")],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_RELATIONSHIP_ID }],
    footerReferences: [{ type: "default", rId: FOOTER_RELATIONSHIP_ID }],
  };
  const materialized = await repackDocx(document, { updateModifiedDate: false });
  const { docx } = await ensureParaIds(materialized);
  return new Uint8Array(docx).buffer;
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
  test("edits a table cell inside shape text and preserves its container", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const target = findBlock(reviewer.snapshot().blocks, "Cell value");

    const result = reviewer.applyOperations(
      [
        {
          id: "replace-shape-table-cell",
          type: "replaceInBlock",
          blockId: target.id,
          find: "Cell value",
          replace: "Updated value",
        },
      ],
      { mode: "direct" },
    );

    expect(result.applied.map(({ id }) => id)).toEqual(["replace-shape-table-cell"]);
    expect(result.skipped).toEqual([]);

    const saved = await reviewer.toBuffer();
    const documentXml = await partText(saved, "word/document.xml");
    expect(documentXml).toContain("<wps:txbx><w:txbxContent>");
    expect(documentXml).toContain("<w:tbl>");
    expect(documentXml).toContain("Updated value");
    expect(documentXml).not.toContain("Cell value");

    const reparsed = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    const shape = findTextBoxShape(reparsed);
    expect(shape.size).toEqual({ width: 1_828_800, height: 914_400 });
    expect(shape.textBody?.margins).toEqual({
      top: 45_720,
      bottom: 45_720,
      left: 91_440,
      right: 91_440,
    });
    expect(shape.textBody?.content.map(({ type }) => type)).toEqual([
      "paragraph",
      "table",
      "paragraph",
    ]);

    const table = shape.textBody?.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a nested table");
    }
    const cellParagraph = table.rows.at(0)?.cells.at(0)?.content.at(0);
    expect(cellParagraph?.type).toBe("paragraph");
    expect(
      await FolioDocxReviewer.fromBuffer(saved).then((reopened) => reopened.getContentAsText()),
    ).toContain("Updated value");
  });

  test("inserts, persists, and undoes a row inside shape text", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const target = findBlock(reviewer.snapshot().blocks, "Cell value");

    const rejected = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      atomic: true,
      operations: [
        {
          id: "insert-row",
          type: "insertTableRow",
          blockId: target.id,
          cellTexts: ["Added value"],
        },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });

    expect(rejected.status).toBe("rejected");
    expect(rejected.applied).toEqual([]);
    expect(reviewer.getContentAsText()).not.toContain("Added value");

    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "insert-row",
          type: "insertTableRow",
          blockId: target.id,
          cellTexts: ["Added value"],
        },
      ],
    });

    expect(result.status).toBe("committed");
    expect(result.applied).toEqual([{ id: "insert-row" }]);
    expect(result.receipts).toEqual([
      {
        operationId: "insert-row",
        operationIndex: 0,
        affected: [
          {
            type: "insertion",
            story: "main",
            anchorBlockId: target.id,
            position: "after",
            content: "tableRow",
          },
        ],
      },
    ]);

    const saved = await reviewer.toBuffer();
    const reparsed = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    const table = findTextBoxShape(reparsed).textBody?.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(table.rows).toHaveLength(2);
    expect(table.rows.at(0)?.cells.at(0)?.content.at(0)?.type).toBe("paragraph");
    expect(table.rows.at(1)?.cells.at(0)?.content.at(0)?.type).toBe("paragraph");
    expect((await FolioDocxReviewer.fromBuffer(saved)).getContentAsText()).toContain("Added value");

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.getContentAsText()).not.toContain("Added value");
    const restored = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const restoredTable = findTextBoxShape(restored).textBody?.content.at(1);
    expect(restoredTable?.type).toBe("table");
    if (restoredTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(restoredTable.rows).toHaveLength(1);
  });

  test("tracks, persists, accepts, and rejects a row insertion", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Cell value");

    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "insert-row",
          type: "insertTableRow",
          blockId: target.id,
          cellTexts: ["Added value"],
        },
      ],
    });

    expect(result.status).toBe("committed");
    expect(result.applied).toHaveLength(1);
    expect(typeof result.applied.at(0)?.revisionId).toBe("number");
    const pendingChange = reviewer.getChanges().find(({ type }) => type === "rowInserted");
    if (!pendingChange) {
      throw new Error("expected a pending row insertion");
    }
    expect(pendingChange.author).toBe("Reviewer");

    const pending = await reviewer.toBuffer();
    const pendingXml = await partText(pending, "word/document.xml");
    expect(pendingXml).toContain("<w:ins ");
    expect(pendingXml).toContain("Added value");

    const reopened = await FolioDocxReviewer.fromBuffer(pending);
    const reopenedChange = reopened.getChanges().find(({ type }) => type === "rowInserted");
    if (!reopenedChange) {
      throw new Error("expected the row insertion after reopen");
    }

    expect(reopened.acceptChange(reopenedChange)).toBe(true);
    const accepted = await reopened.toBuffer();
    expect(await partText(accepted, "word/document.xml")).not.toContain("<w:ins ");
    expect(reopened.getContentAsText()).toContain("Added value");
    const acceptedTable = findTextBoxShape(
      await parseDocx(accepted, { detectVariables: false, preloadFonts: false }),
    ).textBody?.content.at(1);
    expect(acceptedTable?.type).toBe("table");
    if (acceptedTable?.type !== "table") {
      throw new Error("expected an accepted table");
    }
    expect(acceptedTable.rows).toHaveLength(2);

    const rejecting = await FolioDocxReviewer.fromBuffer(pending);
    const rejectChange = rejecting.getChanges().find(({ type }) => type === "rowInserted");
    if (!rejectChange) {
      throw new Error("expected the row insertion before rejection");
    }
    expect(rejecting.rejectChange(rejectChange)).toBe(true);
    const rejected = await rejecting.toBuffer();
    expect(await partText(rejected, "word/document.xml")).not.toContain("<w:ins ");
    expect(rejecting.getContentAsText()).not.toContain("Added value");
    const rejectedTable = findTextBoxShape(
      await parseDocx(rejected, { detectVariables: false, preloadFonts: false }),
    ).textBody?.content.at(1);
    expect(rejectedTable?.type).toBe("table");
    if (rejectedTable?.type !== "table") {
      throw new Error("expected a rejected table");
    }
    expect(rejectedTable.rows).toHaveLength(1);
  });

  test("deletes, persists, and undoes a row inside shape text", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const target = findBlock(reviewer.snapshot().blocks, "Cell value");

    const rejected = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      atomic: true,
      operations: [
        {
          id: "delete-row",
          type: "deleteTableRow",
          blockId: target.id,
        },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });

    expect(rejected.status).toBe("rejected");
    expect(rejected.applied).toEqual([]);
    expect(reviewer.getContentAsText()).toContain("Cell value");

    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "delete-row",
          type: "deleteTableRow",
          blockId: target.id,
        },
      ],
    });

    expect(result.status).toBe("committed");
    expect(result.applied).toEqual([{ id: "delete-row" }]);
    expect(result.receipts).toEqual([
      {
        operationId: "delete-row",
        operationIndex: 0,
        affected: [
          {
            type: "tableRow",
            story: "main",
            anchorBlockId: target.id,
            effect: "deleted",
          },
        ],
      },
    ]);

    const saved = await reviewer.toBuffer();
    const reparsed = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    expect(findTextBoxShape(reparsed).textBody?.content.map(({ type }) => type)).toEqual([
      "paragraph",
      "paragraph",
    ]);
    expect((await FolioDocxReviewer.fromBuffer(saved)).getContentAsText()).not.toContain(
      "Cell value",
    );

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.getContentAsText()).toContain("Cell value");
    const restored = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const restoredTable = findTextBoxShape(restored).textBody?.content.at(1);
    expect(restoredTable?.type).toBe("table");
    if (restoredTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(restoredTable.rows).toHaveLength(1);
  });

  test("inserts, persists, and undoes a column inside shape text", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const target = findBlock(reviewer.snapshot().blocks, "Cell value");

    const rejected = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      atomic: true,
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: target.id,
          cellTexts: ["Added value"],
        },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });

    expect(rejected.status).toBe("rejected");
    expect(rejected.applied).toEqual([]);
    expect(reviewer.getContentAsText()).not.toContain("Added value");

    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: target.id,
          cellTexts: ["Added value"],
        },
      ],
    });

    expect(result.status).toBe("committed");
    expect(result.applied).toEqual([{ id: "insert-column" }]);
    expect(result.receipts).toEqual([
      {
        operationId: "insert-column",
        operationIndex: 0,
        affected: [
          {
            type: "insertion",
            story: "main",
            anchorBlockId: target.id,
            position: "after",
            content: "tableColumn",
          },
        ],
      },
    ]);

    const saved = await reviewer.toBuffer();
    const reparsed = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    const table = findTextBoxShape(reparsed).textBody?.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(table.rows).toHaveLength(1);
    expect(table.rows.at(0)?.cells).toHaveLength(2);
    expect((await FolioDocxReviewer.fromBuffer(saved)).getContentAsText()).toContain("Added value");

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.getContentAsText()).not.toContain("Added value");
    const restored = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const restoredTable = findTextBoxShape(restored).textBody?.content.at(1);
    expect(restoredTable?.type).toBe("table");
    if (restoredTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(restoredTable.rows.at(0)?.cells).toHaveLength(1);
  });

  test("merges, persists, and undoes cells inside shape text", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const initialTarget = findBlock(reviewer.snapshot().blocks, "Cell value");
    const insertion = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId: initialTarget.id,
          cellTexts: ["Added value"],
        },
      ],
    });
    expect(insertion.status).toBe("committed");

    const snapshot = reviewer.snapshot();
    const startTarget = findBlock(snapshot.blocks, "Cell value");
    const endTarget = findBlock(snapshot.blocks, "Added value");
    const rejected = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      atomic: true,
      operations: [
        {
          id: "merge-cells",
          type: "mergeTableCells",
          blockId: startTarget.id,
          endBlockId: endTarget.id,
        },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });

    expect(rejected.status).toBe("rejected");
    expect(rejected.applied).toEqual([]);
    const afterRejected = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const rejectedTable = findTextBoxShape(afterRejected).textBody?.content.at(1);
    expect(rejectedTable?.type).toBe("table");
    if (rejectedTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(rejectedTable.rows.at(0)?.cells).toHaveLength(2);

    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "merge-cells",
          type: "mergeTableCells",
          blockId: startTarget.id,
          endBlockId: endTarget.id,
        },
      ],
    });

    expect(result.status).toBe("committed");
    expect(result.applied).toEqual([{ id: "merge-cells" }]);
    expect(result.receipts).toEqual([
      {
        operationId: "merge-cells",
        operationIndex: 0,
        affected: [
          {
            type: "tableCells",
            story: "main",
            anchorBlockId: startTarget.id,
            endAnchorBlockId: endTarget.id,
            effect: "merged",
          },
        ],
      },
    ]);

    const saved = await reviewer.toBuffer();
    const reparsed = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    const table = findTextBoxShape(reparsed).textBody?.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(table.rows.at(0)?.cells).toHaveLength(1);
    expect((await FolioDocxReviewer.fromBuffer(saved)).getContentAsText()).toContain("Cell value");
    expect((await FolioDocxReviewer.fromBuffer(saved)).getContentAsText()).toContain("Added value");

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    const restored = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const restoredTable = findTextBoxShape(restored).textBody?.content.at(1);
    expect(restoredTable?.type).toBe("table");
    if (restoredTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(restoredTable.rows.at(0)?.cells).toHaveLength(2);
  });

  test("persists a vertical cell merge with continuation geometry", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const initialTarget = findBlock(reviewer.snapshot().blocks, "Cell value");
    const insertion = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "insert-row",
          type: "insertTableRow",
          blockId: initialTarget.id,
          cellTexts: ["Added value"],
        },
      ],
    });
    expect(insertion.status).toBe("committed");

    const snapshot = reviewer.snapshot();
    const startTarget = findBlock(snapshot.blocks, "Cell value");
    const endTarget = findBlock(snapshot.blocks, "Added value");
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "merge-cells",
          type: "mergeTableCells",
          blockId: startTarget.id,
          endBlockId: endTarget.id,
        },
      ],
    });

    expect(result.status).toBe("committed");
    const saved = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const table = findTextBoxShape(saved).textBody?.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(table.rows).toHaveLength(2);
    expect(table.rows.at(0)?.cells.at(0)?.formatting?.vMerge).toBe("restart");
    expect(table.rows.at(1)?.cells.at(0)?.formatting?.vMerge).toBe("continue");

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle).status).toBe("undone");
    const restored = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const restoredTable = findTextBoxShape(restored).textBody?.content.at(1);
    expect(restoredTable?.type).toBe("table");
    if (restoredTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(restoredTable.rows.at(0)?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
    expect(restoredTable.rows.at(1)?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
  });

  test("splits, persists, and undoes a merged cell inside shape text", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const initialTarget = findBlock(reviewer.snapshot().blocks, "Cell value");
    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "insert-column",
            type: "insertTableColumn",
            blockId: initialTarget.id,
            cellTexts: ["Added value"],
          },
        ],
      }).status,
    ).toBe("committed");
    const mergeSnapshot = reviewer.snapshot();
    const mergeStart = findBlock(mergeSnapshot.blocks, "Cell value");
    const mergeEnd = findBlock(mergeSnapshot.blocks, "Added value");
    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "merge",
            type: "mergeTableCells",
            blockId: mergeStart.id,
            endBlockId: mergeEnd.id,
          },
        ],
      }).status,
    ).toBe("committed");

    const splitTarget = findBlock(reviewer.snapshot().blocks, "Cell value");
    const rejected = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      atomic: true,
      operations: [
        { id: "split-cell", type: "splitTableCell", blockId: splitTarget.id },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });
    expect(rejected.status).toBe("rejected");
    expect(rejected.applied).toEqual([]);

    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [{ id: "split-cell", type: "splitTableCell", blockId: splitTarget.id }],
    });

    expect(result.status).toBe("committed");
    expect(result.applied).toEqual([{ id: "split-cell" }]);
    expect(result.receipts).toEqual([
      {
        operationId: "split-cell",
        operationIndex: 0,
        affected: [
          {
            type: "tableCell",
            story: "main",
            anchorBlockId: splitTarget.id,
            effect: "split",
          },
        ],
      },
    ]);
    const saved = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const table = findTextBoxShape(saved).textBody?.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(table.rows.at(0)?.cells).toHaveLength(2);
    expect(table.rows.at(0)?.cells.at(0)?.formatting?.gridSpan).toBeUndefined();
    expect(table.rows.at(0)?.cells.at(1)?.formatting?.gridSpan).toBeUndefined();

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle).status).toBe("undone");
    const restored = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const restoredTable = findTextBoxShape(restored).textBody?.content.at(1);
    expect(restoredTable?.type).toBe("table");
    if (restoredTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(restoredTable.rows.at(0)?.cells).toHaveLength(1);
    expect(restoredTable.rows.at(0)?.cells.at(0)?.formatting?.gridSpan).toBe(2);
  });

  test("clears vertical merge metadata when splitting a cell", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const initialTarget = findBlock(reviewer.snapshot().blocks, "Cell value");
    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "insert-row",
            type: "insertTableRow",
            blockId: initialTarget.id,
            cellTexts: ["Added value"],
          },
        ],
      }).status,
    ).toBe("committed");
    const mergeSnapshot = reviewer.snapshot();
    const mergeStart = findBlock(mergeSnapshot.blocks, "Cell value");
    const mergeEnd = findBlock(mergeSnapshot.blocks, "Added value");
    expect(
      reviewer.applyDocumentOperations({
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "merge",
            type: "mergeTableCells",
            blockId: mergeStart.id,
            endBlockId: mergeEnd.id,
          },
        ],
      }).status,
    ).toBe("committed");

    const splitTarget = findBlock(reviewer.snapshot().blocks, "Cell value");
    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [{ id: "split", type: "splitTableCell", blockId: splitTarget.id }],
    });
    expect(result.status).toBe("committed");

    const saved = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const table = findTextBoxShape(saved).textBody?.content.at(1);
    expect(table?.type).toBe("table");
    if (table?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(table.rows).toHaveLength(2);
    expect(table.rows.at(0)?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
    expect(table.rows.at(1)?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
  });

  test("deletes, persists, and undoes a column inside shape text", async () => {
    const baseline = await buildTextBoxTableDocument();
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const target = findBlock(reviewer.snapshot().blocks, "Cell value");

    const rejected = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      atomic: true,
      operations: [
        {
          id: "delete-column",
          type: "deleteTableColumn",
          blockId: target.id,
        },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });

    expect(rejected.status).toBe("rejected");
    expect(rejected.applied).toEqual([]);
    expect(reviewer.getContentAsText()).toContain("Cell value");

    const result = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "delete-column",
          type: "deleteTableColumn",
          blockId: target.id,
        },
      ],
    });

    expect(result.status).toBe("committed");
    expect(result.applied).toEqual([{ id: "delete-column" }]);
    expect(result.receipts).toEqual([
      {
        operationId: "delete-column",
        operationIndex: 0,
        affected: [
          {
            type: "tableColumn",
            story: "main",
            anchorBlockId: target.id,
            effect: "deleted",
          },
        ],
      },
    ]);

    const saved = await reviewer.toBuffer();
    const reparsed = await parseDocx(saved, { detectVariables: false, preloadFonts: false });
    expect(findTextBoxShape(reparsed).textBody?.content.map(({ type }) => type)).toEqual([
      "paragraph",
      "paragraph",
    ]);
    expect((await FolioDocxReviewer.fromBuffer(saved)).getContentAsText()).not.toContain(
      "Cell value",
    );

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.getContentAsText()).toContain("Cell value");
    const restored = await parseDocx(await reviewer.toBuffer(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const restoredTable = findTextBoxShape(restored).textBody?.content.at(1);
    expect(restoredTable?.type).toBe("table");
    if (restoredTable?.type !== "table") {
      throw new Error("expected a nested table");
    }
    expect(restoredTable.rows.at(0)?.cells).toHaveLength(1);
  });

  test("edits a header through story-scoped document operations", async () => {
    const baseline = await makeHeaderFooterBaseline();
    const story = { type: "header", relationshipId: HEADER_RELATIONSHIP_ID } as const;
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "Reviewer" });
    const snapshot = reviewer.snapshotStory(story);
    if (!snapshot) {
      throw new Error("expected the header story");
    }
    const target = findBlock(snapshot.blocks, "Header text");

    const result = reviewer.applyDocumentOperationsToStory({
      story,
      snapshot,
      batch: {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "header-replace",
            type: "replaceInBlock",
            blockId: target.id,
            find: "Header text",
            replace: "Updated header",
          },
        ],
      },
    });

    expect(result.status).toBe("committed");
    expect(result.applied.map(({ id }) => id)).toEqual(["header-replace"]);
    expect(result.receipts.at(0)?.affected.at(0)).toEqual({
      type: "block",
      story,
      blockId: target.id,
      effect: "updated",
    });
    expect(reviewer.readStory(story)?.text).toBe("Updated header");
    expect(reviewer.getNotesAsText()).toContain("[header default] Updated header");

    const saved = await reviewer.toBuffer();
    const headerXml = await partText(saved, "word/header1.xml");
    expect(headerXml).toContain("Updated header");
    expect(headerXml).not.toContain("Header text");
    expect(headerXml).not.toContain("<w:ins ");
    expect(headerXml).not.toContain("<w:del ");
    expect(await partBytes(saved, "word/document.xml")).toEqual(
      await partBytes(baseline, "word/document.xml"),
    );
    expect(await partBytes(saved, "word/footer1.xml")).toEqual(
      await partBytes(baseline, "word/footer1.xml"),
    );

    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    expect(reopened.readStory(story)?.text).toBe("Updated header");
  });

  test("tracks and undoes a footer operation batch", async () => {
    const baseline = await makeHeaderFooterBaseline();
    const story = { type: "footer", relationshipId: FOOTER_RELATIONSHIP_ID } as const;
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "Reviewer" });
    const snapshot = reviewer.snapshotStory(story);
    if (!snapshot) {
      throw new Error("expected the footer story");
    }
    const target = findBlock(snapshot.blocks, "Footer text");

    const result = reviewer.applyDocumentOperationsToStory({
      story,
      snapshot,
      batch: {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "footer-replace",
            type: "replaceInBlock",
            blockId: target.id,
            find: "Footer text",
            replace: "Updated footer",
          },
        ],
      },
    });

    expect(result.status).toBe("committed");
    const saved = await reviewer.toBuffer();
    const footerXml = await partText(saved, "word/footer1.xml");
    expect(footerXml).toContain("Updated");
    expect(footerXml).toContain("footer");
    expect(footerXml).toContain("<w:ins ");
    expect(footerXml).toContain("<w:del ");
    expect(footerXml).toContain('w:author="Reviewer"');

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.readStory(story)?.text).toBe("Footer text");
    expect(await partBytes(await reviewer.toBuffer(), "word/footer1.xml")).toEqual(
      await partBytes(baseline, "word/footer1.xml"),
    );
  });

  test("rejects a missing editable story before applying operations", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await makeHeaderFooterBaseline());
    const story = { type: "header", relationshipId: "missing" } as const;
    expect(reviewer.snapshotStory(story)).toBeNull();
    expect(() =>
      reviewer.applyDocumentOperationsToStory({
        story,
        batch: { version: 1, operations: [] },
      }),
    ).toThrow(FolioDocumentStoryNotFoundError);
  });

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
    const contentBefore = reviewer.getContentAsText();

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
    expect(result.receipts).toEqual([
      {
        operationId: "v1",
        operationIndex: 0,
        affected: [{ type: "block", story: "main", blockId: target.id, effect: "updated" }],
      },
    ]);
    expect(result.undoHandle).toEqual({
      type: "documentOperationUndo",
      id: expect.any(String),
    });
    expect(await partText(await reviewer.toBuffer(), "word/document.xml")).toContain("<w:ins ");
    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.getContentAsText()).toBe(contentBefore);
    expect(reviewer.getChanges()).toEqual([]);
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
      issues: [
        {
          operationId: "valid",
          operationIndex: 0,
          path: "$.operations[0]",
          code: "atomicBatchRejected",
          retryable: true,
          recovery: "inspectBatch",
        },
        {
          operationId: "missing",
          operationIndex: 1,
          path: "$.operations[1]",
          code: "missingBlock",
          retryable: true,
          recovery: "refreshDocument",
        },
      ],
      receipts: [],
      undoHandle: null,
    });
    expect(reviewer.getContentAsText()).toBe(contentBefore);
    expect(reviewer.getComments()).toEqual([]);
  });

  test("atomic batches commit all operations after a successful preflight", async () => {
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
    expect(result.receipts).toEqual([
      {
        operationId: "replace",
        operationIndex: 0,
        affected: [
          { type: "block", story: "main", blockId: target.id, effect: "updated" },
          { type: "comment", commentId: expect.any(Number) },
        ],
      },
      {
        operationId: "insert",
        operationIndex: 1,
        affected: [
          {
            type: "insertion",
            story: "main",
            anchorBlockId: target.id,
            position: "after",
            content: "block",
          },
        ],
      },
    ]);
    expect(reviewer.getContentAsText()).toContain("Intro paragraph.");
    expect(reviewer.getContentAsText()).toContain("New clause.");
    expect(reviewer.getComments()).toHaveLength(1);
    expect(result.undoHandle).toEqual({
      type: "documentOperationUndo",
      id: expect.any(String),
    });
    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.getContentAsText()).toBe(contentBefore);
    expect(reviewer.getComments()).toEqual([]);

    const reparsed = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(reparsed.getContentAsText()).toBe(contentBefore);
    expect(reparsed.getComments()).toEqual([]);
  });

  test("undo handles reject unknown, out-of-order, and changed document state", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const skipped = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [{ id: "missing", type: "deleteBlock", blockId: "missing" }],
    });
    expect(skipped.status).toBe("committed");
    expect(skipped.undoHandle).toBeNull();
    const heading = findBlock(reviewer.snapshot().blocks, "Heading");
    const first = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "first",
          type: "replaceInBlock",
          blockId: heading.id,
          find: "Heading",
          replace: "Intro",
        },
      ],
    });
    const trailing = findBlock(reviewer.snapshot().blocks, "Trailing");
    const second = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "second",
          type: "replaceInBlock",
          blockId: trailing.id,
          find: "Trailing",
          replace: "Closing",
        },
      ],
    });
    if (!first.undoHandle || !second.undoHandle) {
      throw new Error("expected undo handles");
    }

    expect(reviewer.undoDocumentOperations(first.undoHandle)).toEqual({
      status: "rejected",
      undoHandle: first.undoHandle,
      reason: "notLatest",
    });
    expect(reviewer.undoDocumentOperations(second.undoHandle)).toEqual({
      status: "undone",
      undoHandle: second.undoHandle,
    });
    expect(reviewer.undoDocumentOperations(first.undoHandle)).toEqual({
      status: "undone",
      undoHandle: first.undoHandle,
    });
    const unknownHandle = { type: "documentOperationUndo", id: "unknown" } as const;
    expect(reviewer.undoDocumentOperations(unknownHandle)).toEqual({
      status: "rejected",
      undoHandle: unknownHandle,
      reason: "unknownHandle",
    });

    const changedReviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const changedTarget = findBlock(changedReviewer.snapshot().blocks, "Heading");
    const changed = changedReviewer.applyDocumentOperations({
      version: 1,
      operations: [
        {
          id: "changed",
          type: "replaceInBlock",
          blockId: changedTarget.id,
          find: "Heading",
          replace: "Intro",
        },
      ],
    });
    if (!changed.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(changedReviewer.undoDocumentOperations(first.undoHandle)).toEqual({
      status: "rejected",
      undoHandle: first.undoHandle,
      reason: "unknownHandle",
    });

    const commentReviewer = await FolioDocxReviewer.fromBuffer(baseline);
    const commentTarget = findBlock(commentReviewer.snapshot().blocks, "Heading");
    const commented = commentReviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations: [
        {
          id: "commented",
          type: "commentOnBlock",
          blockId: commentTarget.id,
          comment: { text: "Review this." },
        },
      ],
    });
    const parentComment = commentReviewer.getComments().at(0);
    if (!commented.undoHandle || !parentComment) {
      throw new Error("expected an undo handle and comment");
    }
    commentReviewer.replyTo(parentComment.id, { text: "Keep this reply." });
    expect(commentReviewer.undoDocumentOperations(commented.undoHandle)).toEqual({
      status: "rejected",
      undoHandle: commented.undoHandle,
      reason: "documentChanged",
    });
    expect(commentReviewer.getComments().at(0)?.replies.at(0)?.text).toBe("Keep this reply.");

    changedReviewer.acceptAll();
    expect(changedReviewer.undoDocumentOperations(changed.undoHandle)).toEqual({
      status: "rejected",
      undoHandle: changed.undoHandle,
      reason: "documentChanged",
    });
  });

  test("dry runs predict best-effort results without document or comment changes", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");
    const contentBefore = reviewer.getContentAsText();
    const operations = [
      {
        id: "valid",
        type: "replaceInBlock" as const,
        blockId: target.id,
        find: "Heading",
        replace: "Intro",
        comment: { text: "Review this change." },
      },
      { id: "missing", type: "deleteBlock" as const, blockId: "para-missing" },
    ];

    const preview = reviewer.applyDocumentOperations({
      version: 1,
      dryRun: true,
      mode: "direct",
      operations,
    });

    expect(preview).toEqual({
      version: 1,
      status: "previewed",
      applied: [{ id: "valid" }],
      skipped: [{ id: "missing", reason: "missingBlock" }],
      issues: [
        {
          operationId: "missing",
          operationIndex: 1,
          path: "$.operations[1]",
          code: "missingBlock",
          retryable: true,
          recovery: "refreshDocument",
        },
      ],
      receipts: [
        {
          operationId: "valid",
          operationIndex: 0,
          affected: [{ type: "block", story: "main", blockId: target.id, effect: "updated" }],
        },
      ],
      undoHandle: null,
    });
    expect(reviewer.getContentAsText()).toBe(contentBefore);
    expect(reviewer.getComments()).toEqual([]);

    const committed = reviewer.applyDocumentOperations({
      version: 1,
      mode: "direct",
      operations,
    });
    expect(committed.status).toBe("committed");
    expect(committed.applied.map(({ id }) => id)).toEqual(["valid"]);
    expect(committed.skipped).toEqual([{ id: "missing", reason: "missingBlock" }]);
    expect(reviewer.getContentAsText()).toContain("Intro paragraph.");
    expect(reviewer.getComments()).toHaveLength(1);
  });

  test("dry runs predict atomic rejection without exposing generated ids", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "AI Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");

    const preview = reviewer.applyDocumentOperations({
      version: 1,
      atomic: true,
      dryRun: true,
      operations: [
        {
          id: "valid",
          type: "replaceInBlock",
          blockId: target.id,
          find: "Heading",
          replace: "Intro",
        },
        { id: "missing", type: "deleteBlock", blockId: "para-missing" },
      ],
    });

    expect(preview).toEqual({
      version: 1,
      status: "previewed",
      applied: [],
      skipped: [
        { id: "valid", reason: "atomicBatchRejected" },
        { id: "missing", reason: "missingBlock" },
      ],
      issues: [
        {
          operationId: "valid",
          operationIndex: 0,
          path: "$.operations[0]",
          code: "atomicBatchRejected",
          retryable: true,
          recovery: "inspectBatch",
        },
        {
          operationId: "missing",
          operationIndex: 1,
          path: "$.operations[1]",
          code: "missingBlock",
          retryable: true,
          recovery: "refreshDocument",
        },
      ],
      receipts: [],
      undoHandle: null,
    });
    expect(reviewer.getContentAsText()).toContain("Heading paragraph.");
    expect(reviewer.getChanges()).toEqual([]);
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
  test("reads original, current-markup, and final views without mutating the reviewer", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "Reviewer" });
    const target = findBlock(reviewer.snapshot().blocks, "Heading");
    reviewer.applyOperations([
      {
        id: "reviewed-view-replace",
        type: "replaceInBlock",
        blockId: target.id,
        find: "Heading",
        replace: "Intro",
      },
    ]);
    const changesBefore = reviewer.getChanges();

    const original = reviewer.readReviewedStory({ view: "original" });
    const currentMarkup = reviewer.readReviewedStory({ view: "current-markup" });
    const final = reviewer.readReviewedStory({ view: "final" });

    expect(original?.text).toContain("Heading paragraph.");
    expect(original?.text).not.toContain("Intro paragraph.");
    expect(original?.changes).toEqual([]);
    expect(currentMarkup?.text).toContain('<ins author="Reviewer">Intro</ins>');
    expect(currentMarkup?.text).toContain('<del author="Reviewer">Heading</del>');
    expect(currentMarkup?.changes).toEqual(changesBefore);
    expect(final?.text).toContain("Intro paragraph.");
    expect(final?.text).not.toContain("Heading paragraph.");
    expect(final?.changes).toEqual([]);
    expect(reviewer.getChanges()).toEqual(changesBefore);
    expect(reviewer.getContentAsText()).toContain("Intro paragraph.");
  });

  test("rejects an unsupported reviewed view at the public boundary", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await makeParaIdBaseline(readFixture()));
    expect(() =>
      Reflect.apply(reviewer.readReviewedStory, reviewer, [{ view: "unsupported" }]),
    ).toThrow(UnsupportedFolioReviewedViewError);
  });

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
    '<w:footnote w:id="2"><w:p w14:paraId="B2000001"><w:r><w:t>Injected footnote body text.</w:t></w:r></w:p></w:footnote></w:footnotes>',
  );
  zip.file("word/footnotes.xml", injected);
  zip.file(
    "word/endnotes.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:endnote w:id="3"><w:p w14:paraId="B2000002"><w:r><w:t>Injected endnote body text.</w:t></w:r></w:p></w:endnote></w:endnotes>',
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

describe("headless docx review notes read surface", () => {
  test("edits a footnote through story-scoped document operations", async () => {
    const baseline = await readNotesFixture();
    const story = { type: "footnote", noteId: 2 } as const;
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "Reviewer" });
    const snapshot = reviewer.snapshotStory(story);
    if (!snapshot) {
      throw new Error("expected the footnote story");
    }
    const target = findBlock(snapshot.blocks, "Injected footnote body text.");

    const result = reviewer.applyDocumentOperationsToStory({
      story,
      snapshot,
      batch: {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "footnote-replace",
            type: "replaceInBlock",
            blockId: target.id,
            find: "Injected footnote",
            replace: "Updated footnote",
          },
        ],
      },
    });

    expect(result.status).toBe("committed");
    expect(result.receipts.at(0)?.affected.at(0)).toEqual({
      type: "block",
      story,
      blockId: target.id,
      effect: "updated",
    });
    expect(reviewer.readStory(story)?.text).toBe("Updated footnote body text.");
    expect(reviewer.getNotesAsText()).toContain("[footnote #2] Updated footnote body text.");

    const saved = await reviewer.toBuffer();
    expect(await partText(saved, "word/footnotes.xml")).toContain("Updated footnote");
    expect(await partBytes(saved, "word/document.xml")).toEqual(
      await partBytes(baseline, "word/document.xml"),
    );
    expect(await partBytes(saved, "word/endnotes.xml")).toEqual(
      await partBytes(baseline, "word/endnotes.xml"),
    );

    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    expect(reopened.readStory(story)?.text).toBe("Updated footnote body text.");
  });

  test("tracks and undoes an endnote operation batch", async () => {
    const baseline = await readNotesFixture();
    const story = { type: "endnote", noteId: 3 } as const;
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: "Reviewer" });
    const snapshot = reviewer.snapshotStory(story);
    if (!snapshot) {
      throw new Error("expected the endnote story");
    }
    const target = findBlock(snapshot.blocks, "Injected endnote body text.");

    const result = reviewer.applyDocumentOperationsToStory({
      story,
      snapshot,
      batch: {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "endnote-replace",
            type: "replaceInBlock",
            blockId: target.id,
            find: "Injected endnote",
            replace: "Updated endnote",
          },
        ],
      },
    });

    expect(result.status).toBe("committed");
    expect(reviewer.readReviewedStory({ story, view: "original" })?.text).toContain(
      "Injected endnote body text.",
    );
    expect(reviewer.readReviewedStory({ story, view: "current-markup" })?.text).toContain(
      '<ins author="Reviewer">Updated</ins>',
    );
    expect(reviewer.readReviewedStory({ story, view: "final" })?.text).toContain(
      "Updated endnote body text.",
    );
    const saved = await reviewer.toBuffer();
    const endnotesXml = await partText(saved, "word/endnotes.xml");
    expect(endnotesXml).toContain("Updated");
    expect(endnotesXml).toContain("endnote body text.");
    expect(endnotesXml).toContain("<w:ins ");
    expect(endnotesXml).toContain("<w:del ");
    expect(endnotesXml).toContain('w:author="Reviewer"');

    if (!result.undoHandle) {
      throw new Error("expected an undo handle");
    }
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reviewer.readStory(story)?.text).toBe("Injected endnote body text.");
    expect(await partBytes(await reviewer.toBuffer(), "word/endnotes.xml")).toEqual(
      await partBytes(baseline, "word/endnotes.xml"),
    );
  });

  test("rejects a missing note story before applying operations", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await readNotesFixture());
    const story = { type: "footnote", noteId: 999_999 } as const;
    expect(reviewer.snapshotStory(story)).toBeNull();
    expect(reviewer.readReviewedStory({ story, view: "final" })).toBeNull();
    expect(() =>
      reviewer.applyDocumentOperationsToStory({
        story,
        batch: { version: 1, operations: [] },
      }),
    ).toThrow(FolioDocumentStoryNotFoundError);
  });

  test("listStories discovers typed handles that readStory resolves", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await readNotesFixture());
    const stories = reviewer.listStories();
    expect(stories.at(0)?.handle).toEqual({ type: "main" });
    const footnote = stories.find(({ handle }) => handle.type === "footnote");
    expect(footnote?.text).toBe("Injected footnote body text.");
    expect(footnote ? reviewer.readStory(footnote.handle) : null).toEqual(footnote);
    const endnote = stories.find(({ handle }) => handle.type === "endnote");
    expect(endnote?.text).toBe("Injected endnote body text.");
    expect(endnote ? reviewer.readStory(endnote.handle) : null).toEqual(endnote);
    expect(reviewer.readStory({ type: "footnote", noteId: 999_999 })).toBeNull();
  });

  test("getNotesAsText surfaces header, footer, and footnote text with labels", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await readNotesFixture());
    const notes = reviewer.getNotesAsText();

    expect(notes).toContain(
      "[header default] Docx Editor: Project Charter & Contributor Agreement",
    );
    expect(notes).toContain("[footer default]");
    expect(notes).toContain("[footnote #2] Injected footnote body text.");
    expect(notes).toContain("[endnote #3] Injected endnote body text.");
    // Body content is not folded into the notes surface.
    expect(notes).not.toContain("Project Goals");
  });

  test("getNotesAsText is empty for a document without headers/footers or notes", async () => {
    const baseline = await makeParaIdBaseline(readFixture());
    const reviewer = await FolioDocxReviewer.fromBuffer(baseline);
    expect(reviewer.getNotesAsText()).toBe("");
  });
});
