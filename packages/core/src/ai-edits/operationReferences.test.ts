/**
 * Operation references the document does not define, and the save-time check
 * every applied batch passes (issue #1103).
 *
 * A `numbering.numId` naming no `w:num` used to apply, commit and hand back a
 * receipt, leaving a model only `toBuffer()` refused. It is now skipped before
 * anything is applied, like a block id naming no block. Behind it, every batch
 * result is checked with the validator the save runs, so an operation whose
 * result would not save is refused rather than committed.
 */

import { describe, expect, test } from "bun:test";
import { EditorState, type Transaction } from "prosemirror-state";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  applyFolioDocumentOperations,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_MODES,
  isFolioDocumentOperationModeSupported,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import { schema } from "../prosemirror/schema";
import type { NumberingDefinitions } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { findIntroducedModelErrors } from "./result-validation";
import { createFolioAIEditSnapshot } from "./snapshot";

const ANCHOR = "1.2. Second clause.";
const TAIL = "Tail.";
const MARKDOWN = `# Agreement\n\n## 1. Subject\n\n1.1. First clause.\n\n${ANCHOR}\n\n${TAIL}`;

/** One decimal instance (901) no paragraph uses. */
const UNUSED_901: NumberingDefinitions = {
  abstractNums: [
    {
      abstractNumId: 901,
      multiLevelType: "multilevel",
      levels: [
        {
          ilvl: 0,
          start: 1,
          numFmt: "decimal",
          lvlText: "%1.",
          suffix: "tab",
          pPr: { indentLeft: 720, indentFirstLine: -360, hangingIndent: true },
        },
      ],
    },
  ],
  nums: [{ numId: 901, abstractNumId: 901 }],
};

const PACKAGES = {
  "without a numbering part": undefined,
  "with other instances only": UNUSED_901,
} as const;

const build = async (numbering?: NumberingDefinitions): Promise<Uint8Array> => {
  const model = fromMarkdown(MARKDOWN);
  if (numbering) {
    model.package.numbering = numbering;
  }
  return (await ensureParaIds(await createDocx(model))).docx;
};

const outline = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map((block) => `${block.displayLabel ?? "·"} ${block.text}`);

type NumberingOperationFactory = (
  ids: { anchor: string; tail: string },
  numbering: { numId: number; level: number },
) => FolioDocumentOperation;

/** Every operation that carries a numbering reference, by where it carries it. */
const NUMBERING_OPERATIONS: Record<string, NumberingOperationFactory> = {
  insertAfterBlock: ({ anchor }, numbering) => ({
    id: "op",
    type: "insertAfterBlock",
    blockId: anchor,
    text: "1.3. Inserted clause.",
    numbering,
  }),
  insertBeforeBlock: ({ anchor }, numbering) => ({
    id: "op",
    type: "insertBeforeBlock",
    blockId: anchor,
    text: "1.1a. Inserted clause.",
    numbering,
  }),
  setBlockParagraphProperties: ({ anchor }, numbering) => ({
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: anchor,
    properties: { numbering },
  }),
  "splitBlock (first paragraph)": ({ anchor }, numbering) => ({
    id: "op",
    type: "splitBlock",
    blockId: anchor,
    offset: 5,
    firstParagraphProperties: { numbering },
  }),
  "splitBlock (second paragraph)": ({ anchor }, numbering) => ({
    id: "op",
    type: "splitBlock",
    blockId: anchor,
    offset: 5,
    secondParagraphProperties: { numbering },
  }),
  mergeBlockWithNext: ({ anchor }, numbering) => ({
    id: "op",
    type: "mergeBlockWithNext",
    blockId: anchor,
    mergedParagraphProperties: { numbering },
  }),
};

const cases = Object.entries(PACKAGES).flatMap(([packageName, numbering]) =>
  Object.entries(NUMBERING_OPERATIONS).flatMap(([operationName, factory]) =>
    FOLIO_DOCUMENT_OPERATION_MODES.filter((mode) =>
      isFolioDocumentOperationModeSupported(
        factory({ anchor: "", tail: "" }, { numId: 1, level: 0 }).type,
        mode,
      ),
    ).map((mode) => ({ packageName, numbering, operationName, factory, mode })),
  ),
);

const openReviewer = async (numbering?: NumberingDefinitions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await build(numbering), { author: "AI" });
  const blocks = reviewer.getContent();
  const anchor = blocks.find((block) => block.text === ANCHOR)?.id;
  const tail = blocks.find((block) => block.text === TAIL)?.id;
  if (anchor === undefined || tail === undefined) {
    throw new Error("fixture must expose its anchor and tail blocks");
  }
  return { reviewer, ids: { anchor, tail } };
};

const applyOne = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
  mode: FolioDocumentOperationMode,
) =>
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [operation],
  });

describe("an undefined numbering instance is skipped before anything is applied", () => {
  test.each(cases)(
    "$operationName, package $packageName, $mode mode",
    async ({ numbering, factory, mode }) => {
      const { reviewer, ids } = await openReviewer(numbering);
      const before = outline(reviewer);

      const result = applyOne(reviewer, factory(ids, { numId: 1, level: 0 }), mode);

      expect(result.status).toBe("committed");
      expect(result.applied).toEqual([]);
      expect(result.receipts).toEqual([]);
      expect(result.undoHandle).toBeNull();
      expect(result.skipped).toEqual([
        { id: "op", reason: "missingNumbering", message: expect.stringContaining("numId 1") },
      ]);
      expect(result.issues).toEqual([
        {
          operationId: "op",
          operationIndex: 0,
          path: "$.operations[0]",
          code: "missingNumbering",
          retryable: true,
          recovery: "refreshDocument",
          message: expect.stringContaining("names no numbering instance"),
        },
      ]);
      expect(outline(reviewer)).toEqual(before);
      const saved = await reviewer.toBuffer();
      expect(outline(await FolioDocxReviewer.fromBuffer(saved, { author: "AI" }))).toEqual(before);
    },
  );

  test("the skip names the instances the document does define", async () => {
    const { reviewer, ids } = await openReviewer(UNUSED_901);
    const result = applyOne(
      reviewer,
      NUMBERING_OPERATIONS["insertAfterBlock"]!(ids, { numId: 1, level: 0 }),
      "tracked-changes",
    );
    expect(result.skipped[0]?.message).toBe(
      "numbering.numId 1 names no numbering instance in this document (it defines 901).",
    );
  });

  test.each(FOLIO_DOCUMENT_OPERATION_MODES)(
    "control: the instance the package defines applies and saves (%s)",
    async (mode) => {
      const { reviewer, ids } = await openReviewer(UNUSED_901);
      const result = applyOne(
        reviewer,
        NUMBERING_OPERATIONS["insertAfterBlock"]!(ids, { numId: 901, level: 0 }),
        mode,
      );
      expect(result.applied).toEqual([expect.objectContaining({ id: "op" })]);
      expect(result.issues).toEqual([]);
      expect(outline(reviewer)).toContain("1. 1.3. Inserted clause.");
      await reviewer.toBuffer();
    },
  );

  test("an atomic batch holding one undefined instance commits nothing", async () => {
    const { reviewer, ids } = await openReviewer(UNUSED_901);
    const before = outline(reviewer);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      atomic: true,
      operations: [
        { id: "good", type: "insertAfterBlock", blockId: ids.tail, text: "Fine." },
        { ...NUMBERING_OPERATIONS["insertAfterBlock"]!(ids, { numId: 7, level: 0 }), id: "bad" },
      ],
    });
    expect(result.status).toBe("rejected");
    expect(result.skipped.map(({ id, reason }) => [id, reason])).toEqual([
      ["good", "atomicBatchRejected"],
      ["bad", "missingNumbering"],
    ]);
    expect(outline(reviewer)).toEqual(before);
    await reviewer.toBuffer();
  });

  test("a level the instance does not define is left to the document, as before", async () => {
    const { reviewer, ids } = await openReviewer(UNUSED_901);
    const result = applyOne(
      reviewer,
      NUMBERING_OPERATIONS["insertAfterBlock"]!(ids, { numId: 901, level: 5 }),
      "tracked-changes",
    );
    expect(result.applied).toHaveLength(1);
    await reviewer.toBuffer();
  });
});

describe("a batch result the save validator refuses is not committed", () => {
  // A whitespace-only author makes every tracked change the batch writes one
  // no save accepts — nothing an operation's own fields can be checked for up
  // front, which is what the result check is for.
  const openWithBlankAuthor = async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await build(UNUSED_901), {
      author: "  ",
    });
    const tail = reviewer.getContent().find((block) => block.text === TAIL)?.id;
    if (tail === undefined) {
      throw new Error("fixture must expose its tail block");
    }
    return { reviewer, tail };
  };

  test("the operation is skipped with invalidResult and the validator's reason", async () => {
    const { reviewer, tail } = await openWithBlankAuthor();
    const before = outline(reviewer);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "insert", type: "insertAfterBlock", blockId: tail, text: "Added." }],
    });
    expect(result.status).toBe("committed");
    expect(result.applied).toEqual([]);
    expect(result.undoHandle).toBeNull();
    expect(result.issues).toEqual([
      {
        operationId: "insert",
        operationIndex: 0,
        path: "$.operations[0]",
        code: "invalidResult",
        retryable: false,
        recovery: "refreshDocument",
        message: expect.stringContaining("Tracked change author is empty."),
      },
    ]);
    expect(outline(reviewer)).toEqual(before);
    await reviewer.toBuffer();
  });

  test("in a best-effort batch only the offending operation is refused", async () => {
    const { reviewer, tail } = await openWithBlankAuthor();
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        { id: "direct", type: "insertAfterBlock", blockId: tail, text: "Direct." },
        { id: "noted", type: "commentOnBlock", blockId: tail, comment: { text: "Why?" } },
      ],
    });
    // Direct insertion writes no revision, so no author; the comment's author
    // is the blank one, which no save accepts.
    expect(result.applied.map(({ id }) => id)).toEqual(["direct"]);
    expect(result.skipped).toEqual([
      {
        id: "noted",
        reason: "invalidResult",
        message: expect.stringContaining("Comment author cannot be whitespace-only."),
      },
    ]);
    expect(outline(reviewer)).toContain(`· Direct.`);
    await reviewer.toBuffer();
  });

  test("an atomic batch is rejected whole", async () => {
    const { reviewer, tail } = await openWithBlankAuthor();
    const before = outline(reviewer);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      atomic: true,
      operations: [
        { id: "a", type: "insertAfterBlock", blockId: tail, text: "One." },
        { id: "b", type: "insertBeforeBlock", blockId: tail, text: "Two." },
      ],
    });
    expect(result.status).toBe("rejected");
    expect(result.skipped.map(({ reason }) => reason)).toEqual(["invalidResult", "invalidResult"]);
    expect(outline(reviewer)).toEqual(before);
  });

  test("a dry run reports the refusal without applying anything", async () => {
    const { reviewer, tail } = await openWithBlankAuthor();
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      dryRun: true,
      operations: [{ id: "a", type: "insertAfterBlock", blockId: tail, text: "One." }],
    });
    expect(result.status).toBe("previewed");
    expect(result.skipped.map(({ reason }) => reason)).toEqual(["invalidResult"]);
  });

  test("a live view never receives the refused transaction", () => {
    // The seam a live editor passes: its dispatch is what reaches the screen.
    const model = fromMarkdown(`Alpha.\n\n${TAIL}`);
    let state = EditorState.create({
      schema,
      doc: toProseDoc(model),
      plugins: [createDocumentNumberingPlugin(model.package.numbering)],
    });
    const dispatched: Transaction[] = [];
    const view = {
      get state() {
        return state;
      },
      dispatch: (transaction: Transaction) => {
        dispatched.push(transaction);
        state = state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const tail = snapshot.blocks.find((block) => block.text === TAIL)?.id ?? "";
    const refused = applyFolioDocumentOperations({
      view,
      snapshot,
      author: " ",
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: "a", type: "insertAfterBlock", blockId: tail, text: "One." }],
      },
    });
    expect(refused.skipped.map(({ reason }) => reason)).toEqual(["invalidResult"]);
    expect(dispatched).toEqual([]);

    const accepted = applyFolioDocumentOperations({
      view,
      snapshot,
      author: "Reviewer",
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [{ id: "a", type: "insertAfterBlock", blockId: tail, text: "One." }],
      },
    });
    expect(accepted.applied.map(({ id }) => id)).toEqual(["a"]);
    expect(dispatched).toHaveLength(1);
  });
});

describe("findIntroducedModelErrors", () => {
  const numbered = (numId: number) => {
    const model = fromMarkdown("Alpha.\n\nBeta.\n\nGamma.");
    model.package.numbering = UNUSED_901;
    const doc = toProseDoc(model);
    const state = EditorState.create({ schema, doc });
    const setNumbering = (index: number, id: number) => (tr: Transaction) => {
      let position = 0;
      for (let child = 0; child < index; child += 1) {
        position += tr.doc.child(child).nodeSize;
      }
      return tr.setNodeAttribute(position, "numPr", { kind: "reference", numId: id, ilvl: 0 });
    };
    return { state, setNumbering, numId };
  };

  test("reports a dangling reference the batch wrote, at the story's block index", () => {
    const { state, setNumbering } = numbered(5);
    const after = setNumbering(2, 5)(state.tr).doc;
    expect(
      findIntroducedModelErrors(state.doc, after, { numbering: UNUSED_901, createdCommentIds: [] }),
    ).toEqual([
      {
        path: "package.document.content[2].formatting.numPr.numId",
        message: "Numbering definition 5 is missing.",
        severity: "error",
      },
    ]);
  });

  test("ignores what the window already reported before the batch", () => {
    const { state, setNumbering } = numbered(5);
    const broken = setNumbering(1, 5)(state.tr).doc;
    const after = setNumbering(
      1,
      5,
    )(EditorState.create({ schema, doc: broken }).tr.insertText("x", 2)).doc;
    expect(
      findIntroducedModelErrors(broken, after, { numbering: UNUSED_901, createdCommentIds: [] }),
    ).toEqual([]);
  });

  test("reports an additional error with the same message as an existing error", () => {
    const { state, setNumbering } = numbered(5);
    const broken = setNumbering(1, 5)(state.tr).doc;
    const after = setNumbering(
      2,
      5,
    )(EditorState.create({ schema, doc: broken }).tr.insertText("x", 2)).doc;
    expect(
      findIntroducedModelErrors(broken, after, { numbering: UNUSED_901, createdCommentIds: [] }),
    ).toEqual([
      {
        path: "package.document.content[2].formatting.numPr.numId",
        message: "Numbering definition 5 is missing.",
        severity: "error",
      },
    ]);
  });

  test("defined instances and untouched stories report nothing", () => {
    const { state, setNumbering } = numbered(901);
    const after = setNumbering(0, 901)(state.tr).doc;
    const context = { numbering: UNUSED_901, createdCommentIds: [] };
    expect(findIntroducedModelErrors(state.doc, after, context)).toEqual([]);
    expect(findIntroducedModelErrors(state.doc, state.doc, context)).toEqual([]);
  });

  test("without numbering knowledge a reference is not held against the batch", () => {
    const { state, setNumbering } = numbered(5);
    const after = setNumbering(2, 5)(state.tr).doc;
    expect(
      findIntroducedModelErrors(state.doc, after, { numbering: undefined, createdCommentIds: [] }),
    ).toEqual([]);
  });
});
