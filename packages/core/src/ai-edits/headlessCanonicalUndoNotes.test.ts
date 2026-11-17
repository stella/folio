/**
 * Undo of a committed public operation on the canonical reviewer gives back
 * the exact prior document, notes the operation removed with their
 * references included.
 *
 * The replayed document numbers a paragraph whose insertion wraps a deleted
 * footnote reference and inline math; deleting that paragraph directly
 * removes the note with it, and undo restores both. A property draws where a
 * paragraph's footnote reference stands (plain, inserted, deleted, deleted
 * inside an insertion), whether math stands beside it and whether the
 * paragraphs are numbered, applies a public operation to it directly or
 * tracked, and undoes it: every read and the document are the ones before.
 */

import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import * as documentOps from "@stll/docx-core/ops";

import {
  bodyBlockIds,
  operationFor,
} from "../../../../test/generators/reviewerOperationDraws";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import type { BlockContent, Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const numbered = { numPr: { kind: "reference", numId: 1, ilvl: 0 } } as const;

const BODY = [
  { type: "paragraph", content: [], paraId: "00000001" },
  {
    type: "paragraph",
    content: [
      {
        type: "insertion",
        info: { id: 1, author: "A" },
        content: [
          {
            type: "deletion",
            info: { id: 5000, author: "B", date: "2026-01-02T03:04:05Z" },
            content: [
              { type: "run", content: [{ type: "footnoteRef", id: 2 }] },
            ],
          },
          { type: "mathEquation", display: "inline", ommlXml: "<m:oMath/>" },
        ],
      },
    ],
    formatting: numbered,
    paraId: "00000002",
  },
  {
    type: "paragraph",
    content: [],
    formatting: numbered,
    paraId: "00000003",
  },
] as unknown as BlockContent[];

const PAGE = { pageWidth: 12240, pageHeight: 15840 };

/** The generated document the edits property drew (seed 1462115757). */
const DOCUMENT = {
  package: {
    document: {
      content: BODY,
      sections: [{ properties: PAGE, content: BODY }],
      finalSectionProperties: PAGE,
      comments: [
        {
          id: 7,
          author: "A",
          content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
        },
      ],
    },
    numbering: {
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
        },
      ],
      nums: [{ numId: 1, abstractNumId: 1 }],
    },
    footnotes: [
      {
        type: "footnote",
        id: 2,
        content: [{ type: "paragraph", paraId: "7FFFFFF1", content: [] }],
      },
    ],
    settings: { defaultTabStop: 720 },
    properties: { title: "synthetic" },
  },
  warnings: ["kept"],
} as unknown as Document;

const packaged = async (document: Document): Promise<ArrayBuffer> => {
  const prepared = await prepareCanonicalDocxInput(
    await createDocx(structuredClone(document)),
  );
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Canonical Reviewer",
    session: "canonical",
  });

/** Every read an undo must give back, and the document itself. */
const reads = (reviewer: FolioDocxReviewer) => ({
  content: reviewer.getContent(),
  text: reviewer.getContentAsText({ annotated: true }),
  notes: reviewer.getNotesAsText(),
  stories: reviewer.listStories(),
  changes: reviewer.getChanges(),
  comments: reviewer.getComments(),
  document: reviewer.toDocument(),
});

const applyAndUndo = async (
  operation: FolioDocumentOperation,
  document: Document = DOCUMENT,
  mode: "direct" | "tracked-changes" = "direct",
) => {
  const reviewer = await open(await packaged(document));
  const before = reads(reviewer);
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [operation],
      mode,
    },
    { undefinedReferences: "refuse" },
  );
  return { reviewer, before, result };
};

describe("undo of a direct deletion of a paragraph holding a note reference", () => {
  test("gives back the paragraph, the note and every read", async () => {
    const probe = await open(await packaged(DOCUMENT));
    const operation = operationFor(
      probe,
      { kind: "deleteBlock", block: 1_073_741_824, word: 0, text: "revised" },
      { among: bodyBlockIds(probe.toDocument().package.document.content) },
    );
    if (operation === null) throw new Error("the draw names no block");
    const { reviewer, before, result } = await applyAndUndo(operation);
    expect(result.skipped).toEqual([]);
    if (result.undoHandle === null) throw new Error("no undo handle");
    expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
      status: "undone",
      undoHandle: result.undoHandle,
    });
    expect(reads(reviewer)).toEqual(before);
  }, 60_000);
});

type Placement = "plain" | "inserted" | "deleted" | "deletedInsideInsertion";

type Draw = {
  placement: Placement;
  math: boolean;
  numbered: boolean;
  mode: "direct" | "tracked-changes";
  operation:
    | "deleteBlock"
    | "replaceBlock"
    | "replaceInBlock"
    | "insertAfterBlock"
    | "insertBeforeBlock"
    | "mergeBlockWithNext"
    | "mergeIntoIt"
    | "setBlockParagraphProperties";
};

const INSERTED = { id: 11, author: "A", date: "2026-01-02T03:04:05Z" };
const DELETED = { id: 12, author: "B", date: "2026-01-02T03:04:05Z" };

const textRun = (text: string) => ({
  type: "run",
  content: [{ type: "text", text }],
});

const referenceIn = (placement: Placement) => {
  const reference = { type: "run", content: [{ type: "footnoteRef", id: 2 }] };
  switch (placement) {
    case "plain":
      return reference;
    case "inserted":
      return { type: "insertion", info: INSERTED, content: [reference] };
    case "deleted":
      return { type: "deletion", info: DELETED, content: [reference] };
    case "deletedInsideInsertion":
      return {
        type: "insertion",
        info: INSERTED,
        content: [{ type: "deletion", info: DELETED, content: [reference] }],
      };
    default: {
      const unreachable: never = placement;
      throw new Error(`Unknown placement ${String(unreachable)}`);
    }
  }
};

/** Three paragraphs; the middle one holds the note's reference. */
const documentOf = (draw: Draw): Document => {
  const formatting = draw.numbered ? { formatting: numbered } : {};
  const body = [
    {
      type: "paragraph",
      content: [textRun("Opening words")],
      paraId: "00000001",
      ...formatting,
    },
    {
      type: "paragraph",
      content: [
        textRun("Noted words"),
        referenceIn(draw.placement),
        ...(draw.math
          ? [{ type: "mathEquation", display: "inline", ommlXml: "<m:oMath/>" }]
          : []),
      ],
      paraId: "00000002",
      ...formatting,
    },
    {
      type: "paragraph",
      content: [textRun("Closing words")],
      paraId: "00000003",
      ...formatting,
    },
  ] as unknown as BlockContent[];
  const base = DOCUMENT.package;
  return {
    ...DOCUMENT,
    package: {
      ...base,
      document: {
        ...base.document,
        content: body,
        sections: [{ properties: PAGE, content: body }],
      },
      footnotes: [
        {
          type: "footnote",
          id: 2,
          content: [
            {
              type: "paragraph",
              paraId: "7FFFFFF1",
              content: [textRun("The note")],
            },
          ],
        },
      ],
    },
  } as unknown as Document;
};

const operationOf = (
  draw: Draw,
  ids: { opening: string; noted: string },
): FolioDocumentOperation => {
  const blockId = ids.noted;
  switch (draw.operation) {
    case "deleteBlock":
      return { id: "op", type: "deleteBlock", blockId };
    case "replaceBlock":
      return { id: "op", type: "replaceBlock", blockId, text: "Other words" };
    case "replaceInBlock":
      return {
        id: "op",
        type: "replaceInBlock",
        blockId,
        find: "Noted",
        replace: "Marked",
      };
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return { id: "op", type: draw.operation, blockId, text: "Inserted" };
    case "mergeBlockWithNext":
      return { id: "op", type: "mergeBlockWithNext", blockId };
    case "mergeIntoIt":
      return { id: "op", type: "mergeBlockWithNext", blockId: ids.opening };
    case "setBlockParagraphProperties":
      return {
        id: "op",
        type: "setBlockParagraphProperties",
        blockId,
        properties: { alignment: "center" },
      };
    default: {
      const unreachable: never = draw.operation;
      throw new Error(`Unknown operation ${String(unreachable)}`);
    }
  }
};

const drawArbitrary: fc.Arbitrary<Draw> = fc.record({
  placement: fc.constantFrom<Placement>(
    "plain",
    "inserted",
    "deleted",
    "deletedInsideInsertion",
  ),
  math: fc.boolean(),
  numbered: fc.boolean(),
  mode: fc.constantFrom("direct" as const, "tracked-changes" as const),
  operation: fc.constantFrom<Draw["operation"]>(
    "deleteBlock",
    "replaceBlock",
    "replaceInBlock",
    "insertAfterBlock",
    "insertBeforeBlock",
    "mergeBlockWithNext",
    "mergeIntoIt",
    "setBlockParagraphProperties",
  ),
});

/** The replayed shape: a numbered paragraph, deleted directly. */
const REPLAYED: Draw = {
  placement: "deletedInsideInsertion",
  math: true,
  numbered: true,
  mode: "direct",
  operation: "deleteBlock",
};

const undoLaw = fc.asyncProperty(drawArbitrary, async (draw) => {
  const document = documentOf(draw);
  const probe = await open(await packaged(document));
  const blocks = probe.snapshot().blocks;
  const opening = blocks.find((block) => block.text === "Opening words");
  const noted = blocks.find((block) => block.text.startsWith("Noted words"));
  if (opening === undefined || noted === undefined)
    throw new Error("the drawn document's blocks are missing");
  const { reviewer, before, result } = await applyAndUndo(
    operationOf(draw, { opening: opening.id, noted: noted.id }),
    document,
    draw.mode,
  );
  // An operation the reviewer refuses leaves nothing to undo.
  if (result.undoHandle === null) return;
  expect(reviewer.undoDocumentOperations(result.undoHandle)).toEqual({
    status: "undone",
    undoHandle: result.undoHandle,
  });
  expect(reads(reviewer)).toEqual(before);
});

/** ADD_NOTE owns a body snapshot after the preceding deletion changes list counters. */
const nativeNoteBatchUndoLaw = async (kind: "footnote" | "endnote") => {
  const parsed = await parseDocx(await packaged(documentOf(REPLAYED)), {
    preloadFonts: false,
  });
  const session = createCanonicalSession(parsed).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const before = structuredClone(session.document);
  const beforeProjection = state.doc.toJSON();
  const commit = session
    .prepareOperations(state, [
      {
        type: documentOps.DOCUMENT_OP_TYPES.DELETE_BLOCKS,
        story: documentOps.OP_STORIES.MAIN,
        blockIds: ["00000002"],
      },
      {
        type: documentOps.DOCUMENT_OP_TYPES.ADD_NOTE,
        at: {
          story: documentOps.OP_STORIES.MAIN,
          blockId: "00000003",
          offset: 0,
        },
        note: {
          type: kind,
          id: 3,
          content: [
            {
              type: "paragraph",
              paraId: "7FFFFFF2",
              content: [
                { type: "run", content: [{ type: "text", text: "New note" }] },
              ],
            },
          ],
        },
      },
    ])
    .unwrap();
  let bodyRestores = 0;
  const unsubscribe = session.onCommitted(({ inverse }) => {
    bodyRestores += inverse.filter(
      (op) =>
        op.type === documentOps.DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS &&
        op.parts.body?.content !== undefined,
    ).length;
  });
  const published = publishCanonicalProjection({ session, state, commit });
  unsubscribe();
  state = published.unwrap().state;
  expect(bodyRestores).toBeGreaterThan(0);
  const after = structuredClone(session.document);
  state = publishCanonicalProjection({
    session,
    state,
    commit: session.prepareUndo(state).unwrap(),
  }).unwrap().state;
  expect(structuredClone(session.document)).toStrictEqual(before);
  expect(state.doc.toJSON()).toEqual(beforeProjection);
  publishCanonicalProjection({
    session,
    state,
    commit: session.prepareRedo(state).unwrap(),
  }).unwrap();
  expect(structuredClone(session.document)).toStrictEqual(after);
};

describe("undo of any public operation on a paragraph holding a note reference", () => {
  test(
    "gives back the document and every read",
    async () => {
      await assertProperty(undoLaw, {
        id: "undo gives back a paragraph with a note reference",
        numRuns: 40,
        examples: [[REPLAYED]],
      });
    },
    propertyTestTimeout(180_000),
  );

  test(
    "a native note lifecycle batch undoes after list rendering",
    async () => {
      for (const kind of ["footnote", "endnote"] as const)
        await nativeNoteBatchUndoLaw(kind);
    },
    propertyTestTimeout(60_000),
  );

  test(
    "the native batch law catches a restore that expects the parts before list rendering",
    async () => {
      // The test-only break: the session keeps each restore's expectation
      // as its operation recorded it, before list rendering ran.
      const original = { ...documentOps };
      let rebases = 0;
      mock.module("@stll/docx-core/ops", () => ({
        ...original,
        inverseMeetingDocument: (
          _document: Parameters<typeof original.inverseMeetingDocument>[0],
          inverse: Parameters<typeof original.inverseMeetingDocument>[1],
        ) => {
          rebases++;
          return inverse;
        },
      }));
      try {
        for (const kind of ["footnote", "endnote"] as const)
          await expect(nativeNoteBatchUndoLaw(kind)).rejects.toThrow(
            "Story lifecycle inverse is stale",
          );
        expect(rebases).toBeGreaterThan(0);
      } finally {
        mock.module("@stll/docx-core/ops", () => original);
      }
    },
    propertyTestTimeout(60_000),
  );
});
