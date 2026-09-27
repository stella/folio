/**
 * The story dimension of the public operation set.
 *
 * Operations are written against one story at a time, and every story but the
 * main one reaches the package through its own save path: headers and footers
 * are re-serialized whole, notes are spliced into the original note part. A
 * test that drives an operation only against the body therefore says nothing
 * about the same operation in a note, and a note save that splices only the
 * paragraphs an edit changed silently keeps a deleted paragraph, a removed row
 * or a row-level revision it never looked at.
 *
 * Each operation here runs against the same shape of content in every
 * secondary story — two paragraphs, a 2x2 table, a closing paragraph — and the
 * receipt is held to what the package says after save and reopen: an applied
 * operation's outcome survives, and in tracked mode accepting everything gives
 * the direct outcome. Comments are checked the same way through
 * `getComments()`, whose anchors must come from the story that holds them.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_TYPES,
  type FolioDocumentOperation,
} from "../document-operations";
import type { Paragraph, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer, type FolioEditableDocumentStoryHandle } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIBlock } from "./types";

const HEADER_ID = "rIdStoryHeader";
const FOOTER_ID = "rIdStoryFooter";
const FOOTNOTE_ID = 10;
const ENDNOTE_ID = 11;

type StoryCase = {
  name: string;
  story: FolioEditableDocumentStoryHandle;
  /** Distinct paraId prefix, so every story's ids are unique in the package. */
  prefix: string;
};

const STORIES: readonly StoryCase[] = [
  { name: "header", story: { type: "header", relationshipId: HEADER_ID }, prefix: "51" },
  { name: "footer", story: { type: "footer", relationshipId: FOOTER_ID }, prefix: "52" },
  { name: "footnote", story: { type: "footnote", noteId: FOOTNOTE_ID }, prefix: "53" },
  { name: "endnote", story: { type: "endnote", noteId: ENDNOTE_ID }, prefix: "54" },
];

const paragraph = (text: string, paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const table = (prefix: string): Table => ({
  type: "table",
  columnWidths: [2400, 2400],
  rows: [
    ["A1", "B1"],
    ["A2", "B2"],
  ].map((cells, row) => ({
    type: "tableRow",
    cells: cells.map((text, cell) => ({
      type: "tableCell",
      content: [paragraph(text, `${prefix}0001${row}${cell}`)],
    })),
  })),
});

/** Alpha, Bravo, the 2x2 table, Charlie: every operation has a target. */
const storyContent = (prefix: string): (Paragraph | Table)[] => [
  paragraph("Alpha one.", `${prefix}000001`),
  paragraph("Bravo two.", `${prefix}000002`),
  table(prefix),
  paragraph("Charlie end.", `${prefix}000003`),
];

const buildSeed = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "50000001",
      content: [
        {
          type: "run",
          content: [
            { type: "text", text: "Body" },
            { type: "footnoteRef", id: FOOTNOTE_ID },
            { type: "text", text: " and" },
            { type: "endnoteRef", id: ENDNOTE_ID },
            { type: "text", text: "." },
          ],
        },
      ],
    },
  ];
  document.package.headers = new Map([
    [HEADER_ID, { type: "header", hdrFtrType: "default", content: storyContent("51") }],
  ]);
  document.package.footers = new Map([
    [FOOTER_ID, { type: "footer", hdrFtrType: "default", content: storyContent("52") }],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_ID }],
    footerReferences: [{ type: "default", rId: FOOTER_ID }],
  };
  document.package.footnotes = [
    { type: "footnote", id: FOOTNOTE_ID, noteType: "normal", content: storyContent("53") },
  ];
  document.package.endnotes = [
    { type: "endnote", id: ENDNOTE_ID, noteType: "normal", content: storyContent("54") },
  ];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer;
};

let seedPromise: Promise<ArrayBuffer> | undefined;
const seed = async (): Promise<ArrayBuffer> => (await (seedPromise ??= buildSeed())).slice(0);

/**
 * What one story shows a reader: its paragraphs (with direct alignment and
 * bold runs spelled out, so formatting operations have an outcome to check)
 * and its tables' cells by position.
 */
type StoryShape = { paragraphs: string[]; tables: string[][][] };

const paragraphLabel = (block: FolioAIBlock): string => {
  const runs = block.previewRuns;
  const text =
    runs === undefined
      ? block.text
      : runs.map((run) => (run.bold === true ? `**${run.text}**` : run.text)).join("");
  return block.directAlignment === undefined ? text : `${text} [${block.directAlignment}]`;
};

const shapeOf = (blocks: readonly FolioAIBlock[]): StoryShape => {
  const paragraphs: string[] = [];
  const tables: string[][][] = [];
  for (const block of blocks) {
    if (block.table === undefined) {
      paragraphs.push(paragraphLabel(block));
      continue;
    }
    const { tableIndex, rowIndex, cellIndex } = block.table;
    const rows = (tables[tableIndex] ??= []);
    const cells = (rows[rowIndex] ??= []);
    const existing = cells[cellIndex];
    cells[cellIndex] = existing === undefined ? block.text : `${existing}\n${block.text}`;
  }
  return { paragraphs, tables };
};

const storyShape = (reviewer: FolioDocxReviewer, story: FolioEditableDocumentStoryHandle) => {
  const snapshot = reviewer.snapshotStory(story);
  if (!snapshot) {
    throw new Error(`story ${JSON.stringify(story)} is missing`);
  }
  return shapeOf(snapshot.blocks);
};

const blockWithText = (
  reviewer: FolioDocxReviewer,
  story: FolioEditableDocumentStoryHandle,
  text: string,
): FolioAIBlock => {
  const block = reviewer.snapshotStory(story)?.blocks.find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads ${JSON.stringify(text)} in ${JSON.stringify(story)}`);
  }
  return block;
};

type Mode = "direct" | "tracked-changes";

const rangeOf = (blockId: string, text: string, startOffset: number, endOffset: number) => {
  const range = createFolioAITextRangeHandle({ blockId, text, startOffset, endOffset });
  if (!range) {
    throw new Error(`no range ${startOffset}..${endOffset} in ${JSON.stringify(text)}`);
  }
  return range;
};

const apply = (
  reviewer: FolioDocxReviewer,
  story: FolioEditableDocumentStoryHandle,
  mode: Mode,
  operation: FolioDocumentOperation,
) =>
  reviewer.applyDocumentOperationsToStory({
    story,
    batch: { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations: [operation] },
  });

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Story test" });

const ORIGINAL: StoryShape = {
  paragraphs: ["Alpha one.", "Bravo two.", "Charlie end."],
  tables: [
    [
      ["A1", "B1"],
      ["A2", "B2"],
    ],
  ],
};

/**
 * One operation and the story it should leave behind once applied (and, in
 * tracked mode, accepted). `target` names the block text the operation
 * anchors on; `build` receives the resolved block ids.
 */
type OperationCase = {
  type: FolioDocumentOperation["type"];
  label: string;
  build: (ids: (text: string) => string) => FolioDocumentOperation;
  /**
   * The story after the operation, or `"second table"` where only a new
   * table's presence is asserted. In tracked mode, accepting everything must
   * give the same story as the direct operation.
   */
  expected: StoryShape | "second table";
  /** Modes this operation supports; tracked is omitted where it reports `unsupportedMode`. */
  modes: readonly Mode[];
};

const BOTH: readonly Mode[] = ["direct", "tracked-changes"];

const OPERATION_CASES: readonly OperationCase[] = [
  {
    type: "replaceInBlock",
    label: "replace text in a paragraph",
    build: (ids) => ({
      id: "op",
      type: "replaceInBlock",
      blockId: ids("Alpha one."),
      find: "one",
      replace: "uno",
    }),
    expected: { ...ORIGINAL, paragraphs: ["Alpha uno.", "Bravo two.", "Charlie end."] },
    modes: BOTH,
  },
  {
    type: "replaceRange",
    label: "replace a range in a cell",
    build: (ids) => ({
      id: "op",
      type: "replaceRange",
      range: rangeOf(ids("B2"), "B2", 0, 2),
      replace: "Z9",
    }),
    expected: {
      ...ORIGINAL,
      tables: [
        [
          ["A1", "B1"],
          ["A2", "Z9"],
        ],
      ],
    },
    modes: BOTH,
  },
  {
    type: "formatRange",
    label: "format a range",
    build: (ids) => ({
      id: "op",
      type: "formatRange",
      range: rangeOf(ids("Bravo two."), "Bravo two.", 0, 5),
      formatting: { bold: true },
    }),
    expected: { ...ORIGINAL, paragraphs: ["Alpha one.", "**Bravo** two.", "Charlie end."] },
    modes: BOTH,
  },
  {
    type: "insertAfterBlock",
    label: "insert after a paragraph",
    build: (ids) => ({
      id: "op",
      type: "insertAfterBlock",
      blockId: ids("Alpha one."),
      text: "Inserted.",
    }),
    expected: {
      ...ORIGINAL,
      paragraphs: ["Alpha one.", "Inserted.", "Bravo two.", "Charlie end."],
    },
    modes: BOTH,
  },
  {
    type: "insertBeforeBlock",
    label: "insert before the first paragraph",
    build: (ids) => ({
      id: "op",
      type: "insertBeforeBlock",
      blockId: ids("Alpha one."),
      text: "Inserted.",
    }),
    expected: {
      ...ORIGINAL,
      paragraphs: ["Inserted.", "Alpha one.", "Bravo two.", "Charlie end."],
    },
    modes: BOTH,
  },
  {
    type: "replaceBlock",
    label: "replace a paragraph",
    build: (ids) => ({
      id: "op",
      type: "replaceBlock",
      blockId: ids("Bravo two."),
      text: "Bravo rewritten.",
    }),
    expected: { ...ORIGINAL, paragraphs: ["Alpha one.", "Bravo rewritten.", "Charlie end."] },
    modes: BOTH,
  },
  {
    type: "deleteBlock",
    label: "delete the first paragraph",
    build: (ids) => ({ id: "op", type: "deleteBlock", blockId: ids("Alpha one.") }),
    expected: { ...ORIGINAL, paragraphs: ["Bravo two.", "Charlie end."] },
    modes: BOTH,
  },
  {
    type: "splitBlock",
    label: "split a paragraph",
    build: (ids) => ({
      id: "op",
      type: "splitBlock",
      blockId: ids("Alpha one."),
      offset: 5,
      separator: " ",
    }),
    expected: { ...ORIGINAL, paragraphs: ["Alpha", "one.", "Bravo two.", "Charlie end."] },
    modes: BOTH,
  },
  {
    type: "mergeBlockWithNext",
    label: "merge two paragraphs",
    build: (ids) => ({
      id: "op",
      type: "mergeBlockWithNext",
      blockId: ids("Alpha one."),
      separator: " ",
    }),
    expected: { ...ORIGINAL, paragraphs: ["Alpha one. Bravo two.", "Charlie end."] },
    modes: BOTH,
  },
  {
    type: "setBlockParagraphProperties",
    label: "set paragraph alignment",
    build: (ids) => ({
      id: "op",
      type: "setBlockParagraphProperties",
      blockId: ids("Bravo two."),
      properties: { alignment: "center" },
    }),
    expected: { ...ORIGINAL, paragraphs: ["Alpha one.", "Bravo two. [center]", "Charlie end."] },
    modes: BOTH,
  },
  {
    type: "insertTable",
    label: "insert a table",
    build: (ids) => ({
      id: "op",
      type: "insertTable",
      blockId: ids("Charlie end."),
      rows: [["N1", "N2"]],
    }),
    expected: { ...ORIGINAL, tables: [...ORIGINAL.tables, [["N1", "N2"]]] },
    modes: BOTH,
  },
  {
    type: "deleteTable",
    label: "delete the table",
    build: (ids) => ({ id: "op", type: "deleteTable", blockId: ids("A1") }),
    expected: { ...ORIGINAL, tables: [] },
    modes: BOTH,
  },
  {
    type: "insertSignatureTable",
    label: "insert a signature table",
    build: (ids) => ({
      id: "op",
      type: "insertSignatureTable",
      blockId: ids("Charlie end."),
      parties: [{ name: "Party" }],
    }),
    // The signature block's spacer and rule paragraphs are the template's to
    // decide; the story must show a second table led by the party's name.
    expected: "second table",
    modes: ["direct"],
  },
  {
    type: "insertTableRow",
    label: "insert a row",
    build: (ids) => ({
      id: "op",
      type: "insertTableRow",
      blockId: ids("A1"),
      cellTexts: ["R1", "R2"],
    }),
    expected: {
      ...ORIGINAL,
      tables: [
        [
          ["A1", "B1"],
          ["R1", "R2"],
          ["A2", "B2"],
        ],
      ],
    },
    modes: BOTH,
  },
  {
    type: "deleteTableRow",
    label: "delete a row",
    build: (ids) => ({ id: "op", type: "deleteTableRow", blockId: ids("A1") }),
    expected: { ...ORIGINAL, tables: [[["A2", "B2"]]] },
    modes: BOTH,
  },
  {
    type: "insertTableColumn",
    label: "insert a column",
    build: (ids) => ({
      id: "op",
      type: "insertTableColumn",
      blockId: ids("A1"),
      cellTexts: ["C1", "C2"],
    }),
    expected: {
      ...ORIGINAL,
      tables: [
        [
          ["A1", "C1", "B1"],
          ["A2", "C2", "B2"],
        ],
      ],
    },
    modes: BOTH,
  },
  {
    type: "deleteTableColumn",
    label: "delete a column",
    build: (ids) => ({ id: "op", type: "deleteTableColumn", blockId: ids("A1") }),
    expected: { ...ORIGINAL, tables: [[["B1"], ["B2"]]] },
    modes: BOTH,
  },
  {
    type: "mergeTableCells",
    label: "merge a row's cells",
    build: (ids) => ({
      id: "op",
      type: "mergeTableCells",
      blockId: ids("A1"),
      endBlockId: ids("B1"),
    }),
    expected: {
      ...ORIGINAL,
      tables: [[["A1\nB1"], ["A2", "B2"]]],
    },
    modes: ["direct"],
  },
];

/** Comment operations are checked through `getComments()` below. */
const COMMENT_OPERATION_TYPES = new Set<FolioDocumentOperation["type"]>([
  "commentOnBlock",
  "commentOnRange",
]);
/** Needs a merged cell to act on; covered as the reverse of the merge. */
const SPLIT_AFTER_MERGE = "splitTableCell";

describe("every public operation has a story-dimension case", () => {
  test("no operation type is left out of the matrix", () => {
    const covered = new Set<string>([
      ...OPERATION_CASES.map(({ type }) => type),
      ...COMMENT_OPERATION_TYPES,
      SPLIT_AFTER_MERGE,
    ]);
    expect(FOLIO_DOCUMENT_OPERATION_TYPES.filter((type) => !covered.has(type))).toEqual([]);
  });
});

const expectOutcome = (shape: StoryShape, expected: OperationCase["expected"]): void => {
  if (expected !== "second table") {
    expect(shape).toEqual(expected);
    return;
  }
  expect(shape.paragraphs).toEqual(ORIGINAL.paragraphs);
  expect(shape.tables[0]).toEqual(ORIGINAL.tables[0]);
  expect(shape.tables[1]?.[0]?.[0]?.split("\n")[0]).toBe("Party");
};

describe.each(STORIES)("operations in the $name story survive save and reopen", (storyCase) => {
  const { story } = storyCase;
  const cases = OPERATION_CASES.flatMap((operation) =>
    operation.modes.map((mode) => ({ ...operation, mode })),
  );

  test.each(cases)("$label ($mode)", async (operation) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await seed(), { author: "Story test" });
    const ids = (text: string) => blockWithText(reviewer, story, text).id;
    const result = apply(reviewer, story, operation.mode, operation.build(ids));
    expect(result.issues).toEqual([]);
    expect(result.applied.map(({ id }) => id)).toEqual(["op"]);

    const reopened = await reopen(reviewer);
    if (operation.mode === "direct") {
      expectOutcome(storyShape(reviewer, story), operation.expected);
      expectOutcome(storyShape(reopened, story), operation.expected);
      return;
    }
    reopened.acceptAll();
    expectOutcome(storyShape(reopened, story), operation.expected);
    expectOutcome(storyShape(await reopen(reopened), story), operation.expected);
  });

  test("a split cell comes back split after a merge", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await seed(), { author: "Story test" });
    const ids = (text: string) => blockWithText(reviewer, story, text).id;
    apply(reviewer, story, "direct", {
      id: "merge",
      type: "mergeTableCells",
      blockId: ids("A1"),
      endBlockId: ids("B1"),
    });
    const merged = await reopen(reviewer);
    const result = apply(merged, story, "direct", {
      id: "op",
      type: "splitTableCell",
      blockId: blockWithText(merged, story, "A1").id,
    });
    expect(result.issues).toEqual([]);
    expect(result.applied.map(({ id }) => id)).toEqual(["op"]);
    const expected = storyShape(merged, story);
    expect(expected.tables[0]?.[0]?.length).toBe(2);
    expect(storyShape(await reopen(merged), story)).toEqual(expected);
  });

  test.each(["commentOnBlock", "commentOnRange"] as const)(
    "%s anchors are read back from the story",
    async (type) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await seed(), { author: "Story test" });
      const block = blockWithText(reviewer, story, "Bravo two.");
      const operation: FolioDocumentOperation =
        type === "commentOnBlock"
          ? { id: "op", type, blockId: block.id, comment: { text: "Check." } }
          : {
              id: "op",
              type,
              range: rangeOf(block.id, block.text, 0, 5),
              comment: { text: "Check." },
            };
      const result = apply(reviewer, story, "direct", operation);
      expect(result.issues).toEqual([]);
      const quote = type === "commentOnBlock" ? "Bravo two." : "Bravo";
      const expected = { anchoredText: quote, blockId: block.id, story, text: "Check." };
      const pick = (comments: ReturnType<FolioDocxReviewer["getComments"]>) =>
        comments.map(({ anchoredText, blockId, story: anchorStory, text }) => ({
          anchoredText,
          blockId,
          story: anchorStory,
          text,
        }));
      expect(pick(reviewer.getComments())).toEqual([expected]);
      const reopened = await reopen(reviewer);
      expect(pick(reopened.getComments())).toEqual([
        { ...expected, blockId: blockWithText(reopened, story, "Bravo two.").id },
      ]);
    },
  );
});

const NOTES = STORIES.filter(({ story }) => story.type === "footnote" || story.type === "endnote");
const NOTE_PART = { footnote: "word/footnotes.xml", endnote: "word/endnotes.xml" } as const;

const notePartXml = async (buffer: ArrayBuffer, type: "footnote" | "endnote"): Promise<string> => {
  const file = (await JSZip.loadAsync(buffer)).file(NOTE_PART[type]);
  if (!file) {
    throw new Error(`missing ${NOTE_PART[type]}`);
  }
  return file.async("text");
};

/**
 * A note part is spliced into the original rather than rewritten, so what an
 * edit removed — a paragraph, a table, a row's revision mark — must be read
 * back from the part itself, not only from a reopened reader.
 */
describe.each(NOTES)("the $name part records what an edit removed", ({ story }) => {
  const type = story.type === "endnote" ? "endnote" : "footnote";

  test("a deleted paragraph with no other edit is gone from the saved part", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await seed(), { author: "Story test" });
    const deleted = blockWithText(reviewer, story, "Alpha one.");
    apply(reviewer, story, "direct", { id: "op", type: "deleteBlock", blockId: deleted.id });
    const xml = await notePartXml(await reviewer.toBuffer(), type);
    expect(xml).not.toContain("Alpha one.");
    expect(xml).not.toContain(deleted.id);
    expect(xml).toContain("Bravo two.");
    expect(xml).toContain('w:type="separator"');
  });

  test("a tracked table deletion keeps its row revisions through save", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await seed(), { author: "Story test" });
    apply(reviewer, story, "tracked-changes", {
      id: "op",
      type: "deleteTable",
      blockId: blockWithText(reviewer, story, "A1").id,
    });
    const pending = await reopen(reviewer);
    const changes = pending.readReviewedStory({ story, view: "current-markup" })?.changes ?? [];
    expect(changes.filter((change) => change.type === "rowDeleted")).toHaveLength(2);
    pending.acceptAll();
    expect(storyShape(pending, story).tables).toEqual([]);
    expect(await notePartXml(await pending.toBuffer(), type)).not.toContain("<w:tbl>");
  });
});
