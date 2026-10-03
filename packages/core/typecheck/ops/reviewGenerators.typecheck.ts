/** Tracked-family generators compiled by the package's typecheck contract. */
import type { BlockContent, Document, Paragraph, TableRow } from "@stll/docx-core/model";
import {
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  normalizeForOps,
  OP_STORIES,
  paragraphLength,
  paragraphLogicalText,
  type DocumentOp,
} from "@stll/docx-core/ops";
import { panic } from "better-result";
import fc from "fast-check";
import { parseDocx } from "../../src/docx/parser";
import { createDocx } from "../../src/docx/rezip";

type WithRevision<Op> = Op extends DocumentOp ? ("revision" extends keyof Op ? Op : never) : never;
type TrackedOp = WithRevision<DocumentOp>;

export const FIRST_ID = "00000001";
const SECOND_ID = "00000002";
const NESTED_ID = "00000003";
const ROW_ID = "00000010";
const NEW_ID = "00000020";
const DATE = "2026-05-06T07:08:09.000Z";

export const seedArbitrary = fc.record({
  text: fc.string({ unit: fc.constantFrom("a", "b", " ", "ž", "😀"), minLength: 2, maxLength: 12 }),
  insertion: fc.string({ unit: fc.constantFrom("x", "y", "ž"), minLength: 1, maxLength: 6 }),
  offset: fc.nat(),
  container: fc.constantFrom("body", "cell"),
  nesting: fc.constantFrom("plain", "insertion-deletion"),
  italic: fc.boolean(),
  alignment: fc.constantFrom("start", "end"),
});
type Seed = ReturnType<typeof seedArbitrary.generate>["value"];

type ParagraphOptions = { id: string; text: string; seed: Seed };
const paragraph = ({ id, text, seed }: ParagraphOptions): Paragraph => ({
  type: "paragraph",
  paraId: id,
  formatting: { alignment: seed.alignment },
  content: [
    {
      type: "run",
      ...(seed.italic ? { formatting: { italic: true } } : {}),
      content: [{ type: "text", text }],
    },
  ],
});

const insertedRow = (seed: Seed): TableRow => ({
  type: "tableRow",
  cells: [0, 1, 2].map((index) => ({
    type: "tableCell",
    content: [
      paragraph({
        id: (0x20 + index).toString(16).toUpperCase().padStart(8, "0"),
        text: seed.insertion,
        seed,
      }),
    ],
  })),
});

const tableAt = (blockId = ROW_ID) => ({ story: OP_STORIES.MAIN, blockId });

export const fixture = async (seed: Seed): Promise<Document> => {
  const nested = paragraph({ id: NESTED_ID, text: seed.text, seed });
  if (seed.nesting === "insertion-deletion") {
    nested.content = [
      {
        type: "insertion",
        info: { id: 90, author: "Earlier", date: DATE },
        content: [
          {
            type: "deletion",
            info: { id: 91, author: "Earlier", date: DATE },
            content: nested.content.filter((content) => content.type === "run"),
          },
        ],
      },
    ];
  }
  const targets: BlockContent[] = [
    paragraph({ id: FIRST_ID, text: seed.text, seed }),
    paragraph({ id: SECOND_ID, text: seed.text, seed }),
    nested,
    paragraph({ id: "00000004", text: "tail", seed }),
  ];
  const targetBlocks: BlockContent[] =
    seed.container === "body"
      ? targets
      : [
          {
            type: "table",
            rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: targets }] }],
          },
        ];
  const model: Document = {
    package: {
      document: {
        content: [
          ...targetBlocks,
          {
            type: "table",
            columnWidths: [100, 100, 100],
            rows: [0, 1, 2].map((index) => {
              const rowId = (0x10 + index).toString(16).toUpperCase().padStart(8, "0");
              const rightId = (0x30 + index).toString(16).toUpperCase().padStart(8, "0");
              return {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    formatting: { gridSpan: 2 },
                    content: [
                      index === 0
                        ? paragraph({ id: rowId, text: seed.text, seed })
                        : { type: "paragraph", paraId: rowId, content: [] },
                    ],
                  },
                  {
                    type: "tableCell",
                    content: [paragraph({ id: rightId, text: seed.text, seed })],
                  },
                ],
              };
            }),
          },
          paragraph({ id: "00000005", text: "outside", seed }),
        ],
      },
    },
  };
  return normalizeForOps(
    await parseDocx(await createDocx(model), {
      preloadFonts: false,
      detectVariables: false,
    }),
  );
};

type OperationOptions = { document: Document; seed: Seed };
export const stamp = { id: 1000, author: "Reviewer", date: "2026-06-07T08:09:10.000Z" };
export const newIds = { revision: Array.from({ length: 128 }, (_, index) => 1001 + index) };
export const at = (offset: number) => ({ story: OP_STORIES.MAIN, blockId: FIRST_ID, offset });

export const findParagraph = (
  blocks: readonly BlockContent[],
  id: string,
): Paragraph | undefined => {
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph":
        if (block.paraId === id) return block;
        break;
      case "table":
        for (const row of block.rows)
          for (const cell of row.cells) {
            const found = findParagraph(cell.content, id);
            if (found) return found;
          }
        break;
      case "blockSdt":
      case "blockCustomXml": {
        const found = findParagraph(block.content, id);
        if (found) return found;
        break;
      }
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default:
        return block satisfies never;
    }
  }
  return undefined;
};

const targetLength = (document: Document): number => {
  const target = findParagraph(document.package.document.content, FIRST_ID);
  if (!target) panic("Generated operation target disappeared");
  return paragraphLength(target);
};

/** Offsets at Unicode scalar boundaries, as a character-editing client produces. */
const targetOffset = ({ document, seed }: OperationOptions): number => {
  const target = findParagraph(document.package.document.content, FIRST_ID);
  if (!target) return panic("Generated operation target disappeared");
  const positions = [0];
  let offset = 0;
  for (const character of paragraphLogicalText(target)) {
    offset += character.length;
    positions.push(offset);
  }
  const selected = positions.at(seed.offset % positions.length);
  return selected ?? panic("Generated offset disappeared");
};

/** Total over the tracked members of DocumentOp; a new family requires a case. */
export const operations = {
  [DOCUMENT_OP_TYPES.INSERT_TEXT]: ({ document, seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: at(targetOffset({ document, seed })),
    text: seed.insertion,
    runProps: INHERIT_RUN_PROPS,
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.INSERT_CONTENT]: ({ document, seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at(targetOffset({ document, seed })),
    slice: {
      content: paragraph({ id: NEW_ID, text: seed.insertion, seed }).content,
      openStart: 0,
      openEnd: 0,
    },
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.DELETE_RANGE]: ({ document }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: at(0),
    to: at(targetLength(document)),
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.SET_RUN_PROPS]: ({ document }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
    from: at(0),
    to: at(targetLength(document)),
    patch: { bold: true },
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS]: () => ({
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: FIRST_ID,
    patch: { alignment: "center" },
    revision: stamp,
  }),
  [DOCUMENT_OP_TYPES.SPLIT_BLOCK]: ({ document, seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: at(targetOffset({ document, seed })),
    newBlockId: NEW_ID,
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.JOIN_BLOCKS]: () => ({
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: FIRST_ID,
    nextBlockId: SECOND_ID,
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.INSERT_BLOCKS]: ({ seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId: FIRST_ID },
    blocks: [paragraph({ id: NEW_ID, text: seed.insertion, seed })],
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.DELETE_BLOCKS]: ({ seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: seed.offset % 2 === 0 ? [FIRST_ID] : [FIRST_ID, SECOND_ID],
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.INSERT_TABLE]: ({ seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.INSERT_TABLE,
    story: OP_STORIES.MAIN,
    at: { type: seed.offset % 2 === 0 ? "before" : "after", blockId: FIRST_ID },
    table: { type: "table", columnWidths: [100, 100, 100], rows: [insertedRow(seed)] },
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.DELETE_TABLE]: ({ seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.DELETE_TABLE,
    story: OP_STORIES.MAIN,
    blockId: seed.container === "cell" ? FIRST_ID : ROW_ID,
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.INSERT_ROW]: ({ seed }: OperationOptions) => ({
    type: DOCUMENT_OP_TYPES.INSERT_ROW,
    story: OP_STORIES.MAIN,
    blockId: ROW_ID,
    at: seed.offset % 4,
    row: insertedRow(seed),
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.DELETE_ROW]: () => ({
    type: DOCUMENT_OP_TYPES.DELETE_ROW,
    story: OP_STORIES.MAIN,
    blockId: ROW_ID,
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.INSERT_COLUMN]: () => ({
    type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
    ...tableAt(),
    column: 2,
    width: 100,
    newBlockIds: [0x40, 0x41, 0x42].map((value) =>
      value.toString(16).toUpperCase().padStart(8, "0"),
    ),
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.DELETE_COLUMN]: () => ({
    type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
    ...tableAt(),
    column: 2,
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.MERGE_CELLS]: () => ({
    type: DOCUMENT_OP_TYPES.MERGE_CELLS,
    ...tableAt(),
    top: 0,
    bottom: 3,
    left: 0,
    right: 2,
    newBlockIds: [],
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.SPLIT_CELL]: () => ({
    type: DOCUMENT_OP_TYPES.SPLIT_CELL,
    ...tableAt(),
    newBlockIds: ["00000050"],
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.SET_TABLE_GRID]: () => ({
    type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
    ...tableAt(),
    columnWidths: [110, 90, 100],
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.SET_CELL_PROPS]: () => ({
    type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
    ...tableAt(),
    patch: { verticalAlign: "center" },
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.SET_ROW_PROPS]: () => ({
    type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
    ...tableAt(),
    patch: { cantSplit: true },
    revision: stamp,
    newIds,
  }),
  [DOCUMENT_OP_TYPES.SET_TABLE_PROPS]: () => ({
    type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
    ...tableAt(),
    patch: { justification: "center" },
    revision: stamp,
    newIds,
  }),
} as const satisfies {
  [Kind in TrackedOp["type"]]: (options: OperationOptions) => Extract<TrackedOp, { type: Kind }>;
};
