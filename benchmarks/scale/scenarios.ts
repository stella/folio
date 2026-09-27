import { createDocx, FolioDocxReviewer } from "@stll/folio-core/server";
import { compareDocx } from "@stll/folio-core/compare/compare";
import { layoutDocument, type FlowBlock, type Measure } from "@stll/folio-core/layout-engine";
import type { Document } from "@stll/folio-core";

import {
  commentDocument,
  paragraphDocument,
  tableDocument,
  trackedChangeDocument,
} from "./fixtures";

type Operation = () => Promise<unknown> | unknown;

export type ScaleScenario = {
  readonly name: string;
  readonly sizes: readonly [number, number, number];
  readonly prepare: (size: number) => Promise<Operation> | Operation;
};

const PARAGRAPH_SIZES = [2_500, 5_000, 10_000] as const;
const TABLE_SIZES = [125, 250, 500] as const;
const ANNOTATION_SIZES = [250, 500, 1_000] as const;
const COMPARE_OPTIONS = {
  author: "Scale fixture",
  timestamp: "2000-01-01T00:00:00.000Z",
  onUnverified: "emit",
} as const;

const byteCache = new Map<string, Promise<ArrayBuffer>>();

const bytesFor = (build: (size: number) => Document, size: number): Promise<ArrayBuffer> => {
  const key = `${build.name}:${size}`;
  let bytes = byteCache.get(key);
  if (!bytes) {
    bytes = createDocx(build(size));
    byteCache.set(key, bytes);
  }
  return bytes;
};

type ReadArgs = {
  build: (size: number) => Document;
  size: number;
  expectedBlocks: number;
};

const prepareRead = async ({ build, size, expectedBlocks }: ReadArgs): Promise<Operation> => {
  const bytes = await bytesFor(build, size);
  return async () => {
    const blocks = (await FolioDocxReviewer.fromBuffer(bytes)).getContent();
    if (blocks.length !== expectedBlocks)
      throw new Error(`Expected ${expectedBlocks} content blocks, got ${blocks.length}`);
    return blocks;
  };
};

const prepareEdit = async (build: (size: number) => Document, size: number): Promise<Operation> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await bytesFor(build, size));
  const blocks = reviewer.getContent();
  const block = blocks.at(Math.floor(blocks.length / 2));
  if (!block) throw new Error("The scale fixture has no middle block");
  return () => {
    const result = reviewer.applyOperations(
      [
        {
          id: "scale-edit",
          type: "replaceInBlock",
          blockId: block.id,
          find: "scale",
          replace: "measured",
        },
      ],
      { mode: "direct" },
    );
    if (result.applied.length !== 1)
      throw new Error(`Scale edit was skipped: ${JSON.stringify(result.skipped)}`);
  };
};

const prepareTableEdit = async (size: number): Promise<Operation> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await bytesFor(tableDocument, size));
  const block = reviewer.getContent().find(({ text }) => text === `Row ${Math.floor(size / 2)} A`);
  if (!block) throw new Error("The scale table has no middle cell");
  return () => {
    const result = reviewer.applyOperations(
      [
        {
          id: "scale-table-edit",
          type: "replaceInBlock",
          blockId: block.id,
          find: "Row",
          replace: "Edited row",
        },
      ],
      { mode: "direct" },
    );
    if (result.applied.length !== 1)
      throw new Error(`Scale table edit was skipped: ${JSON.stringify(result.skipped)}`);
  };
};

const prepareSave = async (build: (size: number) => Document, size: number): Promise<Operation> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await bytesFor(build, size));
  return async () => {
    const bytes = await reviewer.toBuffer();
    if (bytes.byteLength === 0) throw new Error("Scale save produced an empty package");
    return bytes;
  };
};

const prepareCommentRead = async (size: number): Promise<Operation> => {
  const bytes = await bytesFor(commentDocument, size);
  return async () => {
    const comments = (await FolioDocxReviewer.fromBuffer(bytes)).getComments();
    if (comments.length !== size)
      throw new Error(`Expected ${size} comments, got ${comments.length}`);
    return comments;
  };
};

const prepareChangeRead = async (size: number): Promise<Operation> => {
  const bytes = await bytesFor(trackedChangeDocument, size);
  return async () => {
    const changes = (await FolioDocxReviewer.fromBuffer(bytes)).getChanges();
    if (changes.length !== size) throw new Error(`Expected ${size} changes, got ${changes.length}`);
    return changes;
  };
};

const prepareAcceptAll = async (size: number): Promise<Operation> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await bytesFor(trackedChangeDocument, size));
  return () => {
    const count = reviewer.acceptAll();
    if (count !== size) throw new Error(`Expected ${size} accepted changes, got ${count}`);
  };
};

const prepareCompare = async (size: number): Promise<Operation> => {
  const base = paragraphDocument(size);
  const target = paragraphDocument(size);
  const middle = target.package.document.content.at(Math.floor(size / 2));
  if (!middle || middle.type !== "paragraph")
    throw new Error("Missing comparison target paragraph");
  middle.content.push({ type: "run", content: [{ type: "text", text: " Revised." }] });
  const [baseBytes, targetBytes] = await Promise.all([createDocx(base), createDocx(target)]);
  return async () => {
    const result = await compareDocx(baseBytes, targetBytes, COMPARE_OPTIONS);
    if (result.isErr()) throw result.error;
    if (result.value.changes.length === 0) throw new Error("Scale compare found no changes");
    return result.value.changes.length;
  };
};

const layoutInput = (size: number): { blocks: FlowBlock[]; measures: Measure[] } => {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  let pmStart = 0;
  for (let index = 0; index < size; index += 1) {
    const text = `Paragraph ${index} has one measured line.`;
    blocks.push({
      kind: "paragraph",
      id: index,
      runs: [{ kind: "text", text, pmStart, pmEnd: pmStart + text.length }],
      attrs: {},
      pmStart,
      pmEnd: pmStart + text.length + 1,
    });
    measures.push({
      kind: "paragraph",
      lines: [
        {
          fromRun: 0,
          fromChar: 0,
          toRun: 0,
          toChar: text.length,
          width: 300,
          ascent: 14,
          descent: 4,
          lineHeight: 18,
        },
      ],
      totalHeight: 18,
    });
    pmStart += text.length + 1;
  }
  return { blocks, measures };
};

const prepareLayout = (size: number): Operation => {
  const { blocks, measures } = layoutInput(size);
  return () => {
    const layout = layoutDocument(blocks, measures, {
      pageSize: { w: 816, h: 1056 },
      margins: { top: 96, right: 96, bottom: 96, left: 96 },
    });
    if (layout.pages.length === 0) throw new Error("Scale layout produced no pages");
    return layout;
  };
};

export const SCALE_SCENARIOS: readonly ScaleScenario[] = [
  {
    name: "paragraph-read",
    sizes: PARAGRAPH_SIZES,
    prepare: (size) => prepareRead({ build: paragraphDocument, size, expectedBlocks: size }),
  },
  {
    name: "paragraph-edit",
    sizes: PARAGRAPH_SIZES,
    prepare: (size) => prepareEdit(paragraphDocument, size),
  },
  {
    name: "paragraph-save",
    sizes: PARAGRAPH_SIZES,
    prepare: (size) => prepareSave(paragraphDocument, size),
  },
  { name: "paragraph-compare", sizes: PARAGRAPH_SIZES, prepare: prepareCompare },
  { name: "paragraph-layout", sizes: PARAGRAPH_SIZES, prepare: prepareLayout },
  {
    name: "table-read",
    sizes: TABLE_SIZES,
    prepare: (size) => prepareRead({ build: tableDocument, size, expectedBlocks: 2 * (size + 1) }),
  },
  { name: "table-edit", sizes: TABLE_SIZES, prepare: prepareTableEdit },
  { name: "table-save", sizes: TABLE_SIZES, prepare: (size) => prepareSave(tableDocument, size) },
  { name: "comment-read", sizes: ANNOTATION_SIZES, prepare: prepareCommentRead },
  {
    name: "comment-save",
    sizes: ANNOTATION_SIZES,
    prepare: (size) => prepareSave(commentDocument, size),
  },
  { name: "change-read", sizes: ANNOTATION_SIZES, prepare: prepareChangeRead },
  {
    name: "change-save",
    sizes: ANNOTATION_SIZES,
    prepare: (size) => prepareSave(trackedChangeDocument, size),
  },
  { name: "change-accept-all", sizes: ANNOTATION_SIZES, prepare: prepareAcceptAll },
];
