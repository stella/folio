/** Reproducible, bounded operation sequences over seeded corpus models. */
import { isDeepStrictEqual } from "node:util";
import { panic } from "better-result";
import type { Document, Paragraph, Table } from "../../../packages/docx-core/src/model/document";
import {
  applyDocumentOp,
  normalizeForOps,
  validateOpsDocument,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  INHERIT_RUN_PROPS,
  REVISION_DECISIONS,
  type DocumentOp,
  type AppliedDocumentOp,
  type NoteStory,
  sectionPropertiesAt,
  type TableEditOp,
} from "../../../packages/docx-core/src/ops/documentOps";
import { cloneDocumentWithParagraphPropertySources } from "@stll/folio-core/docx/document-clone";
import { ensureParaIds } from "@stll/folio-core/docx/ensureParaIds";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { unzipDocx } from "@stll/folio-core/docx/unzip";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { serializeDocument } from "@stll/folio-core/docx/serializer/documentSerializer";
import { storyParagraphs, blockListAt } from "../../../packages/docx-core/src/ops/blocks";
import { leafSpans } from "../../../packages/docx-core/src/ops/leaves";
import { tableGrid } from "../../../packages/docx-core/src/ops/tableGrid";
import { paragraphLength } from "../../../packages/docx-core/src/ops/offsets";
import {
  packageIdentityKeys,
  packageParagraphIds,
  idKey,
} from "../../../packages/docx-core/src/ops/ids";
import type { CorpusInvariantInput } from "./contract";

export const OP_SEQUENCE_SEEDS = [0x17a3, 0x5b91, 0xcf27] as const;
const MAX_NEW_IDS = 64;

/** Exact records, including captures and binary contents; undefined is wire absence. */
export const exactOpModel = (value: unknown): unknown => {
  if (value instanceof Uint8Array) return value.slice();
  if (value instanceof ArrayBuffer) return value.slice(0);
  if (value instanceof Date) return new Date(value);
  if (Array.isArray(value)) return value.map(exactOpModel);
  if (value instanceof Map)
    return new Map([...value].map(([key, entry]) => [key, exactOpModel(entry)]));
  if (
    typeof value !== "object" ||
    value === null ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, entry]) => entry !== undefined)
      .map(([key, entry]) => [key, exactOpModel(entry)]),
  );
};

export const sameOpModel = (left: unknown, right: unknown): boolean =>
  isDeepStrictEqual(exactOpModel(left), exactOpModel(right));

export const serializeOpDocument = (document: Document): string => serializeDocument(document);

/** ZIP metadata/compression is transport; compare all uncompressed part bytes. */
export const serializedOpParts = async (document: Document): Promise<Map<string, Uint8Array>> => {
  const saved = await repackDocx(document, {
    updateModifiedDate: false,
    onDiagnostic: ({ type, part }) => panic(`${type}: ${part}`),
  });
  const raw = await unzipDocx(saved, { extractAllXml: true });
  return new Map(
    await Promise.all(
      Object.entries(raw.originalZip.files)
        .filter(([, file]) => !file.dir)
        .map(async ([path, file]) => [path, await file.async("uint8array")] as const),
    ),
  );
};

export const seedFromBytes = (bytes: Uint8Array): number => {
  let hash = 2166136261;
  for (const byte of bytes) hash = Math.imul(hash ^ byte, 16777619);
  return hash >>> 0;
};

const prepared = new WeakMap<Document, Promise<Document>>();
/** Establish the public operations contract before measuring edits, once per parse. */
export const prepareOpDocument = (input: CorpusInvariantInput): Promise<Document> => {
  const existing = prepared.get(input.parsed);
  if (existing !== undefined) return existing;
  const pending = (async () => {
    const seeded = await ensureParaIds(input.bytes);
    const parsed = seeded.alreadyComplete
      ? input.parsed
      : await parseDocx(Uint8Array.from(seeded.docx).buffer, { preloadFonts: false });
    const document = normalizeForOps(parsed);
    const valid = validateOpsDocument(document);
    if (valid.isErr()) throw valid.error;
    return document;
  })();
  prepared.set(input.parsed, pending);
  return pending;
};

const randomFor = (seed: number) => {
  let state = seed >>> 0;
  return (limit: number): number => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return limit === 0 ? 0 : state % limit;
  };
};

/** Every schema member needs a generator decision when the operations API grows. */
export const OP_GENERATOR_ROLES = {
  [DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER]: "generated",
  [DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER]: "generated",
  [DOCUMENT_OP_TYPES.ADD_NOTE]: "generated",
  [DOCUMENT_OP_TYPES.REMOVE_NOTE]: "generated",
  [DOCUMENT_OP_TYPES.SET_SECTION_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE]: "inverse",
  [DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT]: "inverse",
  [DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS]: "inverse",
  [DOCUMENT_OP_TYPES.INSERT_TEXT]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_CONTENT]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_RANGE]: "generated",
  [DOCUMENT_OP_TYPES.SET_RUN_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.SPLIT_BLOCK]: "generated",
  [DOCUMENT_OP_TYPES.JOIN_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_TABLE]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_TABLE]: "generated",
  [DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_ROW]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_ROW]: "generated",
  [DOCUMENT_OP_TYPES.RESOLVE_REVISION]: "generated",
  [DOCUMENT_OP_TYPES.INSERT_COLUMN]: "generated",
  [DOCUMENT_OP_TYPES.DELETE_COLUMN]: "generated",
  [DOCUMENT_OP_TYPES.MERGE_CELLS]: "generated",
  [DOCUMENT_OP_TYPES.SPLIT_CELL]: "generated",
  [DOCUMENT_OP_TYPES.SET_TABLE_GRID]: "generated",
  [DOCUMENT_OP_TYPES.SET_CELL_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.SET_ROW_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.SET_TABLE_PROPS]: "generated",
  [DOCUMENT_OP_TYPES.SPLIT_INLINE]: "inverse",
  [DOCUMENT_OP_TYPES.JOIN_INLINE]: "inverse",
  [DOCUMENT_OP_TYPES.REPLACE_BLOCKS]: "inverse",
  [DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW]: "inverse",
  [DOCUMENT_OP_TYPES.REPLACE_INLINE]: "inverse",
  [DOCUMENT_OP_TYPES.SET_TABLE_ROWS]: "inverse",
  [DOCUMENT_OP_TYPES.SET_TABLE]: "inverse",
} as const satisfies Record<DocumentOp["type"], "generated" | "inverse">;

type GeneratedOpType = {
  [Kind in keyof typeof OP_GENERATOR_ROLES]: (typeof OP_GENERATOR_ROLES)[Kind] extends "generated"
    ? Kind
    : never;
}[keyof typeof OP_GENERATOR_ROLES];
const isGeneratedFamily = (family: DocumentOp["type"]): family is GeneratedOpType =>
  OP_GENERATOR_ROLES[family] === "generated";
export const OP_SEQUENCE_FAMILIES = Object.values(DOCUMENT_OP_TYPES).filter(isGeneratedFamily);
export const OP_SEQUENCE_LENGTH = OP_SEQUENCE_FAMILIES.length;

/** The semantic subset also has a total decision, including newly added edit kinds. */
const TABLE_EDIT_FAMILIES = {
  insertColumn: true,
  deleteColumn: true,
  mergeCells: true,
  splitCell: true,
  setTableGrid: true,
  setCellProps: true,
  setRowProps: true,
  setTableProps: true,
} as const satisfies Record<TableEditOp["type"], true>;
const isTableEditFamily = (family: GeneratedOpType): family is TableEditOp["type"] =>
  Object.hasOwn(TABLE_EDIT_FAMILIES, family);

/** Resolve the table path already captured by the story walker, without rescanning the story. */
const tableForParagraph = (
  document: Document,
  location: ReturnType<typeof storyParagraphs>[number],
) => {
  const stepIndex = location.list.findLastIndex((step) => step.kind === "tableCell");
  const step = location.list.at(stepIndex);
  if (step?.kind !== "tableCell") return undefined;
  const table = blockListAt(
    document.package.document.content,
    location.list.slice(0, stepIndex),
  ).at(step.block);
  if (table?.type !== "table") return undefined;
  const row = table.rows.at(step.row);
  const cell = row?.cells.at(step.cell);
  if (!row || !cell) return undefined;
  return { table, row, cell, rowIndex: step.row, cellIndex: step.cell };
};
const inferredGridWidth = (table: Table) => {
  const first = table.rows.at(0);
  return (
    table.columnWidths?.length ??
    (first?.formatting?.gridBefore ?? 0) +
      (first?.cells.reduce((sum, cell) => sum + (cell.formatting?.gridSpan ?? 1), 0) ?? 0) +
      (first?.formatting?.gridAfter ?? 0)
  );
};

type CandidateOptions = {
  document: Document;
  choose: (limit: number) => number;
  family: (typeof OP_SEQUENCE_FAMILIES)[number];
  revisions: readonly number[];
  tracked: boolean;
};
const candidate = ({
  document,
  choose,
  family,
  revisions,
  tracked,
}: CandidateOptions): DocumentOp | undefined => {
  const paragraphs = storyParagraphs(document.package.document);
  let candidates = paragraphs;
  if (isTableEditFamily(family)) {
    const inTables = paragraphs.filter(
      (candidateLocation) => tableForParagraph(document, candidateLocation) !== undefined,
    );
    if (inTables.length > 0) candidates = inTables;
    if (family === DOCUMENT_OP_TYPES.SPLIT_CELL) {
      const horizontalSpans = candidates.filter((candidateLocation) => {
        const target = tableForParagraph(document, candidateLocation);
        return (
          (target?.cell.formatting?.gridSpan ?? 1) > 1 &&
          target?.cell.formatting?.vMerge === undefined
        );
      });
      if (horizontalSpans.length > 0) candidates = horizontalSpans;
    }
  }
  let location = candidates.at(choose(candidates.length));
  if (location === undefined && isTableEditFamily(family)) {
    // No main-story paragraph can anchor this family: apply still emits its
    // structured blockNotFound refusal rather than silently skipping coverage.
    location = {
      list: [],
      index: 0,
      paragraph: { type: "paragraph", paraId: "00000001", content: [] },
    };
  }
  if (location === undefined || location.paragraph.paraId === undefined) return undefined;
  const paragraph = location.paragraph;
  const blockId = location.paragraph.paraId;
  const length = paragraphLength(paragraph);
  const at = { story: OP_STORIES.MAIN, blockId, offset: choose(length + 1) };
  const ids = new Set(packageParagraphIds(document.package).map(idKey));
  let paraId = 1;
  while (ids.has(idKey(paraId.toString(16).padStart(8, "0")))) paraId += 1;
  const newBlockId = paraId.toString(16).padStart(8, "0");
  // Oversized tables exhaust the bounded pool and are explicitly refused by apply;
  // never allocate unbounded paragraph ids while probing a corpus file.
  const freshBlockIds = (count: number) => {
    const fresh: string[] = [];
    let next = paraId;
    while (fresh.length < Math.min(count, MAX_NEW_IDS)) {
      const value = next.toString(16).padStart(8, "0");
      if (!ids.has(idKey(value))) fresh.push(value);
      next += 1;
    }
    return fresh;
  };
  const taken = new Set(packageIdentityKeys(document.package));
  const freshRevision: number[] = [];
  const freshControl: number[] = [];
  for (
    let id = 1;
    freshRevision.length < MAX_NEW_IDS || freshControl.length < MAX_NEW_IDS;
    id += 1
  ) {
    if (freshRevision.length < MAX_NEW_IDS && !taken.has(`revision:${id}`)) freshRevision.push(id);
    if (freshControl.length < MAX_NEW_IDS && !taken.has(`control:${id}`)) freshControl.push(id);
  }
  const revisionId = freshRevision.at(0);
  if (revisionId === undefined) return undefined;
  const review = tracked
    ? {
        revision: { id: revisionId, author: "Corpus", date: "2026-01-01T00:00:00Z" },
      }
    : {};
  const newIds = { revision: freshRevision.slice(1), control: freshControl };
  const freshParagraph = {
    type: "paragraph",
    paraId: newBlockId,
    content: [{ type: "run", content: [{ type: "text", text: "‸" }] }],
  } as const satisfies Paragraph;
  const end = { ...at, offset: Math.min(length, at.offset + 1 + choose(3)) };
  const sectionCount =
    document.package.document.content.filter(
      (block) => block.type === "paragraph" && block.sectionProperties !== undefined,
    ).length + 1;
  switch (family) {
    case "createHeaderFooter": {
      const sectionIndex = choose(sectionCount);
      const properties = sectionPropertiesAt(document, sectionIndex);
      if (!properties) return undefined;
      const kind = choose(2) === 0 ? "header" : "footer";
      const references =
        kind === "header" ? properties.headerReferences : properties.footerReferences;
      const availableVariants = (["default", "first", "even"] as const).filter(
        (variant) => !references?.some(({ type }) => type === variant),
      );
      const referenceType = availableVariants.at(choose(availableVariants.length));
      if (!referenceType) return undefined;
      let identity = 1;
      const used = (rId: string) =>
        document.package.headers?.has(rId) ||
        document.package.footers?.has(rId) ||
        document.package.relationships?.has(rId);
      while (used(`rIdCorpus${identity}`)) identity += 1;
      return {
        type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
        sectionIndex,
        story: { kind, rId: `rIdCorpus${identity}` },
        referenceType,
        content: [freshParagraph],
      };
    }
    case "removeHeaderFooter": {
      const removableStories = Array.from({ length: sectionCount }, (_, sectionIndex) => {
        const properties = sectionPropertiesAt(document, sectionIndex);
        const headers =
          properties?.headerReferences
            ?.filter(({ rId }) => document.package.headers?.has(rId))
            .map(
              ({ type, rId }) =>
                ({ sectionIndex, story: { kind: "header", rId }, referenceType: type }) as const,
            ) ?? [];
        const footers =
          properties?.footerReferences
            ?.filter(({ rId }) => document.package.footers?.has(rId))
            .map(
              ({ type, rId }) =>
                ({ sectionIndex, story: { kind: "footer", rId }, referenceType: type }) as const,
            ) ?? [];
        return [...headers, ...footers];
      }).flat();
      const selected = removableStories.at(choose(removableStories.length));
      return selected ? { type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER, ...selected } : undefined;
    }
    case "addNote": {
      const kind = choose(2) === 0 ? "footnote" : "endnote";
      const collection =
        kind === "footnote" ? document.package.footnotes : document.package.endnotes;
      let id = 1;
      while (collection?.some((note) => note.id === id)) id += 1;
      return {
        type: DOCUMENT_OP_TYPES.ADD_NOTE,
        at,
        note: { type: kind, id, content: [freshParagraph] },
      };
    }
    case "removeNote": {
      const references = paragraphs.flatMap(({ paragraph: referenceParagraph }) =>
        leafSpans(referenceParagraph.content).flatMap(({ node, before }) => {
          if (node.type !== "footnoteRef" && node.type !== "endnoteRef") return [];
          if (!referenceParagraph.paraId) return [];
          const kind = node.type === "footnoteRef" ? "footnote" : "endnote";
          const notes =
            kind === "footnote" ? document.package.footnotes : document.package.endnotes;
          if (!notes?.some(({ id }) => id === node.id)) return [];
          return [
            {
              at: {
                story: OP_STORIES.MAIN,
                blockId: referenceParagraph.paraId,
                offset: before.offset,
              },
              story: { kind, id: node.id } satisfies NoteStory,
            },
          ];
        }),
      );
      const selected = references.at(choose(references.length));
      return selected ? { type: DOCUMENT_OP_TYPES.REMOVE_NOTE, ...selected } : undefined;
    }
    case "createNumberingInstance": {
      let numId = 1;
      while (document.package.numbering?.nums.some((num) => num.numId === numId)) numId += 1;
      let abstractNumId = 1;
      while (
        document.package.numbering?.abstractNums.some((num) => num.abstractNumId === abstractNumId)
      )
        abstractNumId += 1;
      return {
        type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
        num: { numId, abstractNumId },
        abstractNum: { abstractNumId, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] },
      };
    }
    case "setSectionProps": {
      const sectionIndex = choose(sectionCount);
      const properties = sectionPropertiesAt(document, sectionIndex);
      if (!properties) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
        sectionIndex,
        patch: {
          marginTop: properties.marginTop === 720 ? 1440 : 720,
          evenAndOddHeaders: ([true, false, null] as const).at(choose(3)) ?? null,
        },
      };
    }
    case "insertText":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at,
        text: ["x", "é", "🧭"][choose(3)] ?? "x",
        runProps: INHERIT_RUN_PROPS,
        newIds,
        ...review,
      };
    case "insertContent":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at,
        slice: { content: freshParagraph.content, openStart: 0, openEnd: 0 },
        newIds,
        ...review,
      };
    case "deleteRange":
      return { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at, to: end, newIds, ...review };
    case "setRunProps":
      return {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        from: at,
        to: end,
        patch: { bold: choose(2) === 0 },
        newIds,
        ...review,
      };
    case "setParagraphProps":
      return {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: OP_STORIES.MAIN,
        blockId,
        patch: { alignment: paragraph.formatting?.alignment === "center" ? "left" : "center" },
        ...review,
      };
    case "splitBlock":
      return { type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, at, newBlockId, newIds, ...review };
    case "joinBlocks": {
      const next = blockListAt(document.package.document.content, location.list).at(
        location.index + 1,
      );
      if (next?.type !== "paragraph" || next.paraId === undefined) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId,
        nextBlockId: next.paraId,
        newIds,
        ...review,
      };
    }
    case "insertBlocks":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        story: OP_STORIES.MAIN,
        at: { type: "before", blockId },
        blocks: [freshParagraph],
        newIds,
        ...review,
      };
    case "deleteBlocks":
      return {
        type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
        story: OP_STORIES.MAIN,
        blockIds: [blockId],
        newIds,
        ...review,
      };
    case "insertTable":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story: OP_STORIES.MAIN,
        at: { type: "before", blockId },
        table: {
          type: "table",
          columnWidths: [900, 900],
          rows: [
            {
              type: "tableRow",
              cells: [
                { type: "tableCell", formatting: { gridSpan: 2 }, content: [freshParagraph] },
              ],
            },
          ],
        },
      };
    case "deleteTable":
      if (!location.list.some((step) => step.kind === "tableCell")) return undefined;
      return { type: DOCUMENT_OP_TYPES.DELETE_TABLE, story: OP_STORIES.MAIN, blockId };
    case "setContainerBlocks": {
      const expected = blockListAt(document.package.document.content, location.list);
      const blocks = [...expected];
      blocks.splice(location.index, 0, {
        type: "table",
        columnWidths: [900, 900],
        rows: [
          {
            type: "tableRow",
            cells: [{ type: "tableCell", formatting: { gridSpan: 2 }, content: [freshParagraph] }],
          },
        ],
      });
      return {
        type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId,
        expected,
        blocks,
      };
    }
    case "insertRow": {
      const target = tableForParagraph(document, location);
      if (target === undefined) return undefined;
      const width = inferredGridWidth(target.table);
      return {
        type: DOCUMENT_OP_TYPES.INSERT_ROW,
        story: OP_STORIES.MAIN,
        blockId,
        at: 0,
        row: {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              ...(width > 1 ? { formatting: { gridSpan: width } } : {}),
              content: [freshParagraph],
            },
          ],
        },
        newIds,
        ...review,
      };
    }
    case "deleteRow":
      if (!location.list.some((step) => step.kind === "tableCell")) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.DELETE_ROW,
        story: OP_STORIES.MAIN,
        blockId,
        newIds,
        ...review,
      };
    case "insertColumn": {
      const target = tableForParagraph(document, location);
      const grid = target === undefined ? undefined : tableGrid(target.table, family);
      const width = target === undefined ? 1 : inferredGridWidth(target.table);
      const column = choose(width + 1);
      const count = grid?.isOk()
        ? grid.value.rows.filter((cells, row) => {
            const before = target?.table.rows.at(row)?.formatting?.gridBefore ?? 0;
            const rowEnd = width - (target?.table.rows.at(row)?.formatting?.gridAfter ?? 0);
            return (
              column >= before &&
              column <= rowEnd &&
              !cells.some((cell) => cell.start < column && column < cell.end)
            );
          }).length
        : 0;
      return {
        type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
        story: OP_STORIES.MAIN,
        blockId,
        column,
        width: 900 + choose(600),
        newBlockIds: freshBlockIds(count),
        newIds,
        ...review,
      };
    }
    case "deleteColumn": {
      const target = tableForParagraph(document, location);
      return {
        type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
        story: OP_STORIES.MAIN,
        blockId,
        column: choose(target === undefined ? 1 : inferredGridWidth(target.table)),
        newIds,
        ...review,
      };
    }
    case "mergeCells": {
      const target = tableForParagraph(document, location);
      const grid = target === undefined ? undefined : tableGrid(target.table, family);
      const rowIndex = target?.rowIndex ?? 0;
      const cells = grid?.isOk() ? grid.value.rows.at(rowIndex) : undefined;
      return {
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        story: OP_STORIES.MAIN,
        blockId,
        top: rowIndex,
        bottom: rowIndex + 1,
        left: cells?.at(0)?.start ?? 0,
        right: cells?.at(-1)?.end ?? 1,
        newBlockIds: [],
        newIds,
        ...review,
      };
    }
    case "splitCell": {
      const target = tableForParagraph(document, location);
      const count = (target?.cell.formatting?.gridSpan ?? 1) - 1;
      return {
        type: DOCUMENT_OP_TYPES.SPLIT_CELL,
        story: OP_STORIES.MAIN,
        blockId,
        newBlockIds: freshBlockIds(count),
        newIds,
        ...review,
      };
    }
    case "setTableGrid": {
      const target = tableForParagraph(document, location);
      const width = target === undefined ? 1 : inferredGridWidth(target.table);
      return {
        type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
        story: OP_STORIES.MAIN,
        blockId,
        columnWidths: Array.from({ length: Math.min(width, MAX_NEW_IDS) }, () => 900 + choose(600)),
        newIds,
        ...review,
      };
    }
    case "setCellProps": {
      const target = tableForParagraph(document, location);
      return {
        type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
        story: OP_STORIES.MAIN,
        blockId,
        patch: { fitText: target?.cell.formatting?.fitText !== true },
        newIds,
        ...review,
      };
    }
    case "setRowProps": {
      const target = tableForParagraph(document, location);
      return {
        type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
        story: OP_STORIES.MAIN,
        blockId,
        patch: { cantSplit: target?.row.formatting?.cantSplit !== true },
        newIds,
        ...review,
      };
    }
    case "setTableProps": {
      const target = tableForParagraph(document, location);
      return {
        type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
        story: OP_STORIES.MAIN,
        blockId,
        patch: { layout: target?.table.formatting?.layout === "fixed" ? "autofit" : "fixed" },
        newIds,
        ...review,
      };
    }
    case "resolveRevision":
      if (revisions.length === 0) return undefined;
      return {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: revisions,
        decision: choose(2) === 0 ? REVISION_DECISIONS.ACCEPT : REVISION_DECISIONS.REJECT,
      };
    default: {
      const unreachable: never = family;
      return unreachable;
    }
  }
};

export type OpSequenceStep = { before: Document; op: DocumentOp; edit: AppliedDocumentOp };
export type OpSequence = {
  original: Document;
  originalModel: unknown;
  originalXml: string;
  document: Document;
  steps: OpSequenceStep[];
  inverse: DocumentOp[];
  mutations: string[];
  refusals: string[];
};

/** A schedule covers every family; the file seed varies positions, modes and order. */
export const generateOpSequence = (document: Document, seed: number): OpSequence => {
  const original = cloneDocumentWithParagraphPropertySources(document);
  const originalModel = exactOpModel(original);
  const originalXml = serializeOpDocument(original);
  const choose = randomFor(seed);
  const shift = choose(OP_SEQUENCE_FAMILIES.length);
  let current = original;
  const steps: OpSequenceStep[] = [];
  const revisions: number[] = [];
  const mutations: string[] = [];
  const refusals: string[] = [];
  const inverseGroups: (readonly DocumentOp[])[] = [];
  for (let index = 0; index < OP_SEQUENCE_LENGTH; index += 1) {
    const family = OP_SEQUENCE_FAMILIES.at((index + shift) % OP_SEQUENCE_FAMILIES.length);
    if (family === undefined)
      panic("The generated operation schedule must cover its declared family.");
    const op = candidate({
      document: current,
      choose,
      family,
      revisions,
      tracked: choose(2) === 0,
    });
    if (op === undefined) continue;
    const before = current;
    const snapshot = exactOpModel(before);
    const applied = applyDocumentOp(before, op);
    if (!isDeepStrictEqual(snapshot, exactOpModel(before))) mutations.push(op.type);
    if (applied.isErr()) {
      refusals.push(`${op.type}:${applied.error.reason}`);
      continue;
    }
    steps.push({ before, op, edit: applied.value });
    inverseGroups.unshift(applied.value.inverse);
    revisions.push(...applied.value.revisions);
    current = applied.value.document;
  }
  return {
    original,
    originalModel,
    originalXml,
    document: current,
    steps,
    inverse: inverseGroups.flat(),
    mutations,
    refusals,
  };
};
