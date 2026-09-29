/**
 * The requested-outcome oracle: what an applied operation must have done,
 * worked out from the document before it and the request alone, never from
 * folio's apply code. A receipt that says `applied` is checked against it
 * after a save and a reopen: in `direct` mode as saved, in `tracked-changes`
 * mode with every change accepted (and rejecting every change must give the
 * document before back). A refused operation must leave the document as it
 * was.
 *
 * The model keeps block texts and requested properties in reading order, plus
 * an independent cell grid and main-story links on unchanged paragraphs. A batch
 * resolves every operation against the document as it was read (see the
 * batch-claims contract), so a batch's expectations compose in order over
 * the same pre-state, and an operation the engine refused contributes
 * nothing. An operation the model cannot predict (a type marked `null` in
 * `EXPECTATIONS`, or a target it cannot place) leaves that batch's text
 * uncompared; supported geometry and links are still checked.
 * FOLIO_ORACLE_GAPS=1 prints each gap.
 */

import assert from "node:assert/strict";

import {
  FOLIO_DOCUMENT_OPERATION_TYPES,
  type FolioDocumentOperationType,
  type FolioDocumentStoryHandle,
  inspectDocumentStylesFromDocx,
} from "@stll/folio-core/server";

import { recordFeatureHit, recordHit, type StepKind } from "./coverage.ts";
import { operationSelection, targetFeatureSignature } from "./feature-coverage.ts";
import { openReviewer, toArrayBuffer } from "./documents.ts";
import { captureLinks, comparePreservedLinks, type LinkSnapshot } from "./link-oracle.ts";
import { coreBatch, type Mode, type Operation } from "./operations.ts";
import {
  applyTableOperation,
  compareTableGeometry,
  modelFromRows,
  UnsupportedTableExpectation,
  type TableModel,
} from "./table-oracle.ts";
import {
  blocksOfStory,
  type Feature,
  type FeatureIndex,
  featureIndex,
  storyKindOf,
  type TargetBlock,
  touchesSurrogate,
} from "./targets.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

type TableLocation = {
  outerTableIndex: number;
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
  gridColumnIndex: number;
  columnSpan: number;
  rowSpan: number;
};

type Run = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  directFormatting?: Record<string, unknown>;
};

/** One block as a reader sees it, with the fields the oracle compares. */
export type Row = {
  id: string;
  text: string;
  kind: string;
  styleId?: string;
  headingLevel?: number;
  listLevel?: number;
  listReference?: { numId: number; level: number };
  displayLabel?: string;
  directAlignment?: string;
  directSpacing?: Record<string, unknown>;
  table?: TableLocation;
  previewRuns?: readonly Run[];
};

/** The fields every block keeps unless an operation asked to change them. */
const KEPT_FIELDS = ["kind", "styleId", "headingLevel", "listLevel"] as const;

/** A field the request leaves to the document (a style's own outline or numbering). */
const ANY = Symbol("any");
type Fields = { [Key in (typeof KEPT_FIELDS)[number] | "directAlignment"]?: Row[Key] | typeof ANY };

type Story = FolioDocumentStoryHandle;
const MAIN: Story = { type: "main" };

/** The blocks of `story` (the body by default) a reader lists. */
export const rowsOf = (reviewer: Reviewer, story: Story = MAIN): Row[] =>
  blocksOfStory(reviewer, story) as unknown as Row[];

const save = async (reviewer: Reviewer): Promise<Uint8Array> =>
  new Uint8Array(await reviewer.toBuffer());

/**
 * `bytes` opened, every change resolved one way, saved and reopened: the
 * blocks a reader is left with, and the saved package.
 */
export const resolvedState = async (
  bytes: Uint8Array,
  resolution: "accept" | "reject",
  story: Story = MAIN,
): Promise<{ rows: Row[]; bytes: Uint8Array; comments: Comment[]; links: LinkSnapshot }> => {
  const reviewer = await openReviewer(bytes);
  if (resolution === "accept") reviewer.acceptAll();
  else reviewer.rejectAll();
  const saved = await save(reviewer);
  const reopened = await openReviewer(saved);
  return {
    rows: rowsOf(reopened, story),
    bytes: saved,
    comments: commentsOf(reopened),
    links: captureLinks(reopened),
  };
};

type Comment = { id: number; text: string; anchor: string; blockId: string | null };
const commentsOf = (reviewer: Reviewer): Comment[] =>
  reviewer.getComments().map((comment) => ({
    id: comment.id,
    text: comment.text,
    anchor: comment.anchoredText ?? "",
    blockId: comment.blockId ?? null,
  }));

// ---------------------------------------------------------------------------
// The model
// ---------------------------------------------------------------------------

type TextEdit = { start: number; end: number; replace: string };

/** A block of the pre-state, or one an operation adds, with what the batch did to it. */
type ModelRow = {
  /** The pre-state block; absent for a block an operation adds. */
  pre?: Row;
  text: string;
  fields: Fields;
  /** Extra checks on the result's block: a message when it fails. */
  checks: ((row: Row) => string | null)[];
  /** Text edits in the pre-state text's coordinates. */
  edits: TextEdit[];
  /** Split points in the pre-state text's coordinates, with what the break consumes. */
  splits: { offset: number; consumed: number; second?: Fields }[];
  mergeSeparator?: string;
  removed: boolean;
  /**
   * The block ends in a paragraph mark pending deletion: a reader lists it
   * and the block after it apart, accepting joins them. What an operation
   * on either does after the join is not modelled yet.
   */
  pendingJoin?: boolean;
  /**
   * A whole paragraph pending deletion, which a reader lists as a blank
   * block: accepting removes it, mark and all, unless an operation wrote
   * into it.
   */
  pendingDeletion?: boolean;
  before: ModelRow[];
  after: ModelRow[];
  /** Formatting that must hold over `[start, end)` of the result's text. */
  formats: { start: number; end: number; property: "bold" | "italic" | "underline" }[];
};

export type Model = {
  rows: ModelRow[];
  mode: Mode;
  tables: TableModel;
  tableGaps: string[];
  /** Existing paragraphs reveal style-derived kind, levels and effective run formatting. */
  styleExamples: Map<string, Row>;
  /** Text of the live blocks an operation can name, before the batch. */
  liveTextById: ReadonlyMap<string, string>;
  /** Live blocks explicitly deleted by applied operations, including pending joins. */
  deletedBlockIds: Set<string>;
  /** Paragraph style ids the saved package must define. */
  styles: Set<string>;
  /** Comments that must exist, with their anchored text when it is known. */
  comments: { text: string; anchor?: () => string | undefined }[];
  /** Operations the oracle has no expectation for; the check notes them. */
  unmodelled: string[];
  /** Blocks that are paragraphs of a text box drawn in the block before them. */
  inTextBox?: ReadonlySet<string>;
};

const fieldsOf = (row: Row): Fields => {
  const fields: Fields = {};
  for (const key of KEPT_FIELDS) (fields as Record<string, unknown>)[key] = row[key];
  return fields;
};

const modelRow = (text: string, fields: Fields = {}, pre?: Row): ModelRow => ({
  ...(pre ? { pre } : {}),
  text,
  fields,
  checks: [],
  edits: [],
  splits: [],
  removed: false,
  before: [],
  after: [],
  formats: [],
});

/**
 * The model of `rows`. `live`, the reviewer's own blocks, adds the blocks a
 * reader still lists but accepting removes (a whole paragraph pending
 * deletion reads as a blank block): an operation may name one, and what it
 * asks for is then placed where the reader saw that block.
 */
export const modelOf = (rows: readonly Row[], live: readonly Row[] = rows): Model => {
  const model: Model = {
    rows: rows.map((row) => modelRow(row.text, fieldsOf(row), row)),
    mode: "direct",
    tables: modelFromRows(rows),
    tableGaps: [],
    styleExamples: new Map(
      rows.flatMap((row) =>
        row.styleId !== undefined && row.text.length > 0 ? [[row.styleId, row] as const] : [],
      ),
    ),
    liveTextById: new Map(live.map(({ id, text }) => [id, text])),
    deletedBlockIds: new Set(),
    styles: new Set(),
    comments: [],
    unmodelled: [],
  };
  const known = new Set(rows.map((row) => row.id));
  const indexOf = (id: string | undefined): number =>
    id === undefined ? -1 : model.rows.findIndex((candidate) => candidate.pre?.id === id);
  let previous: string | undefined;
  for (const row of live) {
    if (known.has(row.id)) {
      previous = row.id;
      continue;
    }
    known.add(row.id);
    const ghost = modelRow("", {}, row);
    if (row.text.length > 0) {
      // Joined into the block before it once accepted.
      const head = model.rows[indexOf(previous)];
      if (head) head.pendingJoin = true;
      ghost.pendingJoin = true;
      ghost.removed = true;
      model.rows.push(ghost);
      continue;
    }
    ghost.pendingDeletion = true;
    model.rows.splice(indexOf(previous) + 1, 0, ghost);
    previous = row.id;
  }
  return model;
};

/** Why an operation could not be modelled; the check reports it as an oracle gap. */
class Unmodelled extends Error {}

const target = (model: Model, blockId: unknown, { adjacent = false } = {}): ModelRow => {
  const row = model.rows.find((candidate) => candidate.pre?.id === blockId);
  if (!row) throw new Unmodelled(`block ${String(blockId)} is not in the resolved pre-state`);
  if (row.pendingJoin) throw new Unmodelled(`block ${String(blockId)} has a pending join`);
  // An insertion only places a block beside its anchor, which may go.
  if (row.removed && !adjacent) {
    throw new Error(
      `applied an operation on block ${String(blockId)}, which an earlier one of the batch removed`,
    );
  }
  return row;
};

const tableOf = (row: ModelRow): TableLocation => {
  const table = row.pre?.table;
  if (!table) throw new Unmodelled("the anchor is not in a table");
  return table;
};

const tableRows = (model: Model, tableIndex: number): ModelRow[] =>
  model.rows.filter((row) => row.pre?.table?.tableIndex === tableIndex);

const hasMergedCells = (model: Model, tableIndex: number): boolean =>
  tableRows(model, tableIndex).some(
    (row) => (row.pre?.table?.rowSpan ?? 1) > 1 || (row.pre?.table?.columnSpan ?? 1) > 1,
  );

const paragraphsOf = (text: string): string[] =>
  text.split(/\r\n|\r|\n/u).filter((line) => line.length > 0);

/**
 * The fields a style id asks for on `pre` (absent: a new block). A style may
 * number its paragraphs, and a direct outline level outlives a restyle, so
 * the kind and levels are the request's only for a plain paragraph made a
 * `HeadingN`.
 */
const styleFields = (model: Model, styleId: string | null, pre: Row | undefined): Fields => {
  if (styleId !== null) model.styles.add(styleId);
  const example = styleId === null ? undefined : model.styleExamples.get(styleId);
  const heading =
    styleId === null ? undefined : /^Heading(?<level>[1-9])$/u.exec(styleId)?.groups?.["level"];
  const plain = pre !== undefined && pre.kind !== "heading";
  let headingLevel: Fields["headingLevel"] = ANY;
  if (example && pre?.headingLevel === undefined) headingLevel = example.headingLevel;
  else if (heading && plain) headingLevel = Number(heading);
  return {
    styleId: styleId ?? undefined,
    kind: example?.kind ?? (heading ? "heading" : ANY),
    headingLevel,
    listLevel: example && pre?.listReference === undefined ? example.listLevel : ANY,
  };
};

const isBullet = (label: string | undefined): boolean =>
  label !== undefined && label.length > 0 && !/[\p{L}\p{N}]/u.test(label);

type ParagraphRequest = Record<string, unknown>;

/** Fields and checks for the paragraph properties an operation sets on `pre` (absent: a new block). */
const paragraphRequest = (
  model: Model,
  request: ParagraphRequest,
  pre: Row | undefined,
): { fields: Fields; checks: ModelRow["checks"] } => {
  const fields: Fields = {};
  const checks: ModelRow["checks"] = [];
  if ("styleId" in request) {
    const styleId = request["styleId"] as string | null;
    Object.assign(fields, styleFields(model, styleId, pre));
    const example = styleId === null ? undefined : model.styleExamples.get(styleId);
    if (
      model.mode !== "suggested" &&
      example &&
      !example.previewRuns?.some((run) => run.directFormatting !== undefined) &&
      !pre?.previewRuns?.some((run) => run.directFormatting !== undefined)
    ) {
      for (const property of ["bold", "italic", "underline"] as const) {
        const flags = formattingAt(example, property) ?? Array(example.text.length).fill(false);
        if (!flags.every((flag) => flag === flags[0])) continue;
        const expected = flags[0] === true;
        checks.push((row) => {
          const actual = formattingAt(row, property) ?? Array(row.text.length).fill(false);
          return actual.every((flag) => flag === expected)
            ? null
            : `effective ${property} differs from style ${styleId}`;
        });
      }
    }
  }
  if ("alignment" in request) {
    fields.directAlignment = (request["alignment"] as string | null) ?? undefined;
  }
  if ("spacing" in request) {
    const spacing = request["spacing"] as Record<string, unknown> | null;
    checks.push((row) => {
      if (spacing === null) {
        return row.directSpacing === undefined
          ? null
          : `direct spacing ${JSON.stringify(row.directSpacing)} was not cleared`;
      }
      const missing = Object.entries(spacing).filter(
        ([key, value]) => row.directSpacing?.[key] !== value,
      );
      return missing.length === 0
        ? null
        : `direct spacing ${JSON.stringify(row.directSpacing)} lacks ${JSON.stringify(spacing)}`;
    });
  }
  if ("numbering" in request) {
    const numbering = request["numbering"] as Record<string, unknown> | null;
    const preKind = pre?.kind;
    if (numbering === null) {
      fields.listLevel = undefined;
      checks.push((row) =>
        row.listReference === undefined
          ? null
          : `still numbered ${JSON.stringify(row.listReference)}`,
      );
      if (preKind === "listItem") fields.kind = "paragraph";
    } else if (numbering["start"] === "new") {
      const kind = numbering["kind"];
      fields.listLevel = typeof numbering["level"] === "number" ? numbering["level"] : 0;
      if (pre && preKind !== "heading") fields.kind = "listItem";
      const used = new Set(
        model.rows.flatMap((row) => (row.pre?.listReference ? [row.pre.listReference.numId] : [])),
      );
      checks.push((row) => {
        if (!row.listReference) return "not numbered";
        if (used.has(row.listReference.numId))
          return `joined list ${row.listReference.numId} instead of starting one`;
        if (kind === "bullet" && !isBullet(row.displayLabel))
          return `a new bullet list reads "${row.displayLabel}"`;
        if (kind === "numbered" && isBullet(row.displayLabel))
          return `a new numbered list reads "${row.displayLabel}"`;
        return null;
      });
    } else {
      const reference = { numId: numbering["numId"], level: numbering["level"] };
      fields.listLevel = reference.level as number;
      if (pre && preKind !== "heading") fields.kind = "listItem";
      checks.push((row) =>
        row.listReference?.numId === reference.numId && row.listReference?.level === reference.level
          ? null
          : `numbered ${JSON.stringify(row.listReference)}, not ${JSON.stringify(reference)}`,
      );
    }
  }
  if (typeof request["listLevel"] === "number" && !("numbering" in request)) {
    fields.listLevel = request["listLevel"];
  }
  return { fields, checks };
};

/**
 * A comment the operation adds. `anchor` reads its anchored text once the
 * whole batch is modelled, or nothing when the batch changed that text.
 */
const addComment = (
  model: Model,
  operation: Operation,
  anchor?: () => string | undefined,
): void => {
  const comment = operation["comment"] as { text?: string } | undefined;
  if (typeof comment?.text === "string") {
    model.comments.push({ text: comment.text, ...(anchor === undefined ? {} : { anchor }) });
  }
};

/** Whether the row's text is still the pre-state's, so offsets into it hold. */
const untouchedText = (row: ModelRow): boolean =>
  row.edits.length === 0 && row.splits.length === 0 && row.mergeSeparator === undefined;

type Expect = (model: Model, operation: Operation) => void;

/**
 * A text box's paragraphs belong to the paragraph it is drawn in, which a
 * reader lists just before them: a block inserted after that paragraph
 * follows them.
 */
const pastTextBoxes = (model: Model, row: ModelRow): ModelRow => {
  const boxed = (candidate: ModelRow | undefined) =>
    candidate?.pre !== undefined && model.inTextBox?.has(candidate.pre.id) === true;
  if (boxed(row)) return row;
  let index = model.rows.indexOf(row);
  while (boxed(model.rows[index + 1])) index += 1;
  return model.rows[index] ?? row;
};

/**
 * Where a block inserted next to `row` goes: next to the row itself, or, for
 * a row inside a table, next to the outermost table (the documented rule:
 * an insertion anchored in a cell lands beside the table, not in the cell).
 */
const insertionAnchor = (model: Model, row: ModelRow, position: "before" | "after"): ModelRow => {
  if (position === "after") row = pastTextBoxes(model, row);
  const table = row.pre?.table;
  if (!table) return row;
  const rows = model.rows.filter(
    (candidate) => candidate.pre?.table?.outerTableIndex === table.outerTableIndex,
  );
  return (position === "before" ? rows[0] : rows.at(-1)) ?? row;
};

const insertBlock =
  (position: "before" | "after"): Expect =>
  (model, operation) => {
    const anchor = insertionAnchor(
      model,
      target(model, operation["blockId"], { adjacent: true }),
      position,
    );
    const { fields, checks } = paragraphRequest(model, operation, undefined);
    const scope = operation["formattingScope"] ?? "firstParagraph";
    const texts =
      operation["lineBreakMode"] === "inline"
        ? [String(operation["text"])]
        : paragraphsOf(String(operation["text"]));
    // `""` inserts a blank paragraph, which no reader lists.
    const rows = texts.map((text, index) => {
      const row = modelRow(text);
      if (index === 0 || scope === "allParagraphs") {
        row.fields = fields;
        row.checks = checks;
      }
      return row;
    });
    gapOf(model, anchor, position).push(...rows);
    addComment(model, operation);
  };

/**
 * The blocks inserted between `row` and its neighbour on `position`'s side.
 * An insertion after one block and one before the next land in one gap, in
 * batch order.
 */
const gapOf = (model: Model, row: ModelRow, position: "before" | "after"): ModelRow[] => {
  if (position === "after") return row.after;
  const previous = model.rows[model.rows.indexOf(row) - 1];
  return previous && !previous.pendingJoin ? previous.after : row.before;
};

const rangeOf = (operation: Operation): { blockId: string; start: number; end: number } => {
  const range = operation["range"] as { blockId: string; startOffset: number; endOffset: number };
  return { blockId: range.blockId, start: range.startOffset, end: range.endOffset };
};

/**
 * One block-text expectation per public operation type. `null` leaves text
 * unmodelled for that type; the separate table model still checks its grid.
 * A new operation type fails the completeness scenario until it is classified.
 */
export const EXPECTATIONS = {
  replaceInBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const find = String(operation["find"]);
    const text = row.pre?.text ?? row.text;
    const start = text.indexOf(find);
    if (start === -1) throw new Error(`applied, but "${find}" is not in "${text}"`);
    row.edits.push({ start, end: start + find.length, replace: String(operation["replace"]) });
    addComment(model, operation);
  },
  replaceRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    target(model, blockId).edits.push({ start, end, replace: String(operation["replace"]) });
    addComment(model, operation);
  },
  formatRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    const row = target(model, blockId);
    const formatting = operation["formatting"] as Record<string, unknown>;
    for (const property of ["bold", "italic", "underline"] as const) {
      if (formatting[property] === true) row.formats.push({ start, end, property });
    }
  },
  commentOnRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    const row = target(model, blockId);
    addComment(model, operation, () =>
      untouchedText(row) && !row.removed ? row.text.slice(start, end) : undefined,
    );
  },
  commentOnBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const quote = operation["quote"];
    const anchored = typeof quote === "string" ? quote : undefined;
    addComment(model, operation, () =>
      untouchedText(row) && !row.removed ? (anchored ?? row.text) : undefined,
    );
  },
  insertAfterBlock: insertBlock("after"),
  insertBeforeBlock: insertBlock("before"),
  replaceBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    row.edits.push({ start: 0, end: row.text.length, replace: String(operation["text"]) });
    if ("styleId" in operation) {
      const { fields, checks } = paragraphRequest(
        model,
        { styleId: operation["styleId"] },
        row.pre,
      );
      Object.assign(row.fields, fields);
      row.checks.push(...checks);
    }
    addComment(model, operation);
  },
  deleteBlock: (model, operation) => {
    target(model, operation["blockId"]).removed = true;
  },
  splitBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const offset = Number(operation["offset"]);
    const separator = typeof operation["separator"] === "string" ? operation["separator"] : "";
    const first = operation["firstParagraphProperties"] as ParagraphRequest | undefined;
    const second = operation["secondParagraphProperties"] as ParagraphRequest | undefined;
    if (first) Object.assign(row.fields, paragraphRequest(model, first, row.pre).fields);
    row.splits.push({
      offset,
      consumed: separator.length,
      ...(second
        ? { second: { ...row.fields, ...paragraphRequest(model, second, row.pre).fields } }
        : {}),
    });
  },
  mergeBlockWithNext: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const next = model.rows[model.rows.indexOf(row) + 1];
    if (row.pendingDeletion || next?.pendingDeletion) {
      // Whether the join keeps that paragraph's pending deletion is not modelled yet.
      throw new Unmodelled("a join with a paragraph pending deletion");
    }
    row.mergeSeparator = typeof operation["separator"] === "string" ? operation["separator"] : "";
    const merged = operation["mergedParagraphProperties"] as ParagraphRequest | undefined;
    if (merged) Object.assign(row.fields, paragraphRequest(model, merged, row.pre).fields);
  },
  setBlockParagraphProperties: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const { fields, checks } = paragraphRequest(
      model,
      operation["properties"] as ParagraphRequest,
      row.pre,
    );
    Object.assign(row.fields, fields);
    row.checks.push(...checks);
  },
  insertTable: (model, operation) => {
    const position = operation["position"] === "before" ? "before" : "after";
    const named = target(model, operation["blockId"], { adjacent: true });
    const anchor = position === "after" ? pastTextBoxes(model, named) : named;
    const rows = operation["rows"] as string[][];
    // A cell's line break starts a paragraph there; a blank one stays blank.
    const texts = rows.flatMap((cells) => cells.flatMap((cell) => cell.split(/\r\n|\r|\n/u)));
    gapOf(model, anchor, position).push(...texts.map((text) => modelRow(text)));
  },
  deleteTable: (model, operation) => {
    const { tableIndex } = tableOf(target(model, operation["blockId"]));
    for (const row of tableRows(model, tableIndex)) row.removed = true;
  },
  insertTableRow: (model, operation) => {
    // A row may go in beside one the batch deletes.
    const anchor = tableOf(target(model, operation["blockId"], { adjacent: true }));
    const inRow = tableRows(model, anchor.tableIndex).filter(
      (row) => row.pre?.table?.rowIndex === anchor.rowIndex,
    );
    const texts = ((operation["cellTexts"] as string[] | undefined) ?? []).flatMap((text) =>
      text.split(/\r\n|\r|\n/u),
    );
    const added = texts.map((text) => modelRow(text));
    // Every supplied text lands in one new row, one cell after another.
    const cells: ModelRow[] = added.filter((row) => row.text.length > 0);
    for (const cell of cells) {
      cell.checks.push((row) => (row.table ? null : "is not in a table"));
    }
    const position = operation["position"] === "before" ? "before" : "after";
    const edge = position === "before" ? inRow[0] : inRow.at(-1);
    if (!edge) throw new Unmodelled("the anchor row has no blocks");
    edge[position].push(...added);
  },
  deleteTableRow: (model, operation) => {
    const anchor = tableOf(target(model, operation["blockId"]));
    if (hasMergedCells(model, anchor.tableIndex)) throw new Unmodelled("merged cells");
    for (const row of tableRows(model, anchor.tableIndex)) {
      if (row.pre?.table?.rowIndex === anchor.rowIndex) row.removed = true;
    }
  },
  insertTableColumn: (model, operation) => {
    const anchor = tableOf(target(model, operation["blockId"]));
    if (hasMergedCells(model, anchor.tableIndex)) throw new Unmodelled("merged cells");
    const texts = (operation["cellTexts"] as string[] | undefined) ?? [];
    const after = operation["position"] !== "before";
    const rows = tableRows(model, anchor.tableIndex);
    const rowIndexes = [...new Set(rows.map((row) => row.pre?.table?.rowIndex ?? 0))].sort(
      (left, right) => left - right,
    );
    texts.forEach((text, index) => {
      const rowIndex = rowIndexes[index];
      if (rowIndex === undefined) throw new Error("more cell texts than rows were applied");
      const inRow = rows.filter((row) => row.pre?.table?.rowIndex === rowIndex);
      // The new cell follows the anchor's column (or precedes it) in each row.
      const column = anchor.gridColumnIndex;
      const cell = modelRow(text);
      const lastBefore = inRow.findLast((row) =>
        after
          ? (row.pre?.table?.gridColumnIndex ?? 0) <= column
          : (row.pre?.table?.gridColumnIndex ?? 0) < column,
      );
      if (lastBefore) lastBefore.after.push(cell);
      else inRow[0]?.before.push(cell);
    });
  },
  deleteTableColumn: (model, operation) => {
    const anchor = tableOf(target(model, operation["blockId"]));
    if (hasMergedCells(model, anchor.tableIndex)) throw new Unmodelled("merged cells");
    for (const row of tableRows(model, anchor.tableIndex)) {
      if (row.pre?.table?.gridColumnIndex === anchor.gridColumnIndex) row.removed = true;
    }
  },
  insertSignatureTable: (model, operation) => {
    const position = operation["position"] === "before" ? "before" : "after";
    const anchor = insertionAnchor(
      model,
      target(model, operation["blockId"], { adjacent: true }),
      position,
    );
    const parties = operation["parties"] as {
      name: string;
      signatory?: string;
      title?: string;
    }[];
    const texts = parties.flatMap(({ name, signatory, title }) => {
      const lines = [name, "", "", "_".repeat(28)];
      if (signatory) lines.push(signatory);
      if (title) lines.push(title);
      return lines;
    });
    gapOf(model, anchor, position).push(...texts.map((text) => modelRow(text)));
    addComment(model, operation);
  },
  mergeTableCells: null,
  splitTableCell: null,
} as const satisfies Record<FolioDocumentOperationType, Expect | null>;

/** The public operation types, for the completeness scenario. */
export const OPERATION_TYPES: readonly string[] = FOLIO_DOCUMENT_OPERATION_TYPES;

/** Apply one applied operation's expectation to `model`. */
export const expectOperation = (model: Model, operation: Operation): void => {
  if (operation.type === "deleteBlock" && typeof operation["blockId"] === "string") {
    model.deletedBlockIds.add(operation["blockId"]);
  }
  if (
    model.mode === "suggested" &&
    ["deleteTable", "deleteTableRow", "deleteTableColumn"].includes(operation.type)
  ) {
    model.tableGaps.push(`${operation.type}: a pending deletion remains in the live grid`);
  } else if (model.tableGaps.length === 0) {
    try {
      applyTableOperation(model.tables, operation);
    } catch (error) {
      if (error instanceof UnsupportedTableExpectation) {
        model.tableGaps.push(`${operation.type}: ${error.message}`);
      } else {
        throw error;
      }
    }
  }
  const expectation = (EXPECTATIONS as Record<string, Expect | null | undefined>)[operation.type];
  if (expectation === undefined) {
    throw new Error(`the oracle has no entry for operation type ${operation.type}`);
  }
  if (expectation === null) {
    model.unmodelled.push(`${operation.type}: no block-text expectation yet`);
    return;
  }
  try {
    expectation(model, operation);
  } catch (error) {
    if (error instanceof Unmodelled) {
      model.unmodelled.push(`${operation.type}: ${error.message}`);
      return;
    }
    throw error;
  }
};

// ---------------------------------------------------------------------------
// Materializing and comparing
// ---------------------------------------------------------------------------

type Expected = {
  text: string;
  fields: Fields;
  checks: ModelRow["checks"];
  formats: ModelRow["formats"];
  /** The pre-state block, while its text is unchanged: formatting outside a range keeps. */
  pre?: Row;
};

const applyEdits = (text: string, edits: readonly TextEdit[], splits: ModelRow["splits"]) => {
  // Everything in pre-state coordinates, applied from the end backwards.
  type Cut = { start: number; end: number; replace: string; split?: ModelRow["splits"][number] };
  const cuts: Cut[] = [
    ...edits.map((edit) => ({ ...edit })),
    ...splits.map((split) => ({
      start: split.offset,
      end: split.offset + split.consumed,
      replace: "",
      split,
    })),
  ].sort((left, right) => right.start - left.start || right.end - left.end);
  const segments: { text: string; split?: ModelRow["splits"][number] }[] = [];
  const tail = text;
  let current = "";
  let cursor = text.length;
  for (const cut of cuts) {
    if (cut.end > cursor) throw new Error(`two applied edits overlap in "${text}"`);
    current = cut.replace + tail.slice(cut.end, cursor) + current;
    cursor = cut.start;
    if (cut.split) {
      segments.unshift({ text: current, split: cut.split });
      current = "";
    }
  }
  current = tail.slice(0, cursor) + current;
  segments.unshift({ text: current });
  return segments;
};

const materialize = (model: Model): Expected[] => {
  const out: Expected[] = [];
  let joinNext: string | undefined;
  const push = (entry: Expected) => {
    if (joinNext !== undefined) {
      const previous = out.at(-1);
      if (previous) {
        previous.text = `${previous.text}${joinNext}${entry.text}`;
        // A join keeps the first paragraph's properties; formatting offsets no longer hold.
        previous.formats = [];
        joinNext = undefined;
        return;
      }
      joinNext = undefined;
    }
    out.push(entry);
  };
  const emit = (row: ModelRow) => {
    for (const inserted of row.before) emit(inserted);
    const gone = row.removed || (row.pendingDeletion === true && row.edits.length === 0);
    if (gone && joinNext !== undefined) {
      const previous = out.at(-1);
      if (previous) previous.text += joinNext;
      joinNext = undefined;
    }
    if (!gone) {
      const segments = applyEdits(row.text, row.edits, row.splits);
      const whole = untouchedText(row);
      segments.forEach((segment, index) =>
        push({
          text: segment.text,
          fields: index === 0 ? row.fields : (segment.split?.second ?? row.fields),
          checks: row.checks,
          formats: whole ? row.formats : [],
          ...(whole && row.pre ? { pre: row.pre } : {}),
        }),
      );
      if (row.mergeSeparator !== undefined) joinNext = row.mergeSeparator;
    }
    for (const inserted of row.after) emit(inserted);
  };
  for (const row of model.rows) emit(row);
  return out;
};

/** A reader lists no blank block; neither side is compared on one. */
const visible = <T extends { text: string }>(rows: readonly T[]): T[] =>
  rows.filter((row) => row.text.length > 0);

const formattingAt = (row: Row, property: "bold" | "italic" | "underline"): boolean[] | null => {
  const runs = row.previewRuns;
  if (!runs || runs.map((run) => run.text).join("") !== row.text) return null;
  return runs.flatMap((run) =>
    Array.from({ length: run.text.length }, () => run[property] === true),
  );
};

/** Compare `actual` with what `model` expects; every mismatch, as text. */
export const compareWithModel = (model: Model, actual: readonly Row[]): string[] => {
  const expected = visible(materialize(model));
  const got = visible(actual);
  const problems: string[] = [];
  const expectedTexts = expected.map((row) => row.text);
  const gotTexts = got.map((row) => row.text);
  if (JSON.stringify(expectedTexts) !== JSON.stringify(gotTexts)) {
    problems.push(
      `block texts differ:\n    expected ${JSON.stringify(expectedTexts)}\n    got      ${JSON.stringify(gotTexts)}`,
    );
    return problems;
  }
  expected.forEach((entry, index) => {
    const row = got[index] as Row;
    for (const [key, value] of Object.entries(entry.fields)) {
      if (value === ANY) continue;
      const actualValue = row[key as keyof Row];
      if (JSON.stringify(actualValue) !== JSON.stringify(value)) {
        problems.push(
          `"${row.text}": ${key} is ${JSON.stringify(actualValue)}, expected ${JSON.stringify(value)}`,
        );
      }
    }
    for (const check of entry.checks) {
      const problem = check(row);
      if (problem !== null) problems.push(`"${row.text}": ${problem}`);
    }
    for (const format of entry.formats) {
      const flags = formattingAt(row, format.property);
      if (flags === null) continue;
      const missing = flags.slice(format.start, format.end).some((flag) => !flag);
      if (missing) {
        problems.push(
          `"${row.text}": [${format.start}, ${format.end}) is not all ${format.property}`,
        );
      }
      // Outside the range, every character keeps what it had.
      const before = entry.pre && formattingAt(entry.pre, format.property);
      const inside = (offset: number) =>
        entry.formats.some(
          (other) =>
            other.property === format.property && offset >= other.start && offset < other.end,
        );
      const changed = before
        ? flags.findIndex((flag, offset) => !inside(offset) && flag !== before[offset])
        : -1;
      if (changed !== -1) {
        problems.push(
          `"${row.text}": ${format.property} changed at ${changed}, outside what was asked`,
        );
      }
    }
  });
  return problems;
};

type AnchorDisposition = "kept" | "removed" | "ambiguous";

/** Decide whether the model removed the anchor, accounting for repeated text. */
const anchorDisposition = (model: Model, entry: Comment): AnchorDisposition => {
  if (entry.blockId === null) return "kept";
  const row = model.rows.find((candidate) => candidate.pre?.id === entry.blockId);
  const explicitlyDeleted = model.deletedBlockIds.has(entry.blockId);
  if (explicitlyDeleted || (row?.removed && !row.pendingJoin)) {
    const liveText = model.liveTextById.get(entry.blockId);
    return entry.anchor === "" || (liveText !== undefined && liveText.includes(entry.anchor))
      ? "removed"
      : "kept";
  }
  // A pending join is already absent from the accepted pre-state; its live
  // block has not been removed by this batch.
  if (!row?.pre || row.removed || entry.anchor === "") return "kept";
  const starts: number[] = [];
  for (let start = 0; start <= row.pre.text.length - entry.anchor.length; start += 1) {
    if (row.pre.text.startsWith(entry.anchor, start)) starts.push(start);
  }
  if (starts.length === 0) return "kept";
  const removed = [
    ...row.edits
      .filter(({ replace }) => replace.length === 0)
      .map(({ start, end }) => ({ start, end })),
    ...row.splits.map(({ offset, consumed }) => ({ start: offset, end: offset + consumed })),
  ];
  const removedPlacements = starts.map((start) =>
    Array.from({ length: entry.anchor.length }, (_, index) => start + index).every((offset) =>
      removed.some((cut) => cut.start <= offset && offset < cut.end),
    ),
  );
  if (removedPlacements.every(Boolean)) return "removed";
  return removedPlacements.some(Boolean) ? "ambiguous" : "kept";
};

/** A full-block anchor has an exact replacement when one edit rewrites the block. */
const expectedKeptAnchor = (model: Model, entry: Comment): string | undefined => {
  const row = model.rows.find((candidate) => candidate.pre?.id === entry.blockId);
  if (!row?.pre || row.removed || row.pendingJoin || row.splits.length > 0) return undefined;
  if (entry.anchor === "" || entry.anchor !== row.pre.text || row.edits.length !== 1) {
    return undefined;
  }
  const edit = row.edits[0];
  if (!edit || edit.start !== 0 || edit.end !== row.pre.text.length || edit.replace === "") {
    return undefined;
  }
  return edit.replace.includes("\n") ? undefined : edit.replace;
};

/** Pending markup may interleave the old text and insertion in one range. */
const containsInOrder = (text: string, expected: string): boolean => {
  let offset = 0;
  for (const character of expected) {
    const index = text.indexOf(character, offset);
    if (index === -1) return false;
    offset = index + character.length;
  }
  return true;
};

const pendingAnchorMatches = (actual: string, old: string, replacement: string): boolean => {
  if (!containsInOrder(actual, old) || !containsInOrder(actual, replacement)) return false;
  // A retained character can serve both texts. With disjoint characters, the
  // deleted text must precede the insertion in the live tracked range.
  const replacementCharacters = new Set(replacement);
  return [...old].some((character) => replacementCharacters.has(character))
    ? true
    : containsInOrder(actual, old + replacement);
};

export const compareComments = (
  model: Model,
  before: readonly Comment[],
  after: readonly Comment[],
  liveBefore: readonly Comment[] = before,
  mode: Mode = "direct",
): string[] => {
  const problems: string[] = [];
  const remaining = [...after];
  const kept: Comment[] = [];
  const gone: Comment[] = [];
  for (const entry of before) {
    const live = liveBefore.find(({ id }) => id === entry.id);
    const disposition = mode === "suggested" ? "kept" : anchorDisposition(model, live ?? entry);
    if (disposition === "removed") gone.push(entry);
    else if (disposition === "kept") kept.push(entry);
  }
  for (const comment of kept) {
    const anchor = expectedKeptAnchor(model, comment);
    const index = remaining.findIndex(
      ({ id, text, anchor: actualAnchor }) =>
        id === comment.id &&
        text === comment.text &&
        (anchor === undefined ||
          (mode === "suggested"
            ? pendingAnchorMatches(actualAnchor, comment.anchor, anchor)
            : actualAnchor === anchor)),
    );
    if (index === -1) {
      problems.push(`no comment ${JSON.stringify(comment)} among ${JSON.stringify(after)}`);
    } else {
      remaining.splice(index, 1);
    }
  }
  const beforeIds = new Set(before.map(({ id }) => id));
  for (const comment of model.comments) {
    const anchor = comment.anchor?.();
    const index = remaining.findIndex(
      (candidate) =>
        !beforeIds.has(candidate.id) &&
        candidate.text === comment.text &&
        (anchor === undefined || candidate.anchor === anchor),
    );
    if (index === -1) {
      problems.push(
        `no comment ${JSON.stringify({ text: comment.text, anchor })} among ${JSON.stringify(after)}`,
      );
    } else {
      remaining.splice(index, 1);
    }
  }
  // Only what no expectation claimed can be the removed comment itself.
  for (const entry of gone) {
    if (remaining.some(({ id }) => id === entry.id)) {
      problems.push(`comment ${JSON.stringify(entry)} outlived the block it anchored`);
    }
  }
  return problems;
};

const compareStyles = async (model: Model, bytes: Uint8Array): Promise<string[]> => {
  if (model.styles.size === 0) return [];
  const catalog = await inspectDocumentStylesFromDocx(toArrayBuffer(bytes));
  const paragraphStyles = new Set(
    catalog.styles.filter((style) => style.type === "paragraph").map((style) => style.styleId),
  );
  return [...model.styles]
    .filter((styleId) => !paragraphStyles.has(styleId))
    .map((styleId) => `the saved package defines no paragraph style "${styleId}"`);
};

// ---------------------------------------------------------------------------
// The check around one apply
// ---------------------------------------------------------------------------

/**
 * The state an apply is checked against. `direct` and `tracked-changes`:
 * saved, reopened and every change accepted, so pending changes from before
 * resolve the same way on both sides. `suggested`: the live reviewer, whose
 * suggestions stay out of the package until accepted.
 */
export type Pre = {
  mode: Mode;
  live: string;
  rows: Row[];
  /** The reviewer's own blocks, when `rows` are another view of them. */
  liveRows: Row[];
  comments: Comment[];
  links: LinkSnapshot;
  liveComments: Comment[];
  rejected?: Row[];
  /** The story the operations target; the body by default. */
  story: Story;
  /** The session the step runs in, for the coverage ledger. */
  step: StepKind;
  /** What each block of the story is, for the coverage ledger. */
  targets: FeatureIndex;
};

export type CaptureOptions = { story?: Story; step?: StepKind };

const liveState = (reviewer: Reviewer, story: Story = MAIN): string =>
  JSON.stringify({
    blocks: reviewer.getContent(),
    ...(story.type === "main" ? {} : { story: rowsOf(reviewer, story) }),
    comments: reviewer.getComments().length,
  });

const project = (rows: readonly Row[]) =>
  visible(rows).map(({ text, kind, styleId, headingLevel, listLevel }) => ({
    text,
    kind,
    styleId,
    headingLevel,
    listLevel,
  }));

/** Per-character effective formatting survives a pending paragraph join. */
const inlineProject = (rows: readonly Row[]) =>
  visible(rows).flatMap((row) =>
    (row.previewRuns ?? [{ text: row.text }]).flatMap((run) =>
      Array.from(run.text, (character) => ({
        character,
        bold: run.bold === true,
        italic: run.italic === true,
        underline: run.underline === true,
      })),
    ),
  );

/** Authored run properties survive style inheritance changes on acceptance. */
const authoredInlineProject = (rows: readonly Row[]) =>
  visible(rows).flatMap((row) =>
    (row.previewRuns ?? [{ text: row.text }]).flatMap((run) =>
      Array.from(run.text, (character) => ({
        character,
        directFormatting: run.directFormatting,
      })),
    ),
  );

/** Reader-visible paragraph, table and effective run properties, without unstable ids. */
const fullProject = (rows: readonly Row[]) =>
  visible(rows).map((row) => ({
    text: row.text,
    kind: row.kind,
    styleId: row.styleId,
    headingLevel: row.headingLevel,
    listLevel: row.listLevel,
    directAlignment: row.directAlignment,
    directSpacing: row.directSpacing,
    table: row.table,
    runs: inlineProject([row]),
  }));

const suggestedProject = (rows: readonly Row[]) =>
  visible(rows).map((row) => ({
    text: row.text,
    kind: row.kind,
    styleId: row.styleId,
    headingLevel: row.headingLevel,
    listLevel: row.listLevel,
    directAlignment: row.directAlignment,
    directSpacing: row.directSpacing,
    table: row.table,
    runs: authoredInlineProject([row]),
  }));

/** Capture what the oracle needs before an apply. */
export const capture = async (
  reviewer: Reviewer,
  mode: Mode,
  { story = MAIN, step = "fresh" }: CaptureOptions = {},
): Promise<Pre> => {
  const live = liveState(reviewer, story);
  const liveComments = commentsOf(reviewer);
  const context = { story, step, targets: featureIndex(reviewer, story) };
  if (mode === "suggested") {
    const rows = rowsOf(reviewer, story);
    return {
      mode,
      live,
      rows,
      liveRows: rows,
      comments: liveComments,
      liveComments,
      links: captureLinks(reviewer),
      ...context,
    };
  }
  const bytes = await save(reviewer);
  const accepted = await resolvedState(bytes, "accept", story);
  return {
    mode,
    live,
    rows: accepted.rows,
    liveRows: rowsOf(reviewer, story),
    comments: accepted.comments,
    links: accepted.links,
    liveComments,
    ...(mode === "tracked-changes"
      ? { rejected: (await resolvedState(bytes, "reject", story)).rows }
      : {}),
    ...context,
  };
};

export type Outcome = {
  /** The operations the receipt says applied, in request order. */
  applied: readonly Operation[];
  /** Every operation asked for, applied or not (the ledger counts refusals); `applied` if absent. */
  attempted?: readonly Operation[];
};

const LEDGER_MODES: Record<Mode, string> = {
  direct: "direct",
  "tracked-changes": "tracked",
  suggested: "suggested",
};

/** The offsets an operation acts at, in its block's text, when it names any. */
const offsetsOf = (operation: Operation, text: string): [number, number] | null => {
  const range = operation["range"] as { startOffset?: unknown; endOffset?: unknown } | undefined;
  if (typeof range?.startOffset === "number" && typeof range.endOffset === "number") {
    return [range.startOffset, range.endOffset];
  }
  if (typeof operation["offset"] === "number") return [operation["offset"], operation["offset"]];
  const find = operation["find"] ?? operation["quote"];
  const at = typeof find === "string" && find.length > 0 ? text.indexOf(find) : -1;
  return at === -1 ? null : [at, at + (find as string).length];
};

/** Record every attempted operation of an outcome in the coverage ledger. */
const recordOutcome = (pre: Pre, outcome: Outcome): void => {
  const applied = new Set(outcome.applied);
  for (const operation of outcome.attempted ?? outcome.applied) {
    if (typeof operation !== "object" || operation === null) continue;
    const range = operation["range"] as { blockId?: unknown } | undefined;
    const blockId = operation["blockId"] ?? range?.blockId;
    const block = pre.liveRows.find((row) => row.id === blockId) as TargetBlock | undefined;
    const features = new Set<Feature | "none">(pre.targets.features.get(String(blockId)) ?? []);
    // A block with an astral character counts as a surrogate boundary only
    // where the operation's own offsets meet one.
    features.delete("surrogateBoundary");
    const offsets = block && offsetsOf(operation, block.text);
    if (block && offsets && touchesSurrogate(block.text, offsets[0], offsets[1])) {
      features.add("surrogateBoundary");
    }
    if (features.size === 0) features.add("none");
    const story = storyKindOf(pre.targets, block);
    for (const feature of targetFeatureSignature(block, features, story)) {
      recordFeatureHit({
        operation: String(operation.type),
        feature,
        selection: operationSelection(operation),
      });
    }
    for (const feature of features) {
      recordHit(
        {
          op: String(operation.type),
          story,
          mode: LEDGER_MODES[pre.mode],
          feature,
          step: pre.step,
        },
        applied.has(operation),
      );
    }
  }
};

/**
 * Check an apply's outcome against what was asked. Throws with every
 * mismatch; returns the gaps the oracle could not model.
 */
export const assertRequestedOutcome = async (
  reviewer: Reviewer,
  pre: Pre,
  outcome: Outcome,
  context: string,
): Promise<string[]> => {
  recordOutcome(pre, outcome);
  if (outcome.applied.length === 0) {
    assert.equal(
      liveState(reviewer, pre.story),
      pre.live,
      `${context}: nothing applied, but the document changed`,
    );
    return [];
  }
  const model = modelOf(pre.rows, pre.liveRows);
  model.inTextBox = pre.targets.inTextBox;
  model.mode = pre.mode;
  for (const operation of outcome.applied) expectOperation(model, operation);
  // An operation the oracle cannot model changes the document in a way it
  // cannot predict; the rest of the batch is not compared either.
  const predictable = model.unmodelled.length === 0;

  let rows: Row[];
  let comments: Comment[];
  let links: LinkSnapshot;
  let bytes: Uint8Array | null = null;
  if (pre.mode === "suggested") {
    rows = rowsOf(reviewer, pre.story);
    comments = commentsOf(reviewer);
    links = captureLinks(reviewer);
  } else {
    const saved = await save(reviewer);
    const accepted = await resolvedState(saved, "accept", pre.story);
    ({ rows, comments, bytes, links } = accepted);
    if (pre.rejected) {
      assert.deepEqual(
        project((await resolvedState(saved, "reject", pre.story)).rows),
        project(pre.rejected),
        `${context}: rejecting every change does not give the document before back`,
      );
    }
  }
  const problems = [
    ...(predictable ? compareWithModel(model, rows) : []),
    ...(model.tableGaps.length === 0 ? compareTableGeometry(model.tables, rows) : []),
    ...comparePreservedLinks({ before: pre.links, after: links, afterRows: rows }).problems,
    ...compareComments(model, pre.comments, comments, pre.liveComments, pre.mode),
    ...(bytes ? await compareStyles(model, bytes) : []),
  ];
  if (process.env["FOLIO_ORACLE_GAPS"] && (!predictable || model.tableGaps.length > 0)) {
    console.log(`oracle gap: ${context}: ${[...model.unmodelled, ...model.tableGaps].join("; ")}`);
  }
  if (problems.length > 0) {
    throw new Error(
      `${context}: the result is not what was asked (${outcome.applied
        .map((operation) => operation.type)
        .join(", ")}):\n  ${problems.join("\n  ")}`,
    );
  }
  return [...model.unmodelled, ...model.tableGaps];
};

/** What resolving every change must leave, and whether block boundaries are known. */
export type Resolution = { rows: Row[]; boundaries: boolean };

/** Apply `operations` as one core batch in `mode`, and check the result is what they asked. */
export const applyChecked = async (
  reviewer: Reviewer,
  operations: readonly Operation[],
  mode: Mode,
  context: string,
): Promise<{ applied: string[]; issues: string[] }> => {
  const pre = await capture(reviewer, mode);
  const batch = coreBatch(operations, mode);
  const result = reviewer.applyDocumentOperations(batch as never);
  const applied = new Set(result.applied.map(({ id }) => id));
  const issues = result.issues.map((issue) => `${issue.operationId}: ${issue.code}`);
  await assertRequestedOutcome(
    reviewer,
    pre,
    { applied: batch.operations.filter((operation) => applied.has(operation.id)) },
    `${context} ${JSON.stringify(operations)} (refused: ${issues.join(", ") || "none"})`,
  );
  return { applied: [...applied], issues };
};

/**
 * What bulk resolution must leave: the accepted (or rejected) view of the
 * saved package. The package excludes pending suggestions, which remain
 * staged in memory. A paragraph-mark deletion can join blocks on acceptance,
 * so the suggested live view may not predict their boundaries.
 */
export const captureResolution = async (
  reviewer: Reviewer,
  _mode: Mode,
  resolution: "accept" | "reject",
): Promise<Resolution> => ({
  rows: (await resolvedState(await save(reviewer), resolution)).rows,
  boundaries: true,
});

/** After `acceptAll()` / `rejectAll()`: the saved package reads `expected`. */
export const assertResolvedTo = async (
  reviewer: Reviewer,
  expected: Resolution,
  context: string,
): Promise<void> => {
  const rows = rowsOf(await openReviewer(await save(reviewer)));
  const message = `${context}: resolving every change does not leave what the reviewer showed`;
  if (expected.boundaries) {
    assert.deepEqual(fullProject(rows), fullProject(expected.rows), message);
  } else {
    assert.deepEqual(authoredInlineProject(rows), authoredInlineProject(expected.rows), message);
    if (
      JSON.stringify(visible(rows).map((row) => row.text)) ===
      JSON.stringify(visible(expected.rows).map((row) => row.text))
    ) {
      assert.deepEqual(suggestedProject(rows), suggestedProject(expected.rows), message);
    }
  }
};
