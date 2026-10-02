/**
 * Metamorphic relations over the seeded flows (support/fuzz.ts): two public
 * paths to what must be the same document, compared with each other rather
 * than with a predicted result. Readers and receipts can all agree on
 * something wrong; two independent paths rarely go wrong the same way.
 *
 * - `directTracked`: the flow's batches replayed on two fresh reviewers of
 *   the fixture, one in `direct` mode and one in `tracked-changes` mode; the
 *   tracked one with every change accepted reads as the direct one (both
 *   saved, reopened, and with the fixture's own changes accepted). The direct
 *   replay applies only what the tracked one applied; once direct refuses an
 *   operation tracked applied, the two documents part and the relation stops
 *   for the flow.
 * - `rejectAll`: the tracked replay, and a `suggested` one, with every change
 *   rejected read as the fixture with every change rejected.
 * - `saveIdempotent`: saving the reopened package gives the same package
 *   parts as the save it was opened from (`docProps/` is left out: it holds
 *   the save's own timestamps).
 * - `undo`: each batch applied to a copy of the document before it and
 *   undone through its receipt's undo handle leaves the copy as it was
 *   (blocks, comments and changes, ids and dates included).
 * - `batchSequential`: a batch leaves what its applied operations leave one
 *   at a time, each re-resolved against the document as it then stands, in
 *   the order the batch contract gives (packages/core/src/ai-edits/
 *   batchOverlap.test.ts): annotations first, then insertions and other
 *   non-column operations, then table-column edits in source-grid order
 *   (table and column descending, insertions before deletions at ties).
 * - `readerStability`: after every save, the live reviewer and the reopened
 *   one give the same `getContent()`, bridge snapshot, `toMarkdown`,
 *   `getChanges()` and `getComments()`.
 *
 * FOLIO_SCENARIO_RELATIONS picks the relations (`all` by default, `none`, or
 * a comma-separated list); FOLIO_SCENARIO_RELATIONS_DEPTH=full checks every
 * batch and every step, where the default samples the heavy relations
 * (`directTracked` / `rejectAll` at the flow's end, `batchSequential` on a
 * seeded share of batches). The flow's own random stream is never drawn
 * from, so a seed replays the same flow with relations on or off.
 * `relationSummary()` reports what ran, what was skipped and why, and how
 * many comparisons each tolerance (a difference an open finding causes,
 * keyed by its `FINDINGS` entry) let pass.
 */

import assert from "node:assert/strict";
import { inflateRawSync } from "node:zlib";

import { createReviewerBridge } from "@stll/folio-agents";
import { toMarkdown } from "@stll/folio-core/markdown";
import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentStoryHandle,
} from "@stll/folio-core/server";

import { openReviewer } from "./documents.ts";
import { projectContentPair, projectSnapshotIdentities } from "./identity.ts";
import type { Finding } from "./known-issues.ts";
import type { Mode } from "./operations.ts";
import { resolvedState, type Row } from "./oracle.ts";
import { createRandom, type Random } from "./random.ts";
import { RELATIONS, type Relation } from "./relation-contract.ts";
import { blocksOfStory } from "./targets.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type Batch = Parameters<Reviewer["applyDocumentOperations"]>[0];
type Result = ReturnType<Reviewer["applyDocumentOperations"]>;
type AnyOperation = { id: string; type: string } & Record<string, unknown>;
type Story = FolioDocumentStoryHandle;
const MAIN: Story = { type: "main" };

export { RELATIONS, type Relation } from "./relation-contract.ts";

const parseRelations = (value: string | undefined): ReadonlySet<Relation> => {
  if (value === undefined || value === "" || value === "all") return new Set(RELATIONS);
  if (value === "none") return new Set();
  const names = value
    .split(",")
    .map((name) => name.trim())
    .filter((name) => name.length > 0);
  for (const name of names) {
    if (!(RELATIONS as readonly string[]).includes(name)) {
      throw new Error(
        `FOLIO_SCENARIO_RELATIONS: unknown relation ${name} (known: ${RELATIONS.join(", ")})`,
      );
    }
  }
  return new Set(names as Relation[]);
};

export const ENABLED_RELATIONS = parseRelations(process.env["FOLIO_SCENARIO_RELATIONS"]);
const FULL = process.env["FOLIO_SCENARIO_RELATIONS_DEPTH"] === "full";
/** The share of batches `batchSequential` checks when not `full`. */
const SEQUENTIAL_SHARE = 0.35;

// ---------------------------------------------------------------------------
// The ledger
// ---------------------------------------------------------------------------

type Tally = { checked: number; skipped: Map<string, number> };
const LEDGER = new Map<Relation, Tally>(
  RELATIONS.map((relation) => [relation, { checked: 0, skipped: new Map() }]),
);

const tally = (relation: Relation): Tally => LEDGER.get(relation) as Tally;
/** Snapshot before a probe so earlier flows cannot satisfy its execution assertion. */
export const relationCheckCount = (relation: Relation): number => tally(relation).checked;
const checked = (relation: Relation): void => {
  tally(relation).checked += 1;
};
const skipped = (relation: Relation, reason: string): void => {
  const { skipped: reasons } = tally(relation);
  reasons.set(reason, (reasons.get(reason) ?? 0) + 1);
};

/** Which relations ran, how often, and what each skipped and why. */
export const relationSummary = (): string =>
  [
    `metamorphic relations (${FULL ? "full" : "sampled"}):`,
    ...RELATIONS.map((relation) => {
      if (!ENABLED_RELATIONS.has(relation)) return `  ${relation}: off`;
      const { checked: count, skipped: reasons } = tally(relation);
      const why = [...reasons].map(([reason, times]) => `${reason} ×${times}`).join("; ");
      return `  ${relation}: ${count} checked${why ? `; skipped: ${why}` : ""}`;
    }),
    "tolerated findings (comparisons each let pass that would have failed without it):",
    ...[...ABSORBED].map(([{ finding, what }, absorbed]) => `  ${finding}: ${what}: ${absorbed}`),
  ].join("\n");

// ---------------------------------------------------------------------------
// Views
// ---------------------------------------------------------------------------

const save = async (reviewer: Reviewer): Promise<Uint8Array> =>
  new Uint8Array(await reviewer.toBuffer());

/** What a reader sees of a block, ids aside: a new block's id is the allocator's choice. */
const blockView = (row: Row) => ({
  kind: row.kind,
  text: row.text,
  styleId: row.styleId,
  headingLevel: row.headingLevel,
  listLevel: row.listLevel,
  displayLabel: row.displayLabel,
  table: row.table && {
    tableIndex: row.table.tableIndex,
    rowIndex: row.table.rowIndex,
    cellIndex: row.table.cellIndex,
    gridColumnIndex: row.table.gridColumnIndex,
    columnSpan: row.table.columnSpan,
    rowSpan: row.table.rowSpan,
  },
});

type Comment = { text: string; anchor: string };
/** Comments by text and anchor; the order two comments on one range come in is not compared. */
const commentView = (comments: readonly Comment[]): string[] =>
  comments.map(({ text, anchor }) => `${text} @ ${anchor}`).sort();

const settledView = (state: { rows: readonly Row[]; comments: readonly Comment[] }) => ({
  blocks: state.rows.map(blockView),
  comments: commentView(state.comments),
});

/** `bytes` with every change resolved one way, saved and reopened. */
const resolvedView = async (
  bytes: Uint8Array,
  resolution: "accept" | "reject",
  story: Story = MAIN,
) => settledView(await resolvedState(bytes, resolution, story));

const rowsOf = (reviewer: Reviewer, story: Story = MAIN): Row[] =>
  blocksOfStory(reviewer, story) as unknown as Row[];
const commentsOf = (reviewer: Reviewer): Comment[] =>
  reviewer
    .getComments()
    .map((comment) => ({ text: comment.text, anchor: comment.anchoredText ?? "" }));

/** Everything the reviewer exposes, exactly, for an undo to restore. */
const exactState = (reviewer: Reviewer): string =>
  JSON.stringify({
    content: reviewer.getContent(),
    stories: reviewer.listStories().map(({ handle }) => ({
      handle,
      snapshot: reviewer.snapshotStory(handle),
    })),
    comments: reviewer.getComments(),
    changes: reviewer.getChanges(),
    notes: reviewer.getNotesAsText(),
  });

/** `:<type>` for an entry with a string `type` (a change, a block), else nothing. */
const typeTag = (value: unknown): string => {
  const type =
    typeof value === "object" && value !== null ? (value as { type?: unknown }).type : undefined;
  return typeof type === "string" ? `:${type}` : "";
};

/**
 * Every path where `a` and `b` differ (at most `limit`), as `path: a → b`.
 * An array entry with a `type` shows it in the path (`[2:insertion].text`),
 * so which kind of change differs is part of the failure's fingerprint.
 */
export const differences = (a: unknown, b: unknown, path = "", limit = 6): string[] => {
  const out: string[] = [];
  const walk = (x: unknown, y: unknown, at: string): void => {
    if (out.length >= limit || Object.is(x, y)) return;
    if (typeof x === "object" && x !== null && typeof y === "object" && y !== null) {
      const keys = [...new Set([...Object.keys(x), ...Object.keys(y)])];
      for (const key of keys) {
        const left = (x as Record<string, unknown>)[key];
        const right = (y as Record<string, unknown>)[key];
        const step = Array.isArray(x) ? `[${key}${typeTag(left) || typeTag(right)}]` : `.${key}`;
        walk(left, right, `${at}${step}`);
      }
      return;
    }
    if (typeof x === "string" && typeof y === "string") {
      // Long texts: from just before where they part.
      let from = 0;
      while (from < x.length && x[from] === y[from]) from += 1;
      const start = Math.max(0, from - 60);
      const cut = (text: string) =>
        `${start > 0 ? "…" : ""}${JSON.stringify(text.slice(start, from + 140))}`;
      out.push(`${at || "(root)"}: ${cut(x)} → ${cut(y)}`);
      return;
    }
    const show = (value: unknown) => {
      const text = JSON.stringify(value) ?? String(value);
      return text.length > 200 ? `${text.slice(0, 200)}…` : text;
    };
    out.push(`${at || "(root)"}: ${show(x)} → ${show(y)}`);
  };
  walk(a, b, path);
  return out;
};

/** Throw `message` with where `actual` differs from `expected`, if it does. */
const assertSame = (actual: unknown, expected: unknown, message: string): void => {
  const found = differences(expected, actual);
  if (found.length > 0) throw new Error(`${message}:\n    ${found.join("\n    ")}`);
};

// ---------------------------------------------------------------------------
// Tolerances
// ---------------------------------------------------------------------------

/**
 * A difference a relation lets through because an open finding causes it.
 * Each is keyed by its `FINDINGS` entry (support/known-issues.ts), whose
 * minimal repro runs as an expected failure: when a fix lands, that expected
 * failure fails, the entry comes out of `FINDINGS`, and the tolerance stops
 * compiling until it is deleted in the same change.
 */
type ToleranceName = {
  finding: Finding;
  /** What the tolerance leaves out of the comparison. */
  what: string;
};
type Tolerance<T> = { name: ToleranceName; apply: (value: T) => T };

/** Every tolerance, with how many comparisons it let pass that would have failed without it. */
const ABSORBED = new Map<ToleranceName, number>();

/**
 * Where `a` and `b` differ once `tolerances` are applied to both. A pair the
 * tolerances reconcile is counted against each tolerance it needed (every
 * one of them when no single one is needed on its own).
 */
const tolerantDifferences = <T>(a: T, b: T, tolerances: readonly Tolerance<T>[]): string[] => {
  const through = (value: T, list: readonly Tolerance<T>[]): T => {
    let current = value;
    for (const tolerance of list) current = tolerance.apply(current);
    return current;
  };
  const found = differences(through(a, tolerances), through(b, tolerances));
  if (found.length > 0 || differences(a, b).length === 0) return found;
  const needed = tolerances.filter((tolerance) => {
    const rest = tolerances.filter((other) => other !== tolerance);
    return differences(through(a, rest), through(b, rest)).length > 0;
  });
  for (const { name } of needed.length > 0 ? needed : tolerances) {
    ABSORBED.set(name, (ABSORBED.get(name) ?? 0) + 1);
  }
  return found;
};

type BlocksView = { blocks: readonly { text: string; table?: unknown }[] };

/**
 * Markdown as a save must keep it: revision ids are renumbered on save, and
 * two adjacent runs with the same emphasis, which a save joins, read as one
 * run. Neither is a finding: the package says the same either way.
 */
const comparableMarkdown = (markdown: string): string =>
  markdown.replace(/(<(?:ins|del)\b[^>]*?) id="[^"]*"/gu, "$1").replaceAll("****", "");

// ---------------------------------------------------------------------------
// Package parts
// ---------------------------------------------------------------------------

const decoder = new TextDecoder();

/** The text parts of a package, by name (a minimal ZIP reader: stored or deflated entries). */
export const packageParts = (bytes: Uint8Array): Map<string, string> => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.length - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06_05_4b_50) end -= 1;
  if (end < 0) throw new Error("not a ZIP package");
  const count = view.getUint16(end + 10, true);
  let offset = view.getUint32(end + 16, true);
  const parts = new Map<string, string>();
  for (let index = 0; index < count; index += 1) {
    const method = view.getUint16(offset + 10, true);
    const size = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const local = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const data = bytes.subarray(start, start + size);
    if (/\.(?:xml|rels)$/u.test(name)) {
      parts.set(name, decoder.decode(method === 0 ? data : inflateRawSync(data)));
    } else {
      parts.set(name, `<${String(size)} bytes, method ${String(method)}>`);
    }
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return parts;
};

/** Parts a save may legitimately change on every save: the timestamps in `docProps/`. */
const VOLATILE_PART = /^docProps\//u;

/** Where two packages' parts differ, or null. */
export const partDifference = (left: Uint8Array, right: Uint8Array): string | null => {
  const a = packageParts(left);
  const b = packageParts(right);
  const names = [...new Set([...a.keys(), ...b.keys()])]
    .filter((name) => !VOLATILE_PART.test(name))
    .sort();
  for (const name of names) {
    const x = a.get(name);
    const y = b.get(name);
    if (x === y) continue;
    if (x === undefined || y === undefined) {
      return `${name} is ${x === undefined ? "missing from the first" : "missing from the second"} save`;
    }
    let at = 0;
    while (at < x.length && x[at] === y[at]) at += 1;
    const around = (text: string) => JSON.stringify(text.slice(Math.max(0, at - 80), at + 120));
    return `${name} differs at ${String(at)}:\n    first:  ${around(x)}\n    second: ${around(y)}`;
  }
  return null;
};

// ---------------------------------------------------------------------------
// Batch ≡ one at a time
// ---------------------------------------------------------------------------

const INSERTIONS = new Set(["insertAfterBlock", "insertBeforeBlock", "insertTable"]);
const ANNOTATIONS = new Set(["formatRange", "commentOnRange", "commentOnBlock"]);
const RANGED = new Set(["replaceRange", "formatRange", "commentOnRange"]);

type Handle = { blockId: string; startOffset: number; endOffset: number };
const handleOf = (operation: AnyOperation): Handle | null =>
  RANGED.has(operation.type) ? (operation["range"] as Handle) : null;
const blockIdOf = (operation: AnyOperation): string =>
  handleOf(operation)?.blockId ?? String(operation["blockId"]);

/**
 * Where an offset of `before` stands in `after`, when the edit between them
 * lies wholly on one side of `[start, end)`; null when it touches the span.
 */
const rebase = (before: string, after: string, start: number, end: number): number | null => {
  if (before === after) return 0;
  let prefix = 0;
  const limit = Math.min(before.length, after.length);
  while (prefix < limit && before[prefix] === after[prefix]) prefix += 1;
  let suffix = 0;
  while (
    suffix < limit - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  if (end <= prefix) return 0;
  if (start >= before.length - suffix) return after.length - before.length;
  return null;
};

/**
 * `operation`, made against `preRows`, restated against `rows`: the same
 * block, the same words. Null when what it names is no longer there.
 */
const restate = (
  operation: AnyOperation,
  preRows: readonly Row[],
  rows: readonly Row[],
  textHashOf: (blockId: string) => string | undefined,
): AnyOperation | null => {
  const blockId = blockIdOf(operation);
  const pre = preRows.find((row) => row.id === blockId);
  const now = rows.find((row) => row.id === blockId);
  if (!pre || !now) return null;
  const restated: AnyOperation = { ...operation };
  const precondition = operation["precondition"] as { blockTextHash?: string } | undefined;
  if (precondition?.blockTextHash !== undefined) {
    // What a caller re-reading the document would send: the block's hash now.
    restated["precondition"] = { ...precondition, blockTextHash: textHashOf(blockId) };
  }
  const handle = handleOf(operation);
  if (handle) {
    const shift = rebase(pre.text, now.text, handle.startOffset, handle.endOffset);
    if (shift === null) return null;
    const range = createFolioAITextRangeHandle({
      blockId,
      text: now.text,
      startOffset: handle.startOffset + shift,
      endOffset: handle.endOffset + shift,
    });
    if (
      !range ||
      now.text.slice(range.startOffset, range.endOffset) !==
        pre.text.slice(handle.startOffset, handle.endOffset)
    ) {
      return null;
    }
    restated["range"] = range;
  }
  if (operation.type === "splitBlock") {
    const offset = Number(operation["offset"]);
    const shift = rebase(pre.text, now.text, offset, offset);
    if (shift === null) return null;
    restated["offset"] = offset + shift;
  }
  for (const key of ["find", "quote"]) {
    const value = operation[key];
    if (typeof value === "string" && value.length > 0 && !now.text.includes(value)) return null;
  }
  return restated;
};

/** Where the batch places an operation, in the pre-state: [block, inside, offset]. */
const placement = (operation: AnyOperation, preRows: readonly Row[]): [number, number, number] => {
  const index = preRows.findIndex((row) => row.id === blockIdOf(operation));
  const text = preRows[index]?.text ?? "";
  const handle = handleOf(operation);
  if (handle) return [index, 1, handle.startOffset];
  switch (operation.type) {
    case "replaceInBlock":
      return [index, 1, Math.max(0, text.indexOf(String(operation["find"])))];
    case "splitBlock":
      return [index, 1, Number(operation["offset"])];
    case "mergeBlockWithNext":
    case "insertAfterBlock":
    case "insertTable":
      return [index + 1, 0, 0];
    case "setBlockParagraphProperties":
    case "insertBeforeBlock":
      return [index, 0, 0];
    default:
      return [index, 1, 0];
  }
};

/** The one-at-a-time groups, in the batch contract's order. */
export const sequentialGroups = (
  applied: readonly AnyOperation[],
  preRows: readonly Row[],
): AnyOperation[][] => {
  const indexed = applied.map((operation, index) => ({ operation, index }));
  const isColumnEdit = (operation: AnyOperation): boolean =>
    operation.type === "insertTableColumn" || operation.type === "deleteTableColumn";
  const rest = indexed
    .filter(
      ({ operation }) =>
        !INSERTIONS.has(operation.type) &&
        !ANNOTATIONS.has(operation.type) &&
        !isColumnEdit(operation),
    )
    .toSorted((left, right) => {
      const a = placement(left.operation, preRows);
      const b = placement(right.operation, preRows);
      return b[0] - a[0] || b[1] - a[1] || b[2] - a[2] || right.index - left.index;
    });
  const columnLocation = (operation: AnyOperation): [number, number] => {
    const targetId = blockIdOf(operation);
    const target = preRows.find((row) => row.id === targetId);
    if (!target?.table)
      throw new Error(`Applied ${operation.type} anchor ${targetId} has no source table.`);
    const sourceColumn = target.table.gridColumnIndex;
    const column =
      operation.type === "insertTableColumn" && operation["position"] !== "before"
        ? sourceColumn + target.table.columnSpan
        : sourceColumn;
    return [target.table.tableIndex, column];
  };
  const columns = indexed
    .filter(({ operation }) => isColumnEdit(operation))
    .toSorted((left, right) => {
      const [leftTable, leftColumn] = columnLocation(left.operation);
      const [rightTable, rightColumn] = columnLocation(right.operation);
      if (leftTable !== rightTable) return rightTable - leftTable;
      if (leftColumn !== rightColumn) return rightColumn - leftColumn;
      const leftIsInsertion = left.operation.type === "insertTableColumn";
      const rightIsInsertion = right.operation.type === "insertTableColumn";
      if (leftIsInsertion !== rightIsInsertion) return leftIsInsertion ? -1 : 1;
      return right.index - left.index;
    });
  const insertions = applied.filter((operation) => INSERTIONS.has(operation.type));
  return [
    ...applied.filter((operation) => ANNOTATIONS.has(operation.type)).map((one) => [one]),
    // Insertions sharing a gap keep their input order, which is the batch's
    // own contract, so they are stated together the same way.
    ...(insertions.length > 0 ? [insertions] : []),
    ...rest.map(({ operation }) => [operation]),
    ...columns.map(({ operation }) => [operation]),
  ];
};

const batchOf = (batch: Batch, mode: Mode, operations: readonly AnyOperation[]): Batch =>
  ({ ...batch, mode, operations }) as unknown as Batch;
const operationsOf = (batch: Batch): AnyOperation[] =>
  batch.operations as unknown as AnyOperation[];
const appliedIds = (result: Result): Set<string> => new Set(result.applied.map(({ id }) => id));
const types = (operations: readonly AnyOperation[]): string =>
  operations.map((operation) => operation.type).join(", ");

/** What a batch left, as the relation compares it in `mode`. */
const outcomeView = async (reviewer: Reviewer, mode: Mode, story: Story) => {
  if (mode === "suggested")
    return settledView({ rows: rowsOf(reviewer, story), comments: commentsOf(reviewer) });
  const bytes = await save(reviewer);
  if (mode !== "direct") return resolvedView(bytes, "accept", story);
  const reopened = await openReviewer(bytes);
  return settledView({ rows: rowsOf(reopened, story), comments: commentsOf(reopened) });
};

// ---------------------------------------------------------------------------
// Direct ≡ tracked, reject-all ≡ original: replays of a flow's batches
// ---------------------------------------------------------------------------

type TraceEntry = { story: Story; operations: AnyOperation[] };

type Shadows = {
  direct: Reviewer;
  tracked: Reviewer;
  suggested: Reviewer | null;
  /** Where direct first refused what tracked applied; the two part there. */
  parted: string | null;
  /** Every batch replayed that tracked applied something of, as sent. */
  trace: TraceEntry[];
};

const openShadows = async (fixture: Uint8Array, suggested: boolean): Promise<Shadows> => ({
  direct: await openReviewer(fixture),
  tracked: await openReviewer(fixture),
  suggested: suggested ? await openReviewer(fixture) : null,
  parted: null,
  trace: [],
});

const contractBatch = (mode: Mode, operations: readonly AnyOperation[]): Batch =>
  ({ version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations }) as unknown as Batch;

/** What an operation names reads as in `reviewer`: its block, and the next one a merge pulls in. */
const targetText = (reviewer: Reviewer, story: Story, operation: AnyOperation): string => {
  const blocks = rowsOf(reviewer, story);
  const index = blocks.findIndex((block) => block.id === blockIdOf(operation));
  if (index === -1) return "(absent)";
  const next = operation.type === "mergeBlockWithNext" ? `\u0000${blocks[index + 1]?.text}` : "";
  return `${blocks[index]?.text}${next}`;
};

/**
 * Replay one batch: tracked (and suggested) as sent, direct with only what
 * tracked applied. Whether tracked applied anything.
 */
const replayInto = (shadows: Shadows, entry: TraceEntry, context: string): boolean => {
  const { story, operations } = entry;
  // An operation means the same in both documents only while what it names
  // reads the same in both: tracked, a merged paragraph reads apart from the
  // next one until the join is accepted, directly it has joined already.
  const targetsBefore = (reviewer: Reviewer) =>
    operations.map((operation) => targetText(reviewer, story, operation));
  const directTargets = shadows.parted === null ? targetsBefore(shadows.direct) : [];
  const trackedTargets = shadows.parted === null ? targetsBefore(shadows.tracked) : [];
  const apply = (reviewer: Reviewer, mode: Mode, selected: readonly AnyOperation[]): Result => {
    const batch = contractBatch(mode, selected);
    return story.type === "main"
      ? reviewer.applyDocumentOperations(batch)
      : reviewer.applyDocumentOperationsToStory({ story, batch });
  };
  const tracked = apply(shadows.tracked, "tracked-changes", operations);
  if (shadows.suggested) apply(shadows.suggested, "suggested", operations);
  const applied = appliedIds(tracked);
  if (applied.size === 0) return false;
  shadows.trace.push(entry);
  if (shadows.parted === null) {
    const differing = operations.filter(
      (operation, index) =>
        applied.has(operation.id) && directTargets[index] !== trackedTargets[index],
    );
    if (differing.length > 0) {
      shadows.parted = `${context}: ${differing.map(({ id }) => id).join(", ")} name blocks that read otherwise`;
      return true;
    }
    const direct = apply(
      shadows.direct,
      "direct",
      operations.filter((operation) => applied.has(operation.id)),
    );
    const refused = direct.skipped.filter(({ id }) => applied.has(id));
    if (refused.length > 0) {
      shadows.parted = `${context}: ${refused.map(({ id, reason }) => `${id} ${reason}`).join(", ")}`;
    }
  }
  return true;
};

type ShadowCheck = "directTracked" | "rejectTracked" | "rejectSuggested";

/**
 * Where the replays break `check`; null when they cannot be compared (direct
 * refused what tracked applied). `rejectSuggested` rejects the suggestions
 * in place.
 */
const shadowProblems = async (
  shadows: Shadows,
  check: ShadowCheck,
  fixture: Uint8Array,
): Promise<string[] | null> => {
  if (check === "directTracked" && shadows.parted !== null) return null;
  const suggested = shadows.suggested;
  if (check === "rejectSuggested" && suggested === null) return [];
  if (check === "rejectSuggested") suggested?.rejectAll();
  const stories = new Map(shadows.trace.map(({ story }) => [JSON.stringify(story), story]));
  const problems: string[] = [];
  for (const [name, story] of stories) {
    let found: string[];
    switch (check) {
      case "directTracked":
        found = tolerantDifferences<BlocksView>(
          await resolvedView(await save(shadows.direct), "accept", story),
          await resolvedView(await save(shadows.tracked), "accept", story),
          [],
        );
        break;
      case "rejectTracked":
        found = tolerantDifferences<BlocksView>(
          { blocks: (await resolvedView(fixture, "reject", story)).blocks },
          { blocks: (await resolvedView(await save(shadows.tracked), "reject", story)).blocks },
          [],
        );
        break;
      case "rejectSuggested":
        if (!suggested) return [];
        found = tolerantDifferences<BlocksView>(
          { blocks: (await resolvedView(fixture, "reject", story)).blocks },
          { blocks: (await resolvedView(await save(suggested), "accept", story)).blocks },
          [],
        );
        break;
    }
    problems.push(...found.map((difference) => `${name}: ${difference}`));
  }
  return problems;
};

/** The fewest batches and operations of `trace` that still break `check`. */
const minimizeTrace = async (
  fixture: Uint8Array,
  trace: readonly TraceEntry[],
  check: ShadowCheck,
): Promise<TraceEntry[]> => {
  const fails = async (candidate: readonly TraceEntry[]): Promise<boolean> => {
    const shadows = await openShadows(fixture, check === "rejectSuggested");
    for (const entry of candidate) replayInto(shadows, entry, "minimizing");
    const problems = await shadowProblems(shadows, check, fixture);
    return problems !== null && problems.length > 0;
  };
  let current = [...trace];
  for (let index = current.length - 1; index >= 0; index -= 1) {
    const candidate = current.toSpliced(index, 1);
    if (await fails(candidate)) current = candidate;
  }
  for (let index = current.length - 1; index >= 0; index -= 1) {
    for (
      let operation = (current[index]?.operations.length ?? 0) - 1;
      operation >= 0;
      operation -= 1
    ) {
      const entry = current[index];
      if (!entry || entry.operations.length < 2) break;
      const candidate = current.with(index, {
        ...entry,
        operations: entry.operations.toSpliced(operation, 1),
      });
      if (await fails(candidate)) current = candidate;
    }
  }
  return current;
};

// ---------------------------------------------------------------------------
// One flow
// ---------------------------------------------------------------------------

/**
 * A live pre-batch copy, including pending suggestions that a DOCX save drops.
 * Applying a no-op transaction gives each copy its own editor and plugin
 * state. Mutable reviewer collections are copied too. The recorder's own
 * method wrappers are omitted, leaving the prototype's operation methods.
 */
const cloneReviewer = (reviewer: Reviewer): Reviewer => {
  // SAFETY: this test-only clone retains FolioDocxReviewer's prototype and
  // every own field. Its EditorState copies have independent plugin state.
  const copy = Object.create(Object.getPrototypeOf(reviewer)) as Reviewer;
  for (const key of Reflect.ownKeys(reviewer)) {
    if (key === "applyDocumentOperations" || key === "applyDocumentOperationsToStory") continue;
    const descriptor = Object.getOwnPropertyDescriptor(reviewer, key);
    if (!descriptor) continue;
    const value = descriptor.value;
    let cloned = value;
    if (key === "state") {
      cloned = value.apply(value.tr);
    } else if (key === "secondaryStoryStates") {
      cloned = new Map(
        [...value].map(([storyKey, entry]) => [
          storyKey,
          {
            ...entry,
            state: entry.state.apply(entry.state.tr),
            initialState: entry.initialState.apply(entry.initialState.tr),
          },
        ]),
      );
    } else if (value instanceof Map) {
      cloned = new Map(value);
    } else if (value instanceof Set) {
      cloned = new Set(value);
    } else if (Array.isArray(value)) {
      cloned = [...value];
    }
    Object.defineProperty(copy, key, {
      ...descriptor,
      value: cloned,
    });
  }
  return copy;
};

type Recorded = { story: Story; batch: Batch; result: Result; pre: Reviewer };

export type FlowRelations = {
  /**
   * After a step and the flow's own save: `saved` is that save and its
   * reopened reviewer. Throws when a relation does not hold.
   */
  afterStep: (
    reviewer: Reviewer,
    saved: { bytes: Uint8Array; reopened: Reviewer },
    context: string,
  ) => Promise<void>;
  /** After the last step. */
  finish: () => Promise<void>;
};

/**
 * Start the relations for one flow over `fixture` (the package the flow's
 * `reviewer` was opened from), in the flow's `mode`. Every batch applied to
 * the flow's reviewer is recorded as it applies (the reviewer's own
 * `applyDocumentOperations`, which `suggest_changes` reaches through the
 * bridge, or `applyDocumentOperationsToStory`) and checked at the step's end.
 */
export const startRelations = async ({
  fixture,
  reviewer,
  mode,
  seed,
}: {
  fixture: Uint8Array;
  reviewer: Reviewer;
  mode: Mode;
  seed: number;
}): Promise<FlowRelations> => {
  const on = (relation: Relation) => ENABLED_RELATIONS.has(relation);
  // Its own stream: the flow's is never drawn from.
  const random: Random = createRandom(seed ^ 0x5eed_4e1a);
  const recorded: Recorded[] = [];
  const instrumented = new WeakSet<Reviewer>();
  const instrument = (target: Reviewer): void => {
    if (instrumented.has(target) || ENABLED_RELATIONS.size === 0) return;
    instrumented.add(target);
    const apply = target.applyDocumentOperations.bind(target);
    target.applyDocumentOperations = (batch, options) => {
      const pre = cloneReviewer(target);
      const result = apply(batch, options);
      recorded.push({ story: MAIN, batch: structuredClone(batch), result, pre });
      return result;
    };
    const applyToStory = target.applyDocumentOperationsToStory.bind(target);
    target.applyDocumentOperationsToStory = (options) => {
      const pre = cloneReviewer(target);
      const result = applyToStory(options);
      recorded.push({
        story: structuredClone(options.story),
        batch: structuredClone(options.batch),
        result,
        pre,
      });
      return result;
    };
  };
  instrument(reviewer);

  // The replays of `directTracked` and `rejectAll`.
  const shadows =
    on("directTracked") || on("rejectAll") ? await openShadows(fixture, on("rejectAll")) : null;
  const check = async (relation: Relation, which: ShadowCheck, context: string, what: string) => {
    if (!shadows) return;
    const problems = await shadowProblems(shadows, which, fixture);
    if (problems === null) {
      skipped(relation, "the direct and tracked replays parted");
      return;
    }
    if (problems.length === 0) {
      checked(relation);
      return;
    }
    const minimal = await minimizeTrace(fixture, shadows.trace, which);
    throw new Error(
      `${context}: [${relation}] ${what}:\n    ${problems.join("\n    ")}\n  minimal replay on the flow's fixture (one batch per line):\n    ${minimal.map((entry) => JSON.stringify(entry)).join("\n    ")}`,
    );
  };

  const compareShadows = async (context: string): Promise<void> => {
    if (!shadows || shadows.trace.length === 0) return;
    if (on("directTracked")) {
      await check(
        "directTracked",
        "directTracked",
        context,
        "the flow's batches replayed tracked and accepted read otherwise than replayed direct (direct → tracked)",
      );
    }
    if (on("rejectAll")) {
      await check(
        "rejectAll",
        "rejectTracked",
        context,
        "the flow's batches replayed tracked and rejected do not give the fixture back (fixture → rejected)",
      );
    }
  };

  const replay = async ({ story, batch }: Recorded, context: string): Promise<void> => {
    if (!shadows) return;
    if (replayInto(shadows, { story, operations: operationsOf(batch) }, context) && FULL)
      await compareShadows(context);
  };

  /** `undo` and `batchSequential` on copies of the document before the batch. */
  const onCopies = async ({ story, batch, result, pre }: Recorded, context: string) => {
    const sequential = on("batchSequential") && (FULL || random.chance(SEQUENTIAL_SHARE));
    if (!on("undo") && !sequential) return;
    const batchMode = (batch.mode ?? "tracked-changes") as Mode;
    const copy = cloneReviewer(pre);
    const preRows = rowsOf(copy, story);
    const preState = exactState(copy);
    const applyToCopy = (target: Reviewer, operations: Batch): Result =>
      story.type === "main"
        ? target.applyDocumentOperations(operations)
        : target.applyDocumentOperationsToStory({ story, batch: operations });
    const again = applyToCopy(copy, structuredClone(batch));
    const applied = operationsOf(batch).filter((operation) => appliedIds(again).has(operation.id));
    assert.deepEqual(
      [...appliedIds(again)].sort(),
      [...appliedIds(result)].sort(),
      `${context}: the live pre-batch copy answered differently than the flow reviewer`,
    );
    const batchOutcome =
      sequential && applied.length >= 2 ? await outcomeView(copy, batchMode, story) : null;

    if (on("undo")) {
      if (again.undoHandle === null) {
        skipped("undo", applied.length === 0 ? "nothing applied" : "no undo handle");
      } else {
        const undone = copy.undoDocumentOperations(again.undoHandle);
        assert.equal(
          undone.status,
          "undone",
          `${context}: [undo] the batch's own undo handle is refused (${types(applied)})`,
        );
        assertSame(
          JSON.parse(exactState(copy)),
          JSON.parse(preState),
          `${context}: [undo] undoing the batch (${types(applied)}) does not give the document before back (before → undone)`,
        );
        checked("undo");
      }
    }

    if (!sequential) return;
    if (batchOutcome === null) {
      skipped("batchSequential", "fewer than two operations applied");
      return;
    }
    const oneByOne = cloneReviewer(pre);
    const problems: string[] = [];
    for (const group of sequentialGroups(applied, preRows)) {
      const snapshot = story.type === "main" ? oneByOne.snapshot() : oneByOne.snapshotStory(story);
      if (!snapshot) throw new Error(`${context}: the story disappeared before sequential replay`);
      const rows = snapshot.blocks as unknown as Row[];
      const textHashOf = (blockId: string) => snapshot.anchors[blockId]?.textHash;
      const restated = group.map((operation) => restate(operation, preRows, rows, textHashOf));
      if (restated.some((operation) => operation === null)) {
        skipped("batchSequential", "an operation does not re-resolve one at a time");
        return;
      }
      const alone = applyToCopy(oneByOne, batchOf(batch, batchMode, restated as AnyOperation[]));
      for (const { id, reason } of alone.skipped) {
        if (reason !== "noopOperation") problems.push(`${id} refused one at a time: ${reason}`);
      }
    }
    const sequentialOutcome = await outcomeView(oneByOne, batchMode, story);
    if (problems.length === 0) {
      problems.push(
        ...differences(batchOutcome, sequentialOutcome).map((d) => `batch → one at a time ${d}`),
      );
    }
    if (problems.length > 0) {
      throw new Error(
        `${context}: [batchSequential] the batch (${batchMode}: ${types(applied)}) leaves otherwise than its operations one at a time:\n  ${problems.join("\n  ")}`,
      );
    }
    checked("batchSequential");
  };

  const readerStability = (live: Reviewer, reopened: Reviewer, context: string): void => {
    const content = projectContentPair({
      leftRows: live.getContent(),
      rightRows: reopened.getContent(),
    });
    const views = (target: Reviewer, side: "left" | "right") => ({
      getContent: content[side],
      snapshot: projectSnapshotIdentities(
        createReviewerBridge(target).snapshot(),
        content[side === "left" ? "leftAliases" : "rightAliases"],
      ),
      toMarkdown: comparableMarkdown(toMarkdown(target.toDocument())) as unknown,
      // Revision ids and dates can change on save. Keep every entry's kind,
      // author, content and containing block, including duplicate entries.
      getChanges: target
        .getChanges()
        .map(({ type, author, text, blockId }) => ({
          type,
          author,
          text,
          blockId:
            blockId === null
              ? null
              : (content[side === "left" ? "leftAliases" : "rightAliases"].get(blockId) ?? blockId),
        }))
        .toSorted((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right))),
      getComments: target.getComments() as unknown,
    });
    const tolerances: Record<keyof ReturnType<typeof views>, Tolerance<unknown>[]> = {
      getContent: [],
      snapshot: [],
      toMarkdown: [],
      getChanges: [],
      getComments: [],
    };
    const a = views(live, "left");
    const b = views(reopened, "right");
    for (const reader of Object.keys(a) as (keyof typeof a)[]) {
      const found = tolerantDifferences(a[reader], b[reader], tolerances[reader]);
      if (found.length > 0) {
        throw new Error(
          `${context}: [readerStability] ${reader} reads otherwise after the save than before it (before → after):\n    ${found.join("\n    ")}`,
        );
      }
    }
    checked("readerStability");
  };

  return {
    afterStep: async (live, saved, context) => {
      const batches = recorded.splice(0);
      instrument(live);
      const pending = mode === "suggested" && live.getChanges().length > 0;
      for (const entry of batches) {
        await replay(entry, context);
        await onCopies(entry, context);
      }
      if (on("saveIdempotent")) {
        const difference = partDifference(saved.bytes, await save(saved.reopened));
        assert.equal(
          difference,
          null,
          `${context}: [saveIdempotent] saving the reopened package changes it: ${difference}`,
        );
        checked("saveIdempotent");
      }
      if (on("readerStability")) {
        if (pending) skipped("readerStability", "pending suggestions stay out of the package");
        else readerStability(live, saved.reopened, context);
      }
    },
    finish: async () => {
      if (!shadows) return;
      if (!FULL) await compareShadows("end of flow");
      if (on("rejectAll") && shadows.suggested && shadows.trace.length > 0) {
        await check(
          "rejectAll",
          "rejectSuggested",
          "end of flow",
          "the flow's batches replayed as suggestions and rejected do not give the fixture back (fixture → rejected)",
        );
      }
    },
  };
};
