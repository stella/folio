/**
 * What the round-trip self-check found, and how to say it safely.
 *
 * The check compares two block projections: what accepting the generated
 * revisions leaves against the target, and what rejecting them leaves against
 * the base. A projection carries each block's container, style, list level,
 * direct alignment, direct spacing and text, so WHICH field diverged names
 * which part of the pipeline lost the difference — and that is worth reporting
 * as a typed cause rather than as one opaque "did not reproduce".
 *
 * Every `detail` string here is structural: counts, offsets, container kinds.
 * Never a phrase of either document, because a caller may log it, put it in a
 * report, or quote it in a review.
 */

import { PARAGRAPH_MARK_CHANGE_KINDS, type ParagraphMarkChangeKind } from "@stll/docx-core/model";

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import type { FolioAIBlock, FolioAIBlockPreviewRun } from "../ai-edits/types";
import { paragraphSpacingEqual } from "../prosemirror/paragraphSpacing";
import { resolveColorToHex } from "../utils/colorResolver";

const normalizeInlineFormattingColor = (color: string | undefined): string | undefined =>
  resolveColorToHex(color === undefined ? undefined : { rgb: color }, null);

/** The two directions of the round trip, each an invariant of its own. */
export const COMPARE_VERIFICATION_INVARIANTS = Object.freeze([
  "accept-reproduces-target",
  "reject-reproduces-base",
] as const);

export type CompareVerificationInvariant = (typeof COMPARE_VERIFICATION_INVARIANTS)[number];

/**
 * What diverged. Ordered from the structural to the textual: the first cause
 * that fits is the one reported, because a block in the wrong container is a
 * different finding from a block whose words are wrong even when both are true.
 */
export const COMPARE_VERIFICATION_CAUSES = Object.freeze([
  /**
   * Every block is present, in order, at coordinates the block model cannot
   * reach, so no operation could have put a block at the expected ones. Not a
   * lost difference.
   *
   * Blank paragraphs no longer cause this: the snapshot carries them. What is
   * left is the row a package hides, whose whole subtree the snapshot skips on
   * purpose. A table hiding a row on one side only shifts every later row's
   * index, and nothing the comparison can do reaches those positions.
   */
  "invisible-structure",
  "block-count",
  "container",
  /**
   * Every block is where it should be and a table's own properties are not:
   * `w:tblPr`, the `w:tblGrid` widths, `w:trPr`, `w:tcPr`. No block carries
   * them, so a projection of blocks alone cannot see them go.
   */
  "table-geometry",
  "style",
  "list-level",
  "alignment",
  "spacing",
  "inline-formatting",
  "whitespace",
  "text",
] as const);

export type CompareVerificationCause = (typeof COMPARE_VERIFICATION_CAUSES)[number];

/** One invariant that did not hold, in one story. */
export type CompareVerificationFailure = {
  invariant: CompareVerificationInvariant;
  cause: CompareVerificationCause;
  story: FolioDocumentStoryHandle;
  /** Structural facts only: counts, offsets, container kinds. Safe to quote. */
  detail: string;
};

/**
 * Whether the redline was proven to round-trip.
 *
 * `unverified` is only ever returned when the caller asked for it with
 * `onUnverified: "emit"`; the default refuses instead, because a redline that
 * reads plausibly and is wrong is worse than no redline.
 */
export type CompareVerification =
  | { status: "verified" }
  | { status: "unverified"; failures: readonly CompareVerificationFailure[] };

const supportedInlineStyle = ({
  bold,
  italic,
  underline,
  strike,
  fontFamily,
  fontSizePt,
  color,
  directFormatting,
}: FolioAIBlockPreviewRun): string =>
  JSON.stringify([
    bold === true,
    italic === true,
    underline === true,
    strike === true,
    fontFamily ?? null,
    fontSizePt ?? null,
    normalizeInlineFormattingColor(color) ?? null,
    directFormatting?.bold ?? null,
    directFormatting?.italic ?? null,
    directFormatting?.underline ?? null,
    directFormatting?.strike ?? null,
    directFormatting?.fontFamily ?? null,
    directFormatting?.fontSizePt ?? null,
    normalizeInlineFormattingColor(directFormatting?.color ?? undefined) ?? null,
  ]);

/** Effective supported formatting with equivalent adjacent runs normalized. */
export const projectSupportedInlineFormatting = ({ text, previewRuns }: FolioAIBlock): string => {
  const projected: { length: number; style: string }[] = [];
  for (const run of previewRuns ?? [{ text }]) {
    if (run.text.length === 0) {
      continue;
    }
    const style = supportedInlineStyle(run);
    const previous = projected.at(-1);
    if (previous?.style === style) {
      previous.length += run.text.length;
      continue;
    }
    projected.push({ length: run.text.length, style });
  }
  return projected.map(({ length, style }) => `${String(length)}:${style}`).join(",");
};

type ProjectedBlock = Pick<
  FolioAIBlock,
  "text" | "table" | "styleId" | "listLevel" | "directAlignment" | "directSpacing"
>;

type ProjectedTableContainer = NonNullable<ProjectedBlock["table"]>;

const sameContainer = (
  left: ProjectedTableContainer | undefined,
  right: ProjectedTableContainer | undefined,
): boolean => {
  if (!left || !right) {
    return left === right;
  }
  return (
    left.outerTableIndex === right.outerTableIndex &&
    left.tableIndex === right.tableIndex &&
    left.rowIndex === right.rowIndex &&
    left.cellIndex === right.cellIndex &&
    left.gridColumnIndex === right.gridColumnIndex &&
    left.columnSpan === right.columnSpan &&
    left.rowSpan === right.rowSpan &&
    left.paragraphIndex === right.paragraphIndex
  );
};

const sameProjectedBlock = (left: ProjectedBlock, right: ProjectedBlock): boolean =>
  sameContainer(left.table, right.table) &&
  left.styleId === right.styleId &&
  left.listLevel === right.listLevel &&
  left.directAlignment === right.directAlignment &&
  paragraphSpacingEqual(left.directSpacing, right.directSpacing) &&
  left.text === right.text;

const containerKind = (container: ProjectedTableContainer | undefined): "body" | "cell" =>
  container ? "cell" : "body";

const collapseWhitespace = (text: string): string => text.replace(/\s+/gu, " ").trim();

const sameBlockProjection = (
  left: readonly ProjectedBlock[],
  right: readonly ProjectedBlock[],
): boolean =>
  left.length === right.length &&
  left.every((entry, index) => {
    const other = right[index];
    return other !== undefined && sameProjectedBlock(entry, other);
  });

const sameStringProjection = (left: readonly string[], right: readonly string[]): boolean =>
  left.length === right.length && left.every((entry, index) => entry === right[index]);

/**
 * The same projection with every table coordinate renumbered by first
 * appearance, so it counts the blocks the model holds rather than the
 * paragraphs the package contains.
 *
 * The snapshot skips a hidden row's whole subtree, so a table that hides a row
 * on one side only reports every later row one position along. No operation
 * can put a block at those coordinates, because none can create or remove the
 * hidden row that produces them. When two projections agree here and disagree
 * on the raw coordinates, the redline holds every block the other side does,
 * in order, and the difference is one the block model cannot see.
 */
const byVisibleOrdinal = (entries: readonly ProjectedBlock[]): ProjectedBlock[] => {
  const ordinals = new Map<string, number>();
  const counts = new Map<string, number>();
  const ordinalWithin = (scope: string, index: string): number => {
    const key = `${scope}:${index}`;
    const existing = ordinals.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const next = counts.get(scope) ?? 0;
    counts.set(scope, next + 1);
    ordinals.set(key, next);
    return next;
  };

  const paragraphCounts = new Map<string, number>();
  return entries.map((entry) => {
    const container = entry.table;
    if (!container) {
      return entry;
    }
    const outerTableOrdinal = ordinalWithin("t", String(container.outerTableIndex));
    const tableOrdinal = ordinalWithin("t", String(container.tableIndex));
    const rowScope = `r${String(tableOrdinal)}`;
    const rowOrdinal = ordinalWithin(rowScope, String(container.rowIndex));
    const cellScope = `c${String(tableOrdinal)}.${String(rowOrdinal)}`;
    const cellOrdinal = ordinalWithin(cellScope, String(container.cellIndex));
    const cellKey = `${cellScope}:${String(container.cellIndex)}`;
    const paragraphOrdinal = paragraphCounts.get(cellKey) ?? 0;
    paragraphCounts.set(cellKey, paragraphOrdinal + 1);
    return {
      ...entry,
      table: {
        ...container,
        outerTableIndex: outerTableOrdinal,
        tableIndex: tableOrdinal,
        rowIndex: rowOrdinal,
        cellIndex: cellOrdinal,
        paragraphIndex: paragraphOrdinal,
      },
    };
  });
};

/**
 * Where two projections first diverge, and what diverged there. A pair that
 * matches everywhere but in length diverges at the shorter one's end.
 */
const firstDivergence = (
  actual: readonly ProjectedBlock[],
  expected: readonly ProjectedBlock[],
): { index: number; actual: ProjectedBlock | null; expected: ProjectedBlock | null } => {
  const shared = Math.min(actual.length, expected.length);
  for (let index = 0; index < shared; index++) {
    const left = actual[index];
    const right = expected[index];
    if (left !== undefined && right !== undefined && !sameProjectedBlock(left, right)) {
      return {
        index,
        actual: left,
        expected: right,
      };
    }
  }
  return {
    index: shared,
    actual: actual.at(shared) ?? null,
    expected: expected.at(shared) ?? null,
  };
};

type ClassifyOptions<T> = {
  invariant: CompareVerificationInvariant;
  story: FolioDocumentStoryHandle;
  /** What the redline actually leaves. */
  actual: readonly T[];
  /** What the invariant says it should leave. */
  expected: readonly T[];
};

/**
 * The failure two table-geometry projections describe, or `null` when they
 * agree. One line per table, so the detail names which table diverged and
 * whether the count itself did — never a property value, which could carry a
 * style name either document chose.
 */
export const classifyGeometryMismatch = ({
  invariant,
  story,
  actual,
  expected,
}: ClassifyOptions<string>): CompareVerificationFailure | null => {
  if (sameStringProjection(actual, expected)) {
    return null;
  }
  const detail =
    actual.length === expected.length
      ? `table ${String(actual.findIndex((entry, index) => entry !== expected[index]))} of ${String(actual.length)} carries different properties`
      : `${String(actual.length)} tables against ${String(expected.length)}`;
  return { invariant, cause: "table-geometry", story, detail };
};

/**
 * The failure two projections describe, or `null` when they agree.
 *
 * Total over the causes by construction: the last branch is unconditional, so
 * a divergence always produces a failure rather than being dropped.
 */
export const classifyProjectionMismatch = ({
  invariant,
  story,
  actual,
  expected,
}: ClassifyOptions<ProjectedBlock>): CompareVerificationFailure | null => {
  if (sameBlockProjection(actual, expected)) {
    return null;
  }
  const failure = (
    cause: CompareVerificationCause,
    detail: string,
  ): CompareVerificationFailure => ({
    invariant,
    cause,
    story,
    detail,
  });

  if (sameBlockProjection(byVisibleOrdinal(actual), byVisibleOrdinal(expected))) {
    return failure(
      "invisible-structure",
      `every block matches once table coordinates count visible blocks (${String(expected.length)} blocks)`,
    );
  }

  const divergence = firstDivergence(actual, expected);
  const at = `at block ${String(divergence.index)}/${String(expected.length)}`;
  const counts = `${String(actual.length)} blocks against ${String(expected.length)}`;
  if (divergence.actual === null || divergence.expected === null) {
    const side = actual.length > expected.length ? "more" : "fewer";
    return failure("block-count", `${side} blocks than expected (${counts}), diverging ${at}`);
  }
  const { actual: left, expected: right } = divergence;
  if (!sameContainer(left.table, right.table)) {
    return failure(
      "container",
      `a block sits in a ${containerKind(left.table)} where it is expected in a ${containerKind(right.table)}, ${at} (${counts})`,
    );
  }
  if (left.text === right.text && left.styleId !== right.styleId) {
    return failure("style", `the paragraph style did not move ${at} (${counts})`);
  }
  if (left.text === right.text && left.listLevel !== right.listLevel) {
    return failure("list-level", `the list level did not move ${at} (${counts})`);
  }
  if (left.text === right.text && left.directAlignment !== right.directAlignment) {
    return failure("alignment", `the direct paragraph alignment did not move ${at} (${counts})`);
  }
  if (left.text === right.text && !paragraphSpacingEqual(left.directSpacing, right.directSpacing)) {
    return failure("spacing", `the direct paragraph spacing did not move ${at} (${counts})`);
  }
  if (collapseWhitespace(left.text) === collapseWhitespace(right.text)) {
    return failure("whitespace", `a block's text differs only in whitespace ${at} (${counts})`);
  }
  if (actual.length !== expected.length) {
    return failure("block-count", `${counts}, first differing ${at}`);
  }
  return failure(
    "text",
    `a ${containerKind(left.table)} block's text does not match ${at}, ` +
      `length ${String(left.text.length)} against ${String(right.text.length)} (${counts})`,
  );
};

/**
 * A container whose last paragraph carries a revision on its mark.
 *
 * Structural facts only — a path through the package model and an index — so
 * the finding is safe to log, report or quote.
 */
export type FinalParagraphMarkRevision = {
  /** Where the container sits in the package model, e.g. `package.document.content`. */
  container: string;
  /** The paragraph's index among its container's children. */
  paragraphIndex: number;
  /** The mark kind found there, `moveFrom` / `moveTo` for a relocation's ends. */
  kind: ParagraphMarkChangeKind;
};

const isAParagraphMarkChangeKind = (value: unknown): value is ParagraphMarkChangeKind =>
  PARAGRAPH_MARK_CHANGE_KINDS.some((kind) => kind === value);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const isParagraph = (value: unknown): value is Record<string, unknown> =>
  isRecord(value) && value["type"] === "paragraph";

/**
 * A row the package marks deleted, whose cells' marks go with it.
 *
 * A deleted table row is written as `w:trPr/w:del` PLUS a deletion on every
 * mark the row's cells end with: accepting takes the whole row away, so those
 * marks never have to join anything and are the shape the format asks for.
 */
const isADeletedTableRow = (value: Record<string, unknown>): boolean => {
  const change = value["structuralChange"];
  return isRecord(change) && change["type"] === "tableRowDeletion";
};

/**
 * Every container in a package whose final paragraph mark carries a revision.
 *
 * A deleted paragraph mark says "join this paragraph with the one after it",
 * and an inserted one says that break was ADDED, so rejecting it closes the
 * paragraph back over the next one. The last paragraph of a body, a table
 * cell, a header or footer, a note or a text box has no paragraph after it, so
 * neither direction states an edit that can be carried out: a consumer refuses
 * the package, or opens it and leaves a revision standing that neither
 * accepting nor rejecting everything can clear. The exception is a cell of a
 * row the package is DELETING: there the mark leaves with its row.
 *
 * The walk is over the package model rather than over a list of the containers
 * known today: a container is any sequence that ends in a paragraph, so a part
 * the model grows later is covered the day it arrives instead of the day
 * someone remembers this function.
 *
 * `since` scopes it to the revisions a comparison MINTED: a base may arrive
 * carrying one of these on a paragraph in a part no story mounts, which folio
 * preserves the way it preserves everything else it parses. What this proves
 * is that the comparison writes none of its own.
 */
export const revisedFinalParagraphMarks = (
  packageModel: unknown,
  { since = 0 }: { since?: number } = {},
): FinalParagraphMarkRevision[] => {
  const found: FinalParagraphMarkRevision[] = [];
  // The path is kept as a stack of raw keys and formatted only where something
  // is found. Building the string at every node instead made the walk cost
  // more than the serialization it guards, on a package where it never has
  // anything to report.
  const trail: (string | number)[] = ["package"];
  const pathOf = (): string => {
    let path = "";
    for (const segment of trail) {
      path += typeof segment === "number" ? `[${String(segment)}]` : `.${segment}`;
    }
    return path.slice(1);
  };
  const visit = (value: unknown, insideADeletedRow: boolean): void => {
    if (Array.isArray(value)) {
      const last: unknown = value.at(-1);
      const mark = isParagraph(last) ? last["pPrMark"] : undefined;
      const kind = isRecord(mark) ? mark["kind"] : undefined;
      const info = isRecord(mark) ? mark["info"] : undefined;
      const revisionId = isRecord(info) ? info["id"] : undefined;
      const isOurs = typeof revisionId === "number" ? revisionId >= since : true;
      if (isAParagraphMarkChangeKind(kind) && isOurs && !insideADeletedRow) {
        found.push({ container: pathOf(), paragraphIndex: value.length - 1, kind });
      }
      for (const [index, item] of value.entries()) {
        trail.push(index);
        visit(item, insideADeletedRow);
        trail.pop();
      }
      return;
    }
    if (value instanceof Map) {
      for (const [key, item] of value) {
        trail.push(String(key));
        visit(item, insideADeletedRow);
        trail.pop();
      }
      return;
    }
    if (!isRecord(value) || value instanceof Date || ArrayBuffer.isView(value)) {
      return;
    }
    const inADeletedRow = insideADeletedRow || isADeletedTableRow(value);
    for (const [key, item] of Object.entries(value)) {
      trail.push(key);
      visit(item, inADeletedRow);
      trail.pop();
    }
  };
  visit(packageModel, false);
  return found;
};
