/**
 * What the round-trip self-check found, and how to say it safely.
 *
 * The check compares two block projections: what accepting the generated
 * revisions leaves against the target, and what rejecting them leaves against
 * the base. A projection carries each block's container, style, list level and
 * text, so WHICH field diverged names which part of the pipeline lost the
 * difference — and that is worth reporting as a typed cause rather than as one
 * opaque "did not reproduce".
 *
 * Every `detail` string here is structural: counts, offsets, container kinds.
 * Never a phrase of either document, because a caller may log it, put it in a
 * report, or quote it in a review.
 */

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import type { FolioAIBlock, FolioAIBlockPreviewRun } from "../ai-edits/types";

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
  "style",
  "list-level",
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
}: FolioAIBlockPreviewRun): string =>
  `${bold === true ? "b" : ""}${italic === true ? "i" : ""}${underline === true ? "u" : ""}${
    strike === true ? "s" : ""
  }`;

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

/**
 * One block of a projection. `projectStory` writes
 * `container|styleId|listLevel|text`, so the text is everything after the
 * third separator and may itself contain one.
 */
type ProjectedBlock = {
  container: string;
  styleId: string;
  listLevel: string;
  text: string;
};

const parseProjection = (entry: string): ProjectedBlock => {
  const first = entry.indexOf("|");
  const second = entry.indexOf("|", first + 1);
  const third = entry.indexOf("|", second + 1);
  if (first === -1 || second === -1 || third === -1) {
    return { container: "", styleId: "", listLevel: "", text: entry };
  }
  return {
    container: entry.slice(0, first),
    styleId: entry.slice(first + 1, second),
    listLevel: entry.slice(second + 1, third),
    text: entry.slice(third + 1),
  };
};

const CONTAINER_PATTERN = /^t(\d+)r(\d+)c(\d+)p(\d+)$/u;

const containerKind = (container: string): "body" | "cell" =>
  CONTAINER_PATTERN.test(container) ? "cell" : "body";

const collapseWhitespace = (text: string): string => text.replace(/\s+/gu, " ").trim();

/** Element by element, never by joining: a block's text may hold the separator. */
export const sameProjection = (left: readonly string[], right: readonly string[]): boolean =>
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
const byVisibleOrdinal = (entries: readonly string[]): string[] => {
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
    const first = entry.indexOf("|");
    const match = CONTAINER_PATTERN.exec(entry.slice(0, first));
    if (!match) {
      return entry;
    }
    const [, table = "", row = "", cell = ""] = match;
    const tableOrdinal = ordinalWithin("t", table);
    const rowScope = `r${String(tableOrdinal)}`;
    const rowOrdinal = ordinalWithin(rowScope, row);
    const cellScope = `c${String(tableOrdinal)}.${String(rowOrdinal)}`;
    const cellOrdinal = ordinalWithin(cellScope, cell);
    const cellKey = `${cellScope}:${cell}`;
    const paragraphOrdinal = paragraphCounts.get(cellKey) ?? 0;
    paragraphCounts.set(cellKey, paragraphOrdinal + 1);
    return (
      `t${String(tableOrdinal)}r${String(rowOrdinal)}c${String(cellOrdinal)}p${String(paragraphOrdinal)}` +
      entry.slice(first)
    );
  });
};

/**
 * Where two projections first diverge, and what diverged there. A pair that
 * matches everywhere but in length diverges at the shorter one's end.
 */
const firstDivergence = (
  actual: readonly string[],
  expected: readonly string[],
): { index: number; actual: ProjectedBlock | null; expected: ProjectedBlock | null } => {
  const shared = Math.min(actual.length, expected.length);
  for (let index = 0; index < shared; index++) {
    if (actual[index] !== expected[index]) {
      return {
        index,
        actual: parseProjection(actual[index] ?? ""),
        expected: parseProjection(expected[index] ?? ""),
      };
    }
  }
  return {
    index: shared,
    actual: shared < actual.length ? parseProjection(actual[shared] ?? "") : null,
    expected: shared < expected.length ? parseProjection(expected[shared] ?? "") : null,
  };
};

type ClassifyOptions = {
  invariant: CompareVerificationInvariant;
  story: FolioDocumentStoryHandle;
  /** What the redline actually leaves. */
  actual: readonly string[];
  /** What the invariant says it should leave. */
  expected: readonly string[];
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
}: ClassifyOptions): CompareVerificationFailure | null => {
  if (sameProjection(actual, expected)) {
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

  if (sameProjection(byVisibleOrdinal(actual), byVisibleOrdinal(expected))) {
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
  if (left.container !== right.container) {
    return failure(
      "container",
      `a block sits in a ${containerKind(left.container)} where it is expected in a ${containerKind(right.container)}, ${at} (${counts})`,
    );
  }
  if (left.text === right.text && left.styleId !== right.styleId) {
    return failure("style", `the paragraph style did not move ${at} (${counts})`);
  }
  if (left.text === right.text && left.listLevel !== right.listLevel) {
    return failure("list-level", `the list level did not move ${at} (${counts})`);
  }
  if (collapseWhitespace(left.text) === collapseWhitespace(right.text)) {
    return failure("whitespace", `a block's text differs only in whitespace ${at} (${counts})`);
  }
  if (actual.length !== expected.length) {
    return failure("block-count", `${counts}, first differing ${at}`);
  }
  return failure(
    "text",
    `a ${containerKind(left.container)} block's text does not match ${at}, ` +
      `length ${String(left.text.length)} against ${String(right.text.length)} (${counts})`,
  );
};

/**
 * A container whose last paragraph carries a deletion on its mark.
 *
 * Structural facts only — a path through the package model and an index — so
 * the finding is safe to log, report or quote.
 */
export type FinalParagraphMarkDeletion = {
  /** Where the container sits in the package model, e.g. `package.document.content`. */
  container: string;
  /** The paragraph's index among its container's children. */
  paragraphIndex: number;
  /** The mark kind found there: `del`, or `moveFrom` for a relocation's source. */
  kind: "del" | "moveFrom";
};

/** The mark kinds that resolve by joining the paragraph with the one after it. */
const JOINS_FORWARD_ON_ACCEPT = Object.freeze(["del", "moveFrom"] as const);

const joinsForwardOnAccept = (value: unknown): value is FinalParagraphMarkDeletion["kind"] =>
  JOINS_FORWARD_ON_ACCEPT.some((kind) => kind === value);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const isParagraph = (value: unknown): value is Record<string, unknown> =>
  isRecord(value) && value["type"] === "paragraph";

/**
 * Every container in a package whose final paragraph mark carries a deletion.
 *
 * A deleted paragraph mark says "join this paragraph with the one after it".
 * The last paragraph of a body, a table cell, a header or footer, a note or a
 * text box has no paragraph after it, so the mark states an edit that cannot
 * be carried out, and a consumer refuses the package rather than opening it.
 *
 * The walk is over the package model rather than over a list of the containers
 * known today: a container is any sequence that ends in a paragraph, so a part
 * the model grows later is covered the day it arrives instead of the day
 * someone remembers this function.
 */
export const deletedFinalParagraphMarks = (packageModel: unknown): FinalParagraphMarkDeletion[] => {
  const found: FinalParagraphMarkDeletion[] = [];
  const visit = (value: unknown, path: string): void => {
    if (Array.isArray(value)) {
      const last: unknown = value.at(-1);
      const mark = isParagraph(last) ? last["pPrMark"] : undefined;
      const kind = isRecord(mark) ? mark["kind"] : undefined;
      if (joinsForwardOnAccept(kind)) {
        found.push({ container: path, paragraphIndex: value.length - 1, kind });
      }
      for (const [index, item] of value.entries()) {
        visit(item, `${path}[${String(index)}]`);
      }
      return;
    }
    if (value instanceof Map) {
      for (const [key, item] of value) {
        visit(item, `${path}.${String(key)}`);
      }
      return;
    }
    if (!isRecord(value) || value instanceof Date || ArrayBuffer.isView(value)) {
      return;
    }
    for (const [key, item] of Object.entries(value)) {
      visit(item, `${path}.${key}`);
    }
  };
  visit(packageModel, "package");
  return found;
};
