import { panic } from "better-result";

import {
  CANONICAL_CAPABILITIES,
  CANONICAL_CONDITION,
  CANONICAL_GAP,
} from "../packages/core/src/types/canonicalCapabilities";
import type {
  CanonicalCondition,
  CanonicalGap,
  CanonicalRule,
} from "../packages/core/src/types/canonicalCapabilities";
import type { Document } from "../packages/core/src/types/document";
import type { BrowserInputAction } from "../tests/visual/browserInputTrace";

export type HarnessRefusal = {
  /** The ledger gap, absent for a rule or condition refusal. */
  gap: CanonicalGap | undefined;
  /** The document rule, present only for a rule refusal, which no row declares. */
  rule?: CanonicalRule;
  /** The session or request condition, present only for a condition refusal. */
  condition?: CanonicalCondition;
  message: string;
  expectation: "declared" | "unexpected";
  row?: string;
};

/** A refusal an input is declared to meet: under a ledger gap or a condition. */
export type HarnessRefusalRow = {
  id: string;
  message: string;
} & (
  | { gap: CanonicalGap; condition?: never }
  | { condition: CanonicalCondition; gap?: never }
);

/** Whether `refusal` is the one `row` declares: the same ground and message. */
export const refusalMatchesRow = (
  row: HarnessRefusalRow,
  refusal: {
    gap?: CanonicalGap | undefined;
    condition?: CanonicalCondition | undefined;
    message: string;
  },
) =>
  row.gap === refusal.gap &&
  row.condition === refusal.condition &&
  row.message === refusal.message;

export const canonicalActivationRefusalRow = (
  _source: Document,
): HarnessRefusalRow | undefined => undefined;

/** Rows are created by the attempted input, rather than mirroring ledger membership. */
export const validateHarnessRefusalRows = (
  rows: readonly HarnessRefusalRow[],
  capabilities: Readonly<Record<string, unknown>> = CANONICAL_CAPABILITIES,
) => {
  for (const row of rows)
    if (row.gap !== undefined && !Object.hasOwn(capabilities, row.gap))
      panic(
        `Retired canonical refusal row must become strict: ${row.id} (${row.gap})`,
      );
};

type HarnessRefusalProblemsOptions = {
  rows: readonly HarnessRefusalRow[];
  refusals: readonly HarnessRefusal[];
};
export const harnessRefusalProblems = ({
  rows,
  refusals,
}: HarnessRefusalProblemsOptions) => {
  validateHarnessRefusalRows(rows);
  return refusals
    .filter(
      (refusal) =>
        !rows.some(
          (row) => row.id === refusal.row && refusalMatchesRow(row, refusal),
        ),
    )
    .map(
      ({ gap, rule, condition, message }) =>
        `${gap ?? rule ?? condition}: ${message}`,
    );
};

/**
 * Input over a cell selection other than typing and deletion needs a per-cell
 * operation the session does not have yet. Each of these leaves the document
 * as it was.
 */
const SELECTED_CELL_REFUSALS = {
  // Typing clears every selected cell and writes into the first.
  typing: [],
  // Enter clears every selected cell and breaks the first one's paragraph.
  enter: [],
  imeReplacement: [
    {
      id: "selected-cells-composition",
      gap: CANONICAL_GAP.storyContentProjection,
      message:
        "Composition changed unsupported content; the canonical document was restored.",
    },
    // The browser's own cleanup of a cancelled proposal, after the refusal above.
    {
      id: "selected-cells-composition-cleanup",
      condition: CANONICAL_CONDITION.invalidRequest,
      message: "Unclassified native text is unavailable in this session.",
    },
  ],
  // A paste clears every selected cell and lands in the first.
  pastePlain: [],
  pasteHtml: [],
  pasteWordHtml: [],
  pasteListHtml: [],
  pasteMultiBlock: [],
  pasteTable: [],
  backspace: [],
  delete: [],
  cut: [],
  dragCellDelete: [],
  undo: [],
  redo: [],
  historyBurst: [],
  selectionDrag: [],
  selectAll: [],
} as const satisfies Record<
  BrowserInputAction["kind"],
  readonly HarnessRefusalRow[]
>;

/** What the attempted input meets, as the session holds it before the input. */
export type CanonicalBrowserInputContext = {
  selection: "cells" | "other";
};

/** Input over cells is explicit; the rest is strict. */
export const canonicalBrowserRefusalRows = (
  action: BrowserInputAction,
  { selection }: CanonicalBrowserInputContext = { selection: "other" },
): readonly HarnessRefusalRow[] => {
  return selection === "cells" ? SELECTED_CELL_REFUSALS[action.kind] : [];
};

type MatchCanonicalRefusalOptions = {
  rows: readonly HarnessRefusalRow[];
  refusal: {
    gap?: CanonicalGap | undefined;
    condition?: CanonicalCondition | undefined;
    message: string;
  };
};
export const matchCanonicalRefusalRow = ({
  rows,
  refusal,
}: MatchCanonicalRefusalOptions) => {
  validateHarnessRefusalRows(rows);
  return rows.find((row) => refusalMatchesRow(row, refusal));
};
