import { panic } from "better-result";

import {
  CANONICAL_CAPABILITIES,
  CANONICAL_GAP,
} from "../packages/core/src/types/canonicalCapabilities";
import type {
  CanonicalGap,
  CanonicalRule,
} from "../packages/core/src/types/canonicalCapabilities";
import type { Document } from "../packages/core/src/types/document";
import type { BrowserInputAction } from "../tests/visual/browserInputTrace";

export type HarnessRefusal = {
  /** The ledger gap, absent for a rule refusal, which no row declares. */
  gap: CanonicalGap | undefined;
  /** The document rule, present only for a rule refusal. */
  rule?: CanonicalRule;
  message: string;
  expectation: "declared" | "unexpected";
  row?: string;
};

export type HarnessRefusalRow = {
  id: string;
  gap: CanonicalGap;
  message: string;
};

export const canonicalActivationRefusalRow = (
  _source: Document,
): HarnessRefusalRow | undefined => undefined;

/** Rows are created by the attempted input, rather than mirroring ledger membership. */
export const validateHarnessRefusalRows = (
  rows: readonly HarnessRefusalRow[],
  capabilities: Readonly<Record<string, unknown>> = CANONICAL_CAPABILITIES,
) => {
  for (const row of rows)
    if (!Object.hasOwn(capabilities, row.gap))
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
          (row) =>
            row.id === refusal.row &&
            row.gap === refusal.gap &&
            row.message === refusal.message,
        ),
    )
    .map(({ gap, message }) => `${gap}: ${message}`);
};

const clipboardTableRows: readonly HarnessRefusalRow[] = [
  {
    id: "clipboard-table",
    gap: CANONICAL_GAP.trackedTableStructure,
    message:
      "Clipboard tables and embedded blocks require canonical table editing.",
  },
  {
    id: "clipboard-nested-blocks",
    gap: CANONICAL_GAP.trackedTableStructure,
    message:
      "Clipboard tables and nested block containers require canonical table editing.",
  },
];

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
      gap: CANONICAL_GAP.dispatch,
      message: "Unclassified native text is unavailable in this session.",
    },
  ],
  // A paste clears every selected cell and lands in the first.
  pastePlain: [],
  pasteHtml: [],
  pasteWordHtml: [],
  pasteListHtml: [],
  pasteMultiBlock: [],
  pasteTable: clipboardTableRows,
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

/** Native table paste and input over cells are explicit; the rest is strict. */
export const canonicalBrowserRefusalRows = (
  action: BrowserInputAction,
  { selection }: CanonicalBrowserInputContext = { selection: "other" },
): readonly HarnessRefusalRow[] => {
  if (selection === "cells") return SELECTED_CELL_REFUSALS[action.kind];
  return action.kind === "pasteTable" ? clipboardTableRows : [];
};

type MatchCanonicalRefusalOptions = {
  rows: readonly HarnessRefusalRow[];
  refusal: { gap: CanonicalGap; message: string };
};
export const matchCanonicalRefusalRow = ({
  rows,
  refusal,
}: MatchCanonicalRefusalOptions) => {
  validateHarnessRefusalRows(rows);
  return rows.find(
    (row) => row.gap === refusal.gap && row.message === refusal.message,
  );
};
