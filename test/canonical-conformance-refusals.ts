import { panic } from "better-result";

import type {
  EditorMode,
  SelectionPlacement,
} from "../packages/core/src/__tests__/editorHarness";
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import type { HarnessRefusalRow } from "./canonical-refusal-rows";

// These are fixture contracts, independent of the command planner's verdict.
const REFUSAL_ROWS = {
  cellSelectionPlacement: {
    id: CANONICAL_GAP.commands,
    gap: CANONICAL_GAP.commands,
    message:
      "A note or table of contents goes at a caret, not over selected cells.",
  },
  // A suggested column deletion keeps its cells until accepted, so a cell
  // spanning the deleted column and another has no record of its narrowing.
  suggestedSpanningColumnDeletion: {
    id: "suggested-spanning-column-deletion",
    gap: CANONICAL_GAP.trackedTableStructure,
    message:
      "A suggested column deletion cannot narrow a cell that spans other columns.",
  },
} as const satisfies Record<string, HarnessRefusalRow>;

// Explicit unsupported cases; a controller descriptor/planner verdict must
// never add an operation to this contract. Commands outside it stay strict.
const COMMAND_REFUSAL_CONTRACTS = {
  [CANONICAL_GAP.commands]: {
    operations: [
      "command:autoFitContents",
      "host:addComment",
      "host:acceptChange",
      "host:rejectChange",
    ],
    rows: [
      {
        id: CANONICAL_GAP.commands,
        gap: CANONICAL_GAP.dispatch,
        message: "Unclassified native text is unavailable in this session.",
      },
    ],
  },
} as const satisfies Record<
  typeof CANONICAL_GAP.commands,
  {
    operations: readonly string[];
    rows: readonly HarnessRefusalRow[];
  }
>;

type RefusalCase = {
  shape: string;
  operation: string;
  placement: SelectionPlacement;
  mode: EditorMode;
};

export const canonicalRefusalCaseId = ({
  shape,
  operation,
  placement,
  mode,
}: RefusalCase) => `${shape} › ${operation} @ ${placement} [${mode}]`;

// The fixture source/heading collector assertion binds this applicability set.
export const TOC_HEADING_SHAPES = [
  "style-numbered-headings",
  "outline-level-numbered",
] as const;

/** Every declared case has an explicit expectation, including supported cases with no rows. */
export const canonicalConformanceRefusalRows = (
  key: RefusalCase,
): readonly HarnessRefusalRow[] => {
  const commandContract = Object.values(COMMAND_REFUSAL_CONTRACTS).find(
    ({ operations }) =>
      operations.some((operation) => operation === key.operation),
  );
  if (commandContract) return commandContract.rows;
  if (
    key.shape === "tables" &&
    key.placement === "cross-paragraph" &&
    (key.operation === "command:generateTOC" ||
      key.operation === "command:insertFootnote" ||
      key.operation === "command:insertEndnote")
  )
    return [REFUSAL_ROWS.cellSelectionPlacement];
  // Suggestions in the fixture's table. Over the whole document the selection
  // is in no table, so the table commands do nothing in either mode.
  const inFixtureTable =
    key.shape === "tables" &&
    key.mode === "suggesting" &&
    key.placement !== "document";
  // Deleting the table or its rows records every row, a nested table's too.
  // Across the focused cell and the next, the selection spans the merged
  // block; at a caret the column holds none.
  if (key.operation === "command:deleteColumn")
    return inFixtureTable && key.placement === "cross-paragraph"
      ? [REFUSAL_ROWS.suggestedSpanningColumnDeletion]
      : [];
  // Link and contents-list suggestions are recorded as struck text and
  // inserted copies, so a fixture refuses none of them.
  return [];
};

/** Build the total case table from declarations, never from observed controller errors. */
export const declareCanonicalRefusalCases = (cases: readonly RefusalCase[]) => {
  const rows = new Map<string, readonly HarnessRefusalRow[]>();
  for (const key of cases) {
    const id = canonicalRefusalCaseId(key);
    if (rows.has(id)) panic(`Duplicate canonical refusal case: ${id}`);
    rows.set(id, canonicalConformanceRefusalRows(key));
  }
  return rows;
};
