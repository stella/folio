import { panic } from "better-result";

import type {
  EditorMode,
  SelectionPlacement,
} from "../packages/core/src/__tests__/editorHarness";
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import type { HarnessRefusalRow } from "./canonical-refusal-rows";

// These are fixture contracts, independent of the command planner's verdict.
const REFUSAL_ROWS = {
  hyperlinkSuggestion: {
    id: CANONICAL_GAP.trackedHyperlinkResolution,
    gap: CANONICAL_GAP.trackedHyperlinkResolution,
    message:
      "Hyperlink and TOC suggestions require serializable wrapper review provenance.",
  },
  cellSelectionPlacement: {
    id: CANONICAL_GAP.commands,
    gap: CANONICAL_GAP.commands,
    message:
      "A note or table of contents goes at a caret, not over selected cells.",
  },
  // The fixture table's third row holds a nested table, which a suggested
  // row change has no record for.
  suggestedNestedTableRows: {
    id: "suggested-nested-table-rows",
    gap: CANONICAL_GAP.trackedTableStructure,
    message: "A suggestion cannot yet record a row that holds a nested table.",
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
  // Deleting the table records every row, the nested table's too.
  if (key.operation === "command:deleteTable")
    return key.shape === "tables" && key.mode === "suggesting"
      ? [REFUSAL_ROWS.suggestedNestedTableRows]
      : [];
  // Across the focused cell and the next, the selection spans the merged
  // block and reaches the row with the nested table; at a caret the row and
  // column hold neither.
  if (key.operation === "command:deleteRow")
    return inFixtureTable && key.placement === "cross-paragraph"
      ? [REFUSAL_ROWS.suggestedNestedTableRows]
      : [];
  if (key.operation === "command:deleteColumn")
    return inFixtureTable && key.placement === "cross-paragraph"
      ? [REFUSAL_ROWS.suggestedSpanningColumnDeletion]
      : [];
  if (key.operation === "command:generateTOC")
    return key.mode === "suggesting" &&
      TOC_HEADING_SHAPES.some((shape) => shape === key.shape)
      ? [REFUSAL_ROWS.hyperlinkSuggestion]
      : [];
  const range = [
    "word",
    "paragraph",
    "cross-paragraph",
    "document",
    "node",
  ].includes(key.placement);
  const hyperlink =
    key.operation === "command:insertHyperlink" ||
    (range &&
      (key.operation === "command:setHyperlink" ||
        key.operation === "command:removeHyperlink"));
  if (!hyperlink) return [];
  // Over selected cells a hyperlink command changes nothing in either mode, so
  // there is nothing to refuse.
  if (key.shape === "tables" && key.placement === "cross-paragraph") return [];
  if (key.mode === "suggesting") return [REFUSAL_ROWS.hyperlinkSuggestion];
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
