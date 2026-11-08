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
  hyperlinkReview: {
    id: CANONICAL_GAP.trackedHyperlinkResolution,
    gap: CANONICAL_GAP.trackedHyperlinkResolution,
    message:
      "Hyperlink edits cannot cut pending review identities or unsupported inline wrappers.",
  },
  cellSelectionPlacement: {
    id: CANONICAL_GAP.commands,
    gap: CANONICAL_GAP.commands,
    message:
      "A note or table of contents goes at a caret, not over selected cells.",
  },
  trackedTableDeletion: {
    id: CANONICAL_GAP.dispatch,
    gap: CANONICAL_GAP.dispatch,
    message: "The rows contain an unsupported tracked structure.",
  },
  trackedTableInsertionAtEnd: {
    id: CANONICAL_GAP.dispatch,
    gap: CANONICAL_GAP.dispatch,
    message:
      "Tracked insertion requires a following paragraph in the same block list.",
  },
  fieldProjection: {
    id: CANONICAL_GAP.storyContentProjection,
    gap: CANONICAL_GAP.storyContentProjection,
    message: "The operations produce unsupported canonical story content.",
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

export const STORY_PROJECTION_REFUSAL_CASES = [
  {
    shape: "fields-links-bookmarks",
    mode: "suggesting",
    operation: "key:Enter",
    placement: "paragraph",
  },
  {
    shape: "fields-links-bookmarks",
    mode: "suggesting",
    operation: "paste:copied-blocks",
    placement: "caret-end",
  },
  {
    shape: "fields-links-bookmarks",
    mode: "suggesting",
    operation: "key:Delete",
    placement: "cross-paragraph",
  },
  {
    shape: "fields-links-bookmarks",
    mode: "suggesting",
    operation: "host:cut",
    placement: "paragraph",
  },
  {
    shape: "fields-links-bookmarks",
    mode: "suggesting",
    operation: "host:cut",
    placement: "cross-paragraph",
  },
] as const satisfies readonly RefusalCase[];

const STORY_PROJECTION_ROWS = {
  "fields-links-bookmarks": REFUSAL_ROWS.fieldProjection,
} as const satisfies Record<
  (typeof STORY_PROJECTION_REFUSAL_CASES)[number]["shape"],
  HarnessRefusalRow
>;

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
  const projection = STORY_PROJECTION_REFUSAL_CASES.find(
    (candidate) =>
      canonicalRefusalCaseId(candidate) === canonicalRefusalCaseId(key),
  );
  if (projection) return [STORY_PROJECTION_ROWS[projection.shape]];
  if (
    key.shape === "tables" &&
    key.placement === "cross-paragraph" &&
    (key.operation === "command:generateTOC" ||
      key.operation === "command:insertFootnote" ||
      key.operation === "command:insertEndnote")
  )
    return [REFUSAL_ROWS.cellSelectionPlacement];
  // The caret's cell paragraph ends its block list: a tracked paragraph
  // insertion needs a paragraph after it. A selection of the whole document,
  // or of cells, puts the table ahead of a paragraph instead: the story's
  // first, or the one after the table.
  if (key.operation === "command:insertTable")
    return key.mode === "suggesting" &&
      key.shape === "tables" &&
      key.placement !== "document" &&
      key.placement !== "cross-paragraph"
      ? [REFUSAL_ROWS.trackedTableInsertionAtEnd]
      : [];
  // The fixture table's merged cells have no tracked row deletion record.
  if (key.operation === "command:deleteTable")
    return key.mode === "suggesting" && key.shape === "tables"
      ? [REFUSAL_ROWS.trackedTableDeletion]
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
  if (key.shape === "tracked-changes") return [REFUSAL_ROWS.hyperlinkReview];
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
