import { panic } from "better-result";
import { CANONICAL_GAP } from "../packages/core/src/types/canonicalCapabilities";
import type { EditorMode, SelectionPlacement } from "../packages/core/src/__tests__/editorHarness";
import type { HarnessRefusalRow } from "./canonical-refusal-rows";

// These are fixture contracts, independent of the command planner's verdict.
const REFUSAL_ROWS = {
  tableActivation: {
    id: "table-session-activation",
    gap: CANONICAL_GAP.tableActivation,
    message: "Canonical sessions cannot activate documents containing tables.",
  },
  hyperlinkSuggestion: {
    id: CANONICAL_GAP.trackedHyperlinkResolution,
    gap: CANONICAL_GAP.trackedHyperlinkResolution,
    message: "Hyperlink suggestions require serializable wrapper review provenance.",
  },
  hyperlinkReview: {
    id: CANONICAL_GAP.trackedHyperlinkResolution,
    gap: CANONICAL_GAP.trackedHyperlinkResolution,
    message: "Hyperlink edits cannot cut pending review identities or unsupported inline wrappers.",
  },
  commentProjection: {
    id: CANONICAL_GAP.storyContentProjection,
    gap: CANONICAL_GAP.storyContentProjection,
    message: "The paragraph cannot be projected as plain text.",
  },
  fieldProjection: {
    id: CANONICAL_GAP.storyContentProjection,
    gap: CANONICAL_GAP.storyContentProjection,
    message: "The operations produce unsupported canonical story content.",
  },
} as const satisfies Record<string, HarnessRefusalRow>;

type RefusalCase = {
  shape: string;
  operation: string;
  placement: SelectionPlacement;
  mode: EditorMode;
};

export const canonicalRefusalCaseId = ({ shape, operation, placement, mode }: RefusalCase) =>
  `${shape} › ${operation} @ ${placement} [${mode}]`;

export const STORY_PROJECTION_REFUSAL_CASES = [
  { shape: "comments", mode: "editing", operation: "key:Enter", placement: "paragraph" },
  { shape: "comments", mode: "editing", operation: "host:cut", placement: "paragraph" },
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
  comments: REFUSAL_ROWS.commentProjection,
  "fields-links-bookmarks": REFUSAL_ROWS.fieldProjection,
} as const satisfies Record<
  (typeof STORY_PROJECTION_REFUSAL_CASES)[number]["shape"],
  HarnessRefusalRow
>;

/** Every declared case has an explicit expectation, including supported cases with no rows. */
export const canonicalConformanceRefusalRows = (key: RefusalCase): readonly HarnessRefusalRow[] => {
  if (key.shape === "tables") return [REFUSAL_ROWS.tableActivation];
  const projection = STORY_PROJECTION_REFUSAL_CASES.find(
    (candidate) => canonicalRefusalCaseId(candidate) === canonicalRefusalCaseId(key),
  );
  if (projection) return [STORY_PROJECTION_ROWS[projection.shape]];
  const range = ["word", "paragraph", "cross-paragraph", "document", "node"].includes(
    key.placement,
  );
  const hyperlink =
    key.operation === "command:insertHyperlink" ||
    (range &&
      (key.operation === "command:setHyperlink" || key.operation === "command:removeHyperlink"));
  if (!hyperlink) return [];
  if (key.mode === "suggesting") return [REFUSAL_ROWS.hyperlinkSuggestion];
  if (key.shape === "comments" || key.shape === "tracked-changes")
    return [REFUSAL_ROWS.hyperlinkReview];
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
