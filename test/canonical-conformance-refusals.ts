import { panic } from "better-result";

import type {
  EditorMode,
  SelectionPlacement,
} from "../packages/core/src/__tests__/editorHarness";
import type { HarnessRefusalRow } from "./canonical-refusal-rows";

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
  _key: RefusalCase,
): readonly HarnessRefusalRow[] => [];

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
