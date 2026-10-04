import { panic } from "better-result";
import { documentStories, findStoryBody } from "../packages/docx-core/src/ops/stories";
import {
  CANONICAL_CAPABILITIES,
  CANONICAL_GAP,
} from "../packages/core/src/types/canonicalCapabilities";
import type { CanonicalGap } from "../packages/core/src/types/canonicalCapabilities";
import type { Document } from "../packages/core/src/types/document";
import type { BrowserInputAction } from "../tests/visual/browserInputTrace";

export type HarnessRefusal = {
  gap: CanonicalGap;
  message: string;
  expectation: "declared" | "unexpected";
  row?: string;
};

export type HarnessRefusalRow = {
  id: string;
  gap: CanonicalGap;
  message: string;
};

export const canonicalActivationRefusalRow = (source: Document): HarnessRefusalRow | undefined =>
  documentStories(source).some((story) =>
    findStoryBody(source, story)?.content.some((block) => block.type === "table"),
  )
    ? {
        id: "table-session-activation",
        gap: CANONICAL_GAP.tableActivation,
        message: "Canonical sessions cannot activate documents containing tables.",
      }
    : undefined;

/** Rows are created by the attempted input, rather than mirroring ledger membership. */
export const validateHarnessRefusalRows = (
  rows: readonly HarnessRefusalRow[],
  capabilities: Readonly<Record<string, unknown>> = CANONICAL_CAPABILITIES,
) => {
  for (const row of rows)
    if (!Object.hasOwn(capabilities, row.gap))
      panic(`Retired canonical refusal row must become strict: ${row.id} (${row.gap})`);
};

type HarnessRefusalProblemsOptions = {
  rows: readonly HarnessRefusalRow[];
  refusals: readonly HarnessRefusal[];
};
export const harnessRefusalProblems = ({ rows, refusals }: HarnessRefusalProblemsOptions) => {
  validateHarnessRefusalRows(rows);
  return refusals
    .filter(
      (refusal) =>
        !rows.some(
          (row) =>
            row.id === refusal.row && row.gap === refusal.gap && row.message === refusal.message,
        ),
    )
    .map(({ gap, message }) => `${gap}: ${message}`);
};

/** Native table paste remains explicit; every other generated input is strict. */
export const canonicalBrowserRefusalRows = (
  action: BrowserInputAction,
): readonly HarnessRefusalRow[] => {
  if (action.kind !== "pasteTable") return [];
  return [
    {
      id: "clipboard-table",
      gap: CANONICAL_GAP.dispatch,
      message: "Clipboard tables and embedded blocks require canonical table editing.",
    },
    {
      id: "clipboard-nested-blocks",
      gap: CANONICAL_GAP.dispatch,
      message: "Clipboard tables and nested block containers require canonical table editing.",
    },
  ];
};

type MatchCanonicalRefusalOptions = {
  rows: readonly HarnessRefusalRow[];
  refusal: { gap: CanonicalGap; message: string };
};
export const matchCanonicalRefusalRow = ({ rows, refusal }: MatchCanonicalRefusalOptions) => {
  validateHarnessRefusalRows(rows);
  return rows.find((row) => row.gap === refusal.gap && row.message === refusal.message);
};
