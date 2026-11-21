import { getFolioParaIdFromBlockId } from "../types/block-id";
import type { FolioAIBlock } from "./types";

/**
 * What a paragraph's copy must not claim: its ids, what its source authored
 * once (the editing sessions its `w:rsid*` attributes name, its bookmark
 * boundaries), and where a past layout broke the page before it. A save
 * writes the authored ones on one record only, the first in document order,
 * so a copy placed before its source would take them from it; a copy placed
 * after it would write the page break again.
 */
const IDENTITY_BLOCK_ATTRS = new Set([
  "paraId",
  "textId",
  "_preservedAttributes",
  "bookmarks",
  "renderedPageBreakBefore",
]);

/** Resolve explicit snapshot provenance, falling back to the legacy id-shape convention. */
export const folioAIBlockIdStability = ({
  id,
  idStability,
}: FolioAIBlock): "stable" | "positional" =>
  idStability ??
  (getFolioParaIdFromBlockId(id) === null ? "positional" : "stable");

/**
 * A block as the package holds it. `idStability` is the reading's fact, not
 * the package's: a save writes a minted id into a paragraph it rewrites, so the
 * same paragraph reads as stated once reopened.
 */
export const persistedBlock = (block: FolioAIBlock): FolioAIBlock => {
  const persisted = { ...block };
  delete persisted.idStability;
  return persisted;
};

export const stripBlockIdentityAttrs = (
  attrs: Record<string, unknown>,
): Record<string, unknown> => {
  const next: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    next[key] = IDENTITY_BLOCK_ATTRS.has(key) ? null : value;
  }
  return next;
};
