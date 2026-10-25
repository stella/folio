import { getFolioParaIdFromBlockId } from "../types/block-id";
import type { FolioAIBlock } from "./types";

/**
 * What a paragraph's copy must not claim: its ids, and what its source
 * authored once (the editing sessions its `w:rsid*` attributes name, its
 * bookmark boundaries). A save writes those on one record only, the first in
 * document order, so a copy placed before its source would take them from it.
 */
const IDENTITY_BLOCK_ATTRS = new Set([
  "paraId",
  "textId",
  "_preservedAttributes",
  "bookmarks",
]);

/** Resolve explicit snapshot provenance, falling back to the legacy id-shape convention. */
export const folioAIBlockIdStability = ({
  id,
  idStability,
}: FolioAIBlock): "stable" | "positional" =>
  idStability ??
  (getFolioParaIdFromBlockId(id) === null ? "positional" : "stable");

export const stripBlockIdentityAttrs = (attrs: Record<string, unknown>) => {
  const next: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    next[key] = IDENTITY_BLOCK_ATTRS.has(key) ? null : value;
  }
  return next;
};
