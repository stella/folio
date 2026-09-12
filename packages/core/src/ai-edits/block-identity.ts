import { getFolioParaIdFromBlockId } from "../types/block-id";
import type { FolioAIBlock } from "./types";

const IDENTITY_BLOCK_ATTRS = new Set(["paraId", "textId"]);

/** Resolve explicit snapshot provenance, falling back to the block-id shape convention. */
export const folioAIBlockIdStability = ({
  id,
  idStability,
}: FolioAIBlock): "stable" | "positional" =>
  idStability ?? (getFolioParaIdFromBlockId(id) === null ? "positional" : "stable");

export const stripBlockIdentityAttrs = (attrs: Record<string, unknown>) => {
  const next: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    next[key] = IDENTITY_BLOCK_ATTRS.has(key) ? null : value;
  }
  return next;
};
