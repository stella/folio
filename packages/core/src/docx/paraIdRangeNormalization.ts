/**
 * Keep every paragraph id a package carries inside the range the schema gives
 * it.
 *
 * `w14:paraId`, `w14:textId` and the comment-part ids that reference a
 * paragraph are `ST_LongHexNumber` with a maximum: the value has to be below
 * `0x80000000`, so the ids are 31-bit. Producers exist that write eight hex
 * digits without that bound, and folio preserves the ids a document arrives
 * with — so a package can carry an out-of-range id in, and a save that copies
 * it through hands a consumer a package it will refuse.
 *
 * {@link paraIdInRange} is the one mapping, and it is a pure function of the
 * value alone. That is what lets the parser and the save agree without
 * consulting each other: a paragraph's id in the model is the id the file gets,
 * so bringing an id into range does not make a document's own identity move
 * under it between reading and writing. The package pass below is the same
 * mapping applied to every attribute that carries such an id, so a paragraph
 * and every reference to it move together.
 */

import { deterministicHexId } from "../utils/hexId";

/** Exclusive upper bound on a paragraph id: the values are 31-bit. */
const MAX_PARA_ID_EXCLUSIVE = 0x8000_0000;

/**
 * Every attribute that carries a paragraph id or the text-revision marker
 * written beside one: the paragraph's own `w14:paraId` / `w14:textId` (the
 * parser accepts a `w:` prefix too), the comment part's `w15:paraId` and the
 * `w15:paraIdParent` that links a reply to its thread, and the durable-comment
 * part's `w16cid:paraId`.
 */
const PARA_ID_ATTRIBUTE =
  /\b(w|w14|w15|w16cid):(paraId|paraIdParent|textId)=(?<quote>["'])([0-9A-Fa-f]{8})\k<quote>/gu;

/** A candidate part is one that mentions any of those attributes at all. */
const PARA_ID_CANDIDATE = /\b(?:w|w14|w15|w16cid):(?:paraId|paraIdParent|textId)=/u;

/**
 * `value` when it already fits, and a value derived from it when it does not.
 *
 * The replacement is derived from the id being replaced and from nothing else.
 * A package-aware search for a free id would read better here and would be
 * wrong: the parser has no package to search, and an id that means one
 * paragraph while reading and another while writing is worse than the
 * vanishing chance of two rewritten ids landing on one value, which is a
 * duplicate rather than a package a consumer refuses.
 */
export const paraIdInRange = (value: string): string =>
  Number.parseInt(value, 16) < MAX_PARA_ID_EXCLUSIVE ? value : deterministicHexId(value);

/**
 * Rewrite out-of-range paragraph ids across a whole package.
 *
 * Returns the parts unchanged when every id already fits, so a save of a
 * document that never carried one is byte-identical.
 */
export const normalizeParaIdRangeInXmlParts = (
  parts: ReadonlyMap<string, string>,
): Map<string, string> => {
  const normalized = new Map(parts);
  for (const [path, xml] of parts) {
    if (!PARA_ID_CANDIDATE.test(xml)) {
      continue;
    }
    const rewritten = xml.replaceAll(
      PARA_ID_ATTRIBUTE,
      (whole: string, prefix: string, name: string, quote: string, value: string) => {
        const replacement = paraIdInRange(value);
        return replacement === value ? whole : `${prefix}:${name}=${quote}${replacement}${quote}`;
      },
    );
    if (rewritten !== xml) {
      normalized.set(path, rewritten);
    }
  }
  return normalized;
};
