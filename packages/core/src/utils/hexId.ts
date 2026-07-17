/**
 * Lifted from
 * https://github.com/eigenpal/docx-editor/blob/main/packages/core/src/utils/hexId.ts
 * (Apache-2.0). Keep the bound and the comments in sync upstream.
 * folio divergence: `00000000` is excluded from both generators — Word
 * treats a zero `w14:paraId` as unassigned, so it must never be minted.
 */

/**
 * Strictest OOXML `ST_LongHexNumber` upper bound (exclusive) across the
 * fields this helper feeds: `w14:paraId` / `w14:textId` / comment
 * `paraId` (`< 0x80000000`) and `w16cid:commentId/@durableId`
 * (`< 0x7FFFFFFF`). Generated ids must stay strictly below this value
 * to survive both Word ("Document Recovery — Table Properties") and
 * strict OOXML validators.
 */
// eslint-disable-next-line unicorn/numeric-separators-style -- hex is the readable form here; the value is the OOXML upper bound documented above.
export const MAX_HEX_ID_EXCLUSIVE = 0x7fffffff;

/**
 * OOXML `ST_LongHexNumber` shape: exactly 8 hex digits (`xsd:hexBinary`,
 * length 4 bytes). `w14:paraId` / `w14:textId` / comment `paraId` /
 * `w15:paraIdParent` are all typed this way. A value that doesn't match this
 * is not a real Word id — it must not be trusted as one (e.g. echoed
 * unescaped into serialized XML attributes).
 */
export const HEX_ID_PATTERN = /^[0-9A-Fa-f]{8}$/u;

/** Whether `value` is a well-formed 8-hex-digit OOXML long-hex id. */
export const isValidHexId = (value: string | undefined | null): boolean =>
  typeof value === "string" && HEX_ID_PATTERN.test(value);

/**
 * Random 8-char uppercase hex id, matching Microsoft's `w14:paraId`
 * extension format (also reused for comment `paraId` / `durableId`).
 *
 * Range is `[1, MAX_HEX_ID_EXCLUSIVE)` = `[1, 0x7FFFFFFE]` — zero is
 * excluded because Word reads `w14:paraId="00000000"` as "no id".
 *
 * Uses `Math.random()` rather than `crypto.randomUUID()` so the
 * generator works in non-secure contexts (file://, web workers).
 */
export const generateHexId = (): string =>
  (Math.floor(Math.random() * (MAX_HEX_ID_EXCLUSIVE - 1)) + 1)
    .toString(16)
    .toUpperCase()
    .padStart(8, "0");

/**
 * Deterministic 8-char uppercase hex id (FNV-1a over `seed`), `< 0x7FFFFFFF`.
 * Re-deriving from the same seed mints the same id, so content/position-derived
 * ids (block paraIds, comment thread keys) are stable across saves.
 *
 * A hash of exactly zero is remapped to `1`: `00000000` is Word's "no id"
 * sentinel, and remapping only that one (already invalid) output keeps every
 * previously shipped deterministic id unchanged.
 */
export const deterministicHexId = (seed: string): string => {
  let hash = 2_166_136_261;
  for (const character of seed) {
    hash = Math.imul(hash ^ (character.codePointAt(0) ?? 0), 16_777_619) >>> 0;
  }
  const value = hash % MAX_HEX_ID_EXCLUSIVE;
  return (value === 0 ? 1 : value).toString(16).toUpperCase().padStart(8, "0");
};
