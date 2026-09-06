/**
 * Single source of truth for folio block ids.
 *
 * Both the in-browser editor snapshot (`createFolioAIEditSnapshot`)
 * and the server-side DOCX extractor (`apps/api/.../docx-blocks.ts`)
 * import {@link deriveBlockId} so a citation written by the server
 * resolves in the editor without a separate mapping table. Any other
 * way of minting an id — `b-${n}`, `${idx}`, hand-rolled prefixes —
 * is by definition unable to produce a {@link FolioBlockId}: the
 * branded type makes the divergence a compile-time error rather
 * than a silent "scrollToBlock retries 20 times and gives up".
 *
 * Three id shapes are allowed:
 * - Word's `w14:paraId` (or any non-empty allocator-generated id),
 *   surfaced verbatim when the source paragraph carries one.
 * - A zero-padded `seq-NNNN` fallback derived from document order,
 *   used when the paragraph has no paraId or the paraId collides
 *   with one already taken in the same derivation pass.
 * - A zero-padded `blank-NNNN` fallback for a paragraph that holds no
 *   text, numbered in its own sequence.
 *
 * The two sequences are separate on purpose. `seq-NNNN` counts the
 * paragraphs that carry text, and that count is the published
 * contract: the server extractor derives the same numbers, and a
 * citation stored against one has to keep naming the same paragraph.
 * Numbering blank paragraphs into the same sequence would renumber
 * every stored citation after the first blank line in a document.
 */

const SEQUENTIAL_BLOCK_ID_PREFIX = "seq-";
const BLANK_BLOCK_ID_PREFIX = "blank-";
const SEQUENTIAL_BLOCK_ID_PADDING = 4;
const SEQUENTIAL_BLOCK_ID_PATTERN = /^seq-\d{4,}$/u;
const BLANK_BLOCK_ID_PATTERN = /^blank-\d{4,}$/u;

export type FolioBlockId = string & { readonly __brand: "folio.blockId" };

export type DeriveBlockIdInput = {
  /**
   * Source paragraph's `w14:paraId` (or any equivalent stable id),
   * or `null` when the paragraph has none.
   */
  paraId: string | null;
  /** 1-based document order for the paragraph being derived. */
  index: number;
  /**
   * Ids already minted in the same derivation pass. Used to bump
   * the sequential fallback past any collisions (with the source
   * paraId set OR with previous fallbacks).
   */
  taken: ReadonlySet<string>;
};

const formatWithPrefix = (prefix: string, index: number): string =>
  `${prefix}${String(index).padStart(SEQUENTIAL_BLOCK_ID_PADDING, "0")}`;

const formatSequentialBlockId = (index: number): string =>
  formatWithPrefix(SEQUENTIAL_BLOCK_ID_PREFIX, index);

/**
 * The opaque `FolioBlockId` brand has no constructor; an unchecked
 * cast is the runtime no-op that mints one. Centralising the cast
 * here means every other site can keep `typescript/no-unsafe-type-assertion`
 * on.
 */
const brand = (value: string): FolioBlockId => value as unknown as FolioBlockId;

/**
 * Whether a source id would be read back as one this module generated.
 *
 * `seq-` and `blank-` are reserved: `getFolioParaIdFromBlockId` reports an id
 * in either shape as having no source paragraph, and `isFolioBlockId` holds
 * both to their exact pattern. A `w14:paraId` that happened to look like one
 * would therefore be told apart from itself, so it is not used verbatim and
 * the paragraph takes a generated id instead. (A real `w14:paraId` is eight
 * hex digits, so this costs nothing on any package Word writes; it is here so
 * a hand-written or synthesized source cannot break the round trip.)
 */
const isReservedShape = (paraId: string): boolean =>
  paraId.startsWith(SEQUENTIAL_BLOCK_ID_PREFIX) || paraId.startsWith(BLANK_BLOCK_ID_PREFIX);

const usableParaId = (paraId: string | null, taken: ReadonlySet<string>): string | null =>
  paraId !== null && paraId.length > 0 && !taken.has(paraId) && !isReservedShape(paraId)
    ? paraId
    : null;

export const deriveBlockId = ({ paraId, index, taken }: DeriveBlockIdInput): FolioBlockId => {
  const usable = usableParaId(paraId, taken);
  if (usable !== null) {
    return brand(usable);
  }
  let candidate = index;
  let formatted = formatSequentialBlockId(candidate);
  while (taken.has(formatted)) {
    candidate += 1;
    formatted = formatSequentialBlockId(candidate);
  }
  return brand(formatted);
};

/**
 * The id for a paragraph that holds no text, when it carries no
 * paraId to be named by. Numbered in its own sequence so the
 * published `seq-NNNN` positions stay where they are.
 */
export const deriveBlankBlockId = ({ paraId, index, taken }: DeriveBlockIdInput): FolioBlockId => {
  const usable = usableParaId(paraId, taken);
  if (usable !== null) {
    return brand(usable);
  }
  let candidate = index;
  let formatted = formatWithPrefix(BLANK_BLOCK_ID_PREFIX, candidate);
  while (taken.has(formatted)) {
    candidate += 1;
    formatted = formatWithPrefix(BLANK_BLOCK_ID_PREFIX, candidate);
  }
  return brand(formatted);
};

export const isSequentialFolioBlockId = (id: string): boolean =>
  SEQUENTIAL_BLOCK_ID_PATTERN.test(id);

export const isBlankFolioBlockId = (id: string): boolean => BLANK_BLOCK_ID_PATTERN.test(id);

/**
 * Extract the paraId an id was derived from, if any. Returns `null`
 * for sequential fallbacks. Accepts plain `string` so consumers that
 * have a raw id (snapshot anchors, operation blockIds) can call it
 * without an upfront brand check.
 */
export const getFolioParaIdFromBlockId = (id: string): string | null =>
  isSequentialFolioBlockId(id) || isBlankFolioBlockId(id) ? null : id;

/**
 * The 1-based document position encoded in a sequential fallback id
 * (`seq-0011` -> 11), or `null` for paraId-backed ids. The number is
 * the block's position in the same non-empty-block walk both the
 * server DOCX extractor and `createFolioAIEditSnapshot` use, so it
 * indexes a snapshot's ordered `blocks` array directly.
 */
export const getSequentialFolioBlockIdIndex = (id: string): number | null => {
  if (!SEQUENTIAL_BLOCK_ID_PATTERN.test(id)) {
    return null;
  }
  const index = Number(id.slice(SEQUENTIAL_BLOCK_ID_PREFIX.length));
  return Number.isInteger(index) && index > 0 ? index : null;
};

/**
 * Runtime refinement for ids coming back from the DB / API. Accepts
 * the same two shapes {@link deriveBlockId} produces and nothing
 * else — so legacy `b-NNNN` rows stop counting as valid here.
 */
export const isFolioBlockId = (value: unknown): value is FolioBlockId => {
  if (typeof value !== "string" || value.length === 0) {
    return false;
  }
  // Malformed sequential ids ("seq-", "seq-abc") are rejected so
  // typos can't pass as ids. The paraId arm is intentionally lenient
  // (any non-empty non-seq string) to match what real DOCX sources
  // and test fixtures put in `w14:paraId` — structural divergence is
  // already prevented by routing every mint through {@link deriveBlockId},
  // not by an exhaustive format check here.
  if (value.startsWith(SEQUENTIAL_BLOCK_ID_PREFIX)) {
    return SEQUENTIAL_BLOCK_ID_PATTERN.test(value);
  }
  if (value.startsWith(BLANK_BLOCK_ID_PREFIX)) {
    return BLANK_BLOCK_ID_PATTERN.test(value);
  }
  return true;
};
