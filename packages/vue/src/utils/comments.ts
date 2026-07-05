/**
 * Framework-agnostic comment + tracked-change helpers shared by adapters. The
 * data shapes and string-formatting rules here are part of the visible UI
 * (avatar colors, date strings) so keep this file as the single source of
 * truth. CSS-property factories live in adapter-specific files.
 * @packageDocumentation
 * @public
 */
import type { Paragraph } from "@stll/folio-core/types/content";

/** Extract plain text from a Comment's paragraph content. */
export function getCommentText(paragraphs?: Paragraph[]): string {
  if (!paragraphs?.length) return "";
  return paragraphs
    .flatMap((p) =>
      p.content
        .filter((c) => c.type === "run")
        .flatMap((r) => ("content" in r ? r.content : []))
        .filter((c) => c.type === "text")
        .map((t) => ("text" in t ? t.text : "")),
    )
    .join("");
}

export function formatDate(dateStr?: string): string {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return d.toLocaleString(undefined, {
    hour: "numeric",
    minute: "2-digit",
    month: "short",
    day: "numeric",
  });
}

export function getInitials(name: string): string {
  return name
    .split(/\s+/)
    .map((w) => w[0])
    .join("")
    .toUpperCase()
    .slice(0, 2);
}

// Kibana-style avatar palette — deterministic per author name.
const AVATAR_COLORS = [
  "#6DCCB1",
  "#79AAD9",
  "#EE789D",
  "#A987D1",
  "#E6A85F",
  "#F2CC8F",
  "#68B3A2",
  "#B07AA1",
  "#59A14F",
  "#FF9DA7",
  "#E15759",
  "#76B7B2",
];

export function getAvatarColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length];
}

export function truncateText(text: string, maxLength = 50): string {
  return text.length > maxLength ? text.slice(0, maxLength) + "..." : text;
}

/**
 * One tracked change surfaced by `extractTrackedChanges`. Each entry
 * groups all sites of one revision into a single row that the sidebar
 * renders as one card. Resolve via the by-id resolver for any type — it walks
 * every site sharing the id so coalesced edits clear in one click.
 *
 * @public
 */
export interface TrackedChangeEntry {
  /**
   * Revision shape. Inline shapes (`insertion`, `deletion`, `replacement`)
   * wrap text runs; the rest are structural revisions on node attrs.
   *
   * - `insertion` — text was added (`<w:ins>`).
   * - `deletion` — text was struck through but not removed (`<w:del>`).
   * - `replacement` — a deletion + insertion by the same author at the
   *   same position+time; sidebar shows one combined card. `deletedText`
   *   and `insertionRevisionId` are set on this variant.
   * - `paragraphMarkInsertion` / `paragraphMarkDeletion` — Enter /
   *   Backspace produced a tracked paragraph break (`<w:pPr><w:rPr><w:ins/>` /
   *   `<w:del/>`).
   * - `paragraphPropertiesChanged` — formatting (alignment, spacing,
   *   etc.) on the paragraph was changed (`<w:pPrChange>`).
   * - `rowInserted` / `rowDeleted` / `rowPropertiesChanged` — table
   *   row authored / removed / formatted (`<w:trPr><w:ins/>` / `<w:del/>`
   *   / `<w:trPrChange>`).
   * - `cellInserted` / `cellDeleted` / `cellMerged` /
   *   `cellPropertiesChanged` — per-cell revisions
   *   (`<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>` / `<w:tcPrChange>`).
   * - `tablePropertiesChanged` — table-level formatting
   *   (`<w:tblPrChange>`).
   */
  type:
    | "insertion"
    | "deletion"
    | "replacement"
    | "paragraphMarkInsertion"
    | "paragraphMarkDeletion"
    | "paragraphPropertiesChanged"
    | "rowInserted"
    | "rowDeleted"
    | "rowPropertiesChanged"
    | "cellInserted"
    | "cellDeleted"
    | "cellMerged"
    | "cellPropertiesChanged"
    | "tableInserted"
    | "tableDeleted"
    | "tablePropertiesChanged";
  /**
   * Affected text. For inline types this is the run's text; for
   * structural types it's the surrounding paragraph / cell content
   * (truncated by the sidebar before display).
   */
  text: string;
  /**
   * Only set when `type === 'replacement'` — the text the user removed.
   * The inserted text lives in {@link TrackedChangeEntry.text}.
   */
  deletedText?: string;
  /** Author that minted the revision (`w:author`). */
  author: string;
  /** ISO timestamp the revision was minted (`w:date`). May be undefined for legacy imports. */
  date?: string;
  /**
   * Document position where the revision starts. For inline types this
   * is the start of the marked text run; for structural types it's the
   * containing paragraph / row / cell / table node's start position.
   * Used by the sidebar to anchor the card at the correct vertical
   * offset.
   */
  from: number;
  /**
   * Document position where the revision ends. For inline coalesced
   * runs that span multiple paragraphs, this is the END position of the
   * LAST run in the group; the intervening structural positions are not
   * preserved.
   */
  to: number;
  /**
   * The `w:id` of the revision. Pass to the by-id resolver to resolve every
   * site sharing this id — including pPrIns paragraph attrs and
   * subsequent typed runs in the same editing session.
   */
  revisionId: number;
  /**
   * Only set when `type === 'replacement'` — the insertion half carries
   * a DIFFERENT `w:id` from the deletion (sharing would trip the OOXML
   * move-pair serializer). Card Accept handlers dispatch BOTH ids to
   * clear the deletion and the insertion + any coalesced paragraph-marks.
   */
  insertionRevisionId?: number;
  /**
   * Extra `w:id`s that map to the same logical revision as this card.
   * Populated when the extractor coalesces a burst of distinct ids by
   * (author, date) — e.g. a foreign document where the source editor
   * minted a fresh id per atomic edit. Accept/reject handlers must
   * resolve every id in this list in addition to {@link revisionId}.
   */
  coalescedRevisionIds?: number[];
}
