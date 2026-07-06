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
  // Modulo the non-empty palette length always yields a valid index; the `??`
  // literal satisfies noUncheckedIndexedAccess without a cast.
  return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length] ?? "#888888";
}

export function truncateText(text: string, maxLength = 50): string {
  return text.length > maxLength ? text.slice(0, maxLength) + "..." : text;
}

/**
 * One tracked change surfaced by `extractTrackedChanges`. Re-exported verbatim
 * from core so the sidebar consumes the extractor's exact output type: the Vue
 * adapter used to keep a stricter hand-written copy (required where core's is
 * optional, `?: T` where core's is `?: T | undefined`), which made
 * `extractTrackedChanges(state).entries` unassignable to the sidebar props under
 * `exactOptionalPropertyTypes`. Sharing core's type removes that drift entirely
 * and keeps the accept/reject-by-id contract (`revisionId`,
 * `insertionRevisionId`, `coalescedRevisionIds`) in one place.
 *
 * @public
 */
export type { TrackedChangeEntry } from "@stll/folio-core/prosemirror/utils/extractTrackedChanges";
