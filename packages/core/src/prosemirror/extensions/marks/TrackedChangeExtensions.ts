/**
 * Tracked Change Mark Extensions — insertion and deletion marks
 *
 * Renders insertions with colored underline and deletions with colored
 * strikethrough, matching the standard MS Word display for tracked changes.
 * Colors are assigned per author via CSS custom properties (see editor.css).
 */

import { getAuthorColorIdx, AUTHOR_COLORS } from "../../../utils/authorColors";
import { expectTrackedChangeMarkAttrs } from "../../attrs";
import { createMarkExtension } from "../create";

/**
 * Build an inline style string for the tracked change author color.
 * ProseMirror toDOM runs before the layout painter, so we set the
 * decoration here for correct rendering in both the layout-painter
 * path and the ProseMirror DOM path.
 */
const insertionStyle = (color: string): string =>
  `color: ${color}; text-decoration: underline; text-decoration-color: ${color};`;

const deletionStyle = (color: string): string =>
  `color: ${color}; text-decoration: line-through; text-decoration-color: ${color};`;

/**
 * Suggested (AI-proposed) tracked changes reuse the insertion/deletion visual
 * grammar but with a DOTTED stroke and a dedicated suggestion hue, kept
 * distinct from the per-author redline colors. The hue and tint are driven by
 * the `--suggestion-color` / `--suggestion-bg` CSS custom properties (see
 * editor.css); the inline fallbacks below keep the paged-canvas projection —
 * which reuses this same `toDOM` style string — legible when the stylesheet
 * variables are not in scope.
 */
const SUGGESTION_COLOR = "var(--suggestion-color, #6d3bd6)";
const SUGGESTION_TINT = "var(--suggestion-bg, color-mix(in oklch, #6d3bd6 12%, transparent))";
// Layered as a translucent background-image (not background-color) so authored
// highlight/shading beneath the suggestion stays visible under the tint.
const SUGGESTION_TINT_LAYER = `linear-gradient(${SUGGESTION_TINT}, ${SUGGESTION_TINT})`;

const suggestedInsertionStyle = (): string =>
  `color: ${SUGGESTION_COLOR}; text-decoration: underline; text-decoration-style: dotted; text-decoration-color: ${SUGGESTION_COLOR}; background-image: ${SUGGESTION_TINT_LAYER};`;

const suggestedDeletionStyle = (): string =>
  `color: ${SUGGESTION_COLOR}; text-decoration: line-through; text-decoration-style: dotted; text-decoration-color: ${SUGGESTION_COLOR}; background-image: ${SUGGESTION_TINT_LAYER};`;

/**
 * Insertion mark — text added in tracked changes
 * Renders with per-author colored underline.
 */
export const InsertionExtension = createMarkExtension({
  name: "insertion",
  schemaMarkName: "insertion",
  markSpec: {
    attrs: {
      revisionId: { default: 0 },
      author: { default: "" },
      date: { default: null },
      // `"moveTo"` distinguishes inserted text that originated as a
      // `w:moveTo` (the destination half of an OOXML move) from a
      // plain `w:ins`. Carried through PM so `fromProseDoc` can
      // re-emit the correct OOXML element without relying on
      // brittle revisionId pairing across the doc.
      moveKind: { default: null },
      // Optional author initials (w:initials), carried for round-trip.
      initials: { default: null },
      // `"suggested"` marks are AI-proposed edits: rendered like tracked
      // changes but stripped from serialized DOCX until accepted. Parsing
      // DOCX/HTML never sets this (see parseDOM), so provenance defaults to
      // `"user"` and round-trips remain suggestion-free.
      provenance: { default: "user" },
      suggestionId: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: "span.docx-insertion",
        getAttrs(dom) {
          return {
            revisionId: Number.parseInt(dom.dataset["revisionId"] ?? "0", 10),
            author: dom.dataset["author"] ?? "",
            date: dom.dataset["date"] ?? null,
          };
        },
      },
    ],
    toDOM(mark) {
      const { revisionId, author, date, provenance, suggestionId } =
        expectTrackedChangeMarkAttrs(mark);
      const idx = getAuthorColorIdx(author);
      // SAFETY: getAuthorColorIdx returns modulo AUTHOR_COLORS.length
      const color = AUTHOR_COLORS[idx] ?? "#000000";
      const datePart = date ? new Date(date).toLocaleDateString() : "";
      const titleParts = [author, datePart].filter(Boolean);
      const suggested = provenance === "suggested";
      return [
        "span",
        {
          class: suggested ? "docx-insertion docx-insertion--suggested" : "docx-insertion",
          "data-revision-id": String(revisionId),
          "data-author": author,
          "data-tc-author-idx": String(idx),
          ...(suggested ? { "data-provenance": "suggested" } : {}),
          ...(suggested && suggestionId ? { "data-suggestion-id": suggestionId } : {}),
          ...(date ? { "data-date": date } : {}),
          ...(titleParts.length > 0
            ? { title: `${suggested ? "Suggested" : "Inserted"}: ${titleParts.join(", ")}` }
            : {}),
          style: suggested ? suggestedInsertionStyle() : insertionStyle(color),
        },
        0,
      ];
    },
  },
});

/**
 * Deletion mark — text removed in tracked changes
 * Renders with per-author colored strikethrough.
 */
export const DeletionExtension = createMarkExtension({
  name: "deletion",
  schemaMarkName: "deletion",
  markSpec: {
    attrs: {
      revisionId: { default: 0 },
      author: { default: "" },
      date: { default: null },
      // `"moveFrom"` distinguishes deleted text that originated as a
      // `w:moveFrom` (the source half of an OOXML move) from a plain
      // `w:del`. Carried through PM so `fromProseDoc` can re-emit
      // the correct OOXML element without relying on brittle
      // revisionId pairing across the doc.
      moveKind: { default: null },
      // Optional author initials (w:initials), carried for round-trip.
      initials: { default: null },
      // See the insertion mark: `"suggested"` is an AI proposal stripped
      // from serialized output until accepted.
      provenance: { default: "user" },
      suggestionId: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: "span.docx-deletion",
        getAttrs(dom) {
          return {
            revisionId: Number.parseInt(dom.dataset["revisionId"] ?? "0", 10),
            author: dom.dataset["author"] ?? "",
            date: dom.dataset["date"] ?? null,
          };
        },
      },
    ],
    toDOM(mark) {
      const { revisionId, author, date, provenance, suggestionId } =
        expectTrackedChangeMarkAttrs(mark);
      const idx = getAuthorColorIdx(author);
      // SAFETY: getAuthorColorIdx returns modulo AUTHOR_COLORS.length
      const color = AUTHOR_COLORS[idx] ?? "#000000";
      const datePart = date ? new Date(date).toLocaleDateString() : "";
      const titleParts = [author, datePart].filter(Boolean);
      const suggested = provenance === "suggested";
      return [
        "span",
        {
          class: suggested ? "docx-deletion docx-deletion--suggested" : "docx-deletion",
          "data-revision-id": String(revisionId),
          "data-author": author,
          "data-tc-author-idx": String(idx),
          ...(suggested ? { "data-provenance": "suggested" } : {}),
          ...(suggested && suggestionId ? { "data-suggestion-id": suggestionId } : {}),
          ...(date ? { "data-date": date } : {}),
          ...(titleParts.length > 0
            ? {
                title: `${suggested ? "Suggested deletion" : "Deleted"}: ${titleParts.join(", ")}`,
              }
            : {}),
          style: suggested ? suggestedDeletionStyle() : deletionStyle(color),
        },
        0,
      ];
    },
  },
});

/**
 * Run property change — formatting changed while review tracking was active.
 *
 * The current formatting remains represented by the normal formatting marks;
 * this mark carries the previous run properties and revision metadata needed
 * to serialize, list, accept, and reject the change.
 */
export const RunPropertyChangeExtension = createMarkExtension({
  name: "runPropertyChange",
  schemaMarkName: "runPropertyChange",
  markSpec: {
    attrs: {
      changes: { default: [] },
      // See the insertion/deletion marks: a `"suggested"` run-property change
      // is reverted to its original formatting on serialize until accepted.
      provenance: { default: "user" },
      suggestionId: { default: null },
    },
    inclusive: false,
    parseDOM: [{ tag: "span.docx-run-property-change" }],
    toDOM(mark) {
      const suggested = mark.attrs["provenance"] === "suggested";
      const suggestionId = mark.attrs["suggestionId"];
      return [
        "span",
        {
          class: suggested
            ? "docx-run-property-change docx-run-property-change--suggested"
            : "docx-run-property-change",
          ...(suggested ? { "data-provenance": "suggested" } : {}),
          ...(suggested && typeof suggestionId === "string"
            ? { "data-suggestion-id": suggestionId }
            : {}),
        },
        0,
      ];
    },
  },
});
