/**
 * Comment range markers a table keeps between its rows, or a row between its
 * cells.
 *
 * Neither container has a node for one, so the parser keeps the marker as
 * positioned markup and the serializer writes it back where it stood. A
 * comment's range that runs to the end of a table is written this way: the end
 * is a child of `w:tbl` (or `w:tr`) while the start and the
 * `w:commentReference` sit in cell paragraphs. Every pass that pairs a range's
 * start with its end reads these too, or the range looks unbalanced and its
 * paragraph-side markers get repaired away. The model validator in docx-core
 * counts them the same way.
 */

import type { PreservedChild } from "@stll/docx-core/model";

export type PreservedCommentRangeMarker = {
  readonly type: "commentRangeStart" | "commentRangeEnd";
  readonly id: number;
};

const PRESERVED_COMMENT_RANGE_MARKER =
  /<(?:[\w.-]+:)?commentRange(Start|End)(?=[\s/>])[^>]*?\s(?:[\w.-]+:)?id=["'](\d+)["']/gu;

/** The comment range markers one positioned child holds, in source order. */
export const preservedCommentRangeMarkers = (
  child: Pick<PreservedChild, "xml">,
): PreservedCommentRangeMarker[] =>
  [...child.xml.matchAll(PRESERVED_COMMENT_RANGE_MARKER)].map((match) => ({
    type: match[1] === "Start" ? "commentRangeStart" : "commentRangeEnd",
    id: Number(match[2]),
  }));

/** The markers of the children positioned at `at` (before the row or cell it names). */
export const preservedCommentRangeMarkersAt = (
  children: readonly PreservedChild[] | undefined,
  at: number,
): PreservedCommentRangeMarker[] =>
  (children ?? []).flatMap((child) =>
    child.index === at ? preservedCommentRangeMarkers(child) : [],
  );
