/**
 * What an inline wrapper may hold, decided once per paragraph-content member.
 *
 * A run-level tracked change and an inline content control each admit a subset
 * of paragraph content, and folio asks the question in two places: the parser
 * lifts out whatever the wrapper may not hold, and the editor's save path
 * filters the content it rebuilt. Both were hand-written lists of `type ===`
 * comparisons, each a mirror of the model's union, and the second even carried
 * a comment asking the next person to keep it in sync. A mirror drifts: a
 * member added to the union and to one list is a member the other list drops
 * silently, which is the wrapper losing content on the way out of the editor.
 *
 * So the membership is one total map per wrapper. `satisfies
 * Record<ParagraphContent["type"], boolean>` makes the compiler refuse a map
 * that has not decided about a new member, and both callers read the map.
 */

import type { InlineSdt, ParagraphContent, TrackedRunChange } from "../types/document";

/**
 * `CT_RunTrackChange` holds run-level content: runs, links, fields, bookmark
 * boundaries, equations and nested revisions, plus markup folio does not
 * model. Everything else the paragraph parser produced — comment and move
 * ranges, bidirectional wrappers, an inline content control — is lifted out
 * as a sibling of the wrapper, because the wrapper would not be valid holding
 * it.
 */
const TRACKED_CHANGE_WRAPPER_CONTENT = {
  bookmarkEnd: true,
  bookmarkStart: true,
  complexField: true,
  deletion: true,
  hyperlink: true,
  insertion: true,
  mathEquation: true,
  moveFrom: true,
  moveTo: true,
  preservedInline: true,
  run: true,
  simpleField: true,
  bidiWrapper: false,
  commentRangeEnd: false,
  commentRangeStart: false,
  commentReference: false,
  inlineSdt: false,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveToRangeEnd: false,
  moveToRangeStart: false,
} as const satisfies Record<ParagraphContent["type"], boolean>;

/**
 * `CT_SdtContentRun` holds runs, links, fields, nested controls, revisions,
 * equations and markup folio does not model. Bookmarks and range markers are
 * lifted out as siblings so the control itself stays valid.
 *
 * Mirror of upstream eigenpal/docx-editor PR #482 (commit 29f95751d).
 */
const INLINE_SDT_CONTENT = {
  complexField: true,
  deletion: true,
  hyperlink: true,
  inlineSdt: true,
  insertion: true,
  mathEquation: true,
  moveFrom: true,
  moveTo: true,
  preservedInline: true,
  run: true,
  simpleField: true,
  bidiWrapper: false,
  bookmarkEnd: false,
  bookmarkStart: false,
  commentRangeEnd: false,
  commentRangeStart: false,
  commentReference: false,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveToRangeEnd: false,
  moveToRangeStart: false,
} as const satisfies Record<ParagraphContent["type"], boolean>;

export const isTrackedChangeWrapperChild = (
  content: ParagraphContent,
): content is TrackedRunChange["content"][number] => TRACKED_CHANGE_WRAPPER_CONTENT[content.type];

export const isInlineSdtContent = (
  content: ParagraphContent,
): content is InlineSdt["content"][number] => INLINE_SDT_CONTENT[content.type];
