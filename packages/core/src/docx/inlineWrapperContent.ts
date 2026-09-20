/**
 * Which paragraph content an inline wrapper is allowed to keep inside itself.
 *
 * A revision wrapper (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) and an
 * inline content control (`w:sdt`) are both parsed by recursing into their
 * children and then re-wrapping what came back, and folio asks the question in
 * two places: the parser lifts out whatever the wrapper may not hold, and the
 * editor's save path filters the content it rebuilt. Lifting content out moves
 * it from inside the wrapper to outside it — a run lifted out of a `w:ins` is
 * no longer inserted, and accepting or rejecting the revision both keep it.
 *
 * So the admission is a fact about the content model, written here once and
 * bound to that model: `AdmissionMap<T>` demands `true` for exactly the types
 * `T` holds, so the map cannot say yes to something the union rejects, and the
 * union cannot gain a member the map still lifts out.
 */

import type { InlineSdt, ParagraphContent, TrackedRunContent } from "../types/document";

/** `true` for exactly the paragraph-content types `Admitted` holds. */
type AdmissionMap<Admitted extends ParagraphContent> = {
  [Type in ParagraphContent["type"]]: Type extends Admitted["type"] ? true : false;
};

/**
 * What a run-level tracked-change wrapper keeps inside itself (CT_RunTrackChange).
 *
 * Runs, links, fields, bookmark boundaries, equations, nested revisions and
 * markup folio does not model.
 *
 * The transparent wrappers are admitted too: `w:bdo` / `w:dir` state how their
 * content is laid out, and `w:sdt` states what the content is bound to.
 * Neither says anything about the revision, so lifting one out of a revision
 * takes its text out of the revision with it.
 *
 * The range markers are not: a `w:commentRangeStart` or `w:moveFromRangeStart`
 * inside a revision is a marker the revision does not own, and the pairing
 * passes read it as a paragraph-level sibling.
 */
export const TRACKED_CHANGE_WRAPPER_CONTENT = {
  inlineWrapper: true,
  bookmarkEnd: true,
  bookmarkStart: true,
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
  commentRangeEnd: false,
  commentRangeStart: false,
  commentReference: false,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveToRangeEnd: false,
  moveToRangeStart: false,
} as const satisfies AdmissionMap<TrackedRunContent>;

/**
 * What an inline content control keeps inside `<w:sdtContent>`.
 *
 * Mirror of upstream eigenpal/docx-editor PR #482 (commit 29f95751d), plus
 * the bidirectional wrapper: OOXML allows runs, hyperlinks, simple/complex
 * fields, nested SDTs, tracked insertions/deletions/moves, math, markup folio
 * does not model, and the bidirectional controls directly inside
 * `<w:sdtContent>`.
 *
 * The bookmark boundaries are admitted for the reason the revision wrapper
 * admits them: `CT_SdtContentRun` reaches `EG_RangeMarkupElements` through
 * `EG_RunLevelElts`, so the marker is valid where Word wrote it, and lifting
 * it out changes what the bookmark covers. A range whose extent was the
 * control's content comes back covering the control plus whatever follows, so
 * a `REF` field or a link to it resolves to the wrong text.
 *
 * The other range markers are not: a `w:commentRangeStart` or a
 * `w:moveFromRangeStart` inside the control is a marker the control does not
 * own, and the pairing passes read it as a paragraph-level sibling.
 */
export const INLINE_SDT_CONTENT = {
  inlineWrapper: true,
  bookmarkEnd: true,
  bookmarkStart: true,
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
  commentRangeEnd: false,
  commentRangeStart: false,
  commentReference: false,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveToRangeEnd: false,
  moveToRangeStart: false,
} as const satisfies AdmissionMap<InlineSdt["content"][number]>;

/**
 * Whether a revision wrapper keeps this content inside itself.
 *
 * Generic in the input so a caller that has already narrowed its content
 * keeps that narrowing: filtering a list the bidirectional wrappers are
 * already flattened out of must not put one back into the result type.
 */
export const isTrackedChangeWrapperChild = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, TrackedRunContent> => TRACKED_CHANGE_WRAPPER_CONTENT[content.type];

/** Whether an inline content control keeps this content inside `<w:sdtContent>`. */
export const isInlineSdtContent = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, InlineSdt["content"][number]> => INLINE_SDT_CONTENT[content.type];
