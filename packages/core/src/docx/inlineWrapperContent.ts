/**
 * Which paragraph content a container that holds inline content admits.
 *
 * A revision wrapper (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`), an inline
 * content control (`w:sdt`), a link and a simple field are all parsed by
 * recursing into their children and then re-wrapping what came back, and folio
 * asks the question in two places: the parser lifts out whatever the container
 * may not hold, and the editor's save path filters the content it rebuilt.
 * Lifting content out moves it from inside the container to outside it — a run
 * lifted out of a `w:ins` is no longer inserted, and accepting or rejecting the
 * revision both keep it.
 *
 * So the admission is a fact about the content model, written here once and
 * bound to that model: `AdmissionMap<T>` demands `true` for exactly the types
 * `T` holds, so the map cannot say yes to something the union rejects, and the
 * union cannot gain a member the map still lifts out.
 */

import type {
  Hyperlink,
  InlineSdt,
  ParagraphContent,
  SimpleField,
  TrackedRunContent,
} from "../types/document";

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
 * `<w:sdtContent>`. Bookmarks and range markers are lifted out as siblings of
 * the SDT so the control itself stays valid.
 */
export const INLINE_SDT_CONTENT = {
  inlineWrapper: true,
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
  bookmarkEnd: false,
  bookmarkStart: false,
  commentRangeEnd: false,
  commentRangeStart: false,
  commentReference: false,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveToRangeEnd: false,
  moveToRangeStart: false,
} as const satisfies AdmissionMap<InlineSdt["content"][number]>;

/**
 * What a `w:hyperlink` keeps inside itself (CT_Hyperlink).
 *
 * Runs, the bookmark boundaries a link's target sits on, markup folio does not
 * model, and the four transparent wrappers: `CT_Hyperlink` is `EG_PContent`,
 * which declares `w:bdo`, `w:dir`, `w:smartTag` and the run-level
 * `w:customXml`, so `w:hyperlink > w:bdo > w:r` is a link an author may write
 * and folio reads it as authored.
 *
 * A revision is not admitted, and the exclusion is the model's, not the
 * schema's: OOXML nests `w:ins` inside `w:hyperlink` and folio nests the link
 * inside the revision, so the paragraph parser hoists one around the link.
 */
export const HYPERLINK_CONTENT = {
  bookmarkEnd: true,
  bookmarkStart: true,
  inlineWrapper: true,
  preservedInline: true,
  run: true,
  commentRangeEnd: false,
  commentRangeStart: false,
  commentReference: false,
  complexField: false,
  deletion: false,
  hyperlink: false,
  inlineSdt: false,
  insertion: false,
  mathEquation: false,
  moveFrom: false,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveTo: false,
  moveToRangeEnd: false,
  moveToRangeStart: false,
  simpleField: false,
} as const satisfies AdmissionMap<Hyperlink["children"][number]>;

/**
 * What a `w:fldSimple` keeps inside itself (CT_SimpleField).
 *
 * The cached result's runs and the link one may carry, the same four
 * transparent wrappers `EG_PContent` declares, and markup folio does not
 * model. `w:fldData` is the field's own child rather than its content, so it
 * rides `dispatchChildren`'s sink like any other capture.
 */
export const SIMPLE_FIELD_CONTENT = {
  hyperlink: true,
  inlineWrapper: true,
  preservedInline: true,
  run: true,
  bookmarkEnd: false,
  bookmarkStart: false,
  commentRangeEnd: false,
  commentRangeStart: false,
  commentReference: false,
  complexField: false,
  deletion: false,
  inlineSdt: false,
  insertion: false,
  mathEquation: false,
  moveFrom: false,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveTo: false,
  moveToRangeEnd: false,
  moveToRangeStart: false,
  simpleField: false,
} as const satisfies AdmissionMap<SimpleField["content"][number]>;

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

/** Whether a `w:hyperlink` keeps this content inside itself. */
export const isHyperlinkContent = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, Hyperlink["children"][number]> => HYPERLINK_CONTENT[content.type];

/** Whether a `w:fldSimple` keeps this content inside itself. */
export const isSimpleFieldContent = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, SimpleField["content"][number]> =>
  SIMPLE_FIELD_CONTENT[content.type];
