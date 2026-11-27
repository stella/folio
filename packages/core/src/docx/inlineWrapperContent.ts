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

import { panic } from "better-result";

import type {
  Hyperlink,
  InlineSdt,
  InlineWrapper,
  ParagraphContent,
  PreservedInline,
  Run,
  SimpleField,
  TrackedRunChange,
  TrackedRunContent,
} from "../types/document";

/** `true` for exactly the paragraph-content types `Admitted` holds. */
type AdmissionMap<Admitted extends ParagraphContent> = {
  [Type in ParagraphContent["type"]]: Type extends Admitted["type"]
    ? true
    : false;
};

/**
 * What a run-level tracked-change wrapper keeps inside itself (CT_RunTrackChange).
 *
 * Runs, links, fields, bookmark and move-range boundaries, equations, nested
 * revisions and markup folio does not model.
 *
 * The transparent wrappers are admitted too: `w:bdo` / `w:dir` state how their
 * content is laid out, and `w:sdt` states what the content is bound to.
 * Neither says anything about the revision, so lifting one out of a revision
 * takes its text out of the revision with it.
 *
 * Range markers stay inside the revision that owns their authored position.
 * Comment range markers are positions even though the editor projects their
 * extent as marks; lifting a marker changes which side of the tracked content
 * owns that position when the revision is resolved.
 */
export const TRACKED_CHANGE_WRAPPER_CONTENT: {
  readonly inlineWrapper: true;
  readonly bookmarkEnd: true;
  readonly bookmarkStart: true;
  readonly complexField: true;
  readonly deletion: true;
  readonly hyperlink: true;
  readonly inlineSdt: true;
  readonly insertion: true;
  readonly mathEquation: true;
  readonly moveFrom: true;
  readonly moveTo: true;
  readonly preservedInline: true;
  readonly run: true;
  readonly simpleField: true;
  readonly commentRangeEnd: true;
  readonly commentRangeStart: true;
  // The model lifts w:commentReference out of its w:r carrier. Keeping the
  // record here writes that carrier within CT_RunTrackChange at the same
  // position.
  readonly commentReference: true;
  readonly moveFromRangeEnd: true;
  readonly moveFromRangeStart: true;
  readonly moveToRangeEnd: true;
  readonly moveToRangeStart: true;
} = {
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
  commentRangeEnd: true,
  commentRangeStart: true,
  // The model lifts w:commentReference out of its w:r carrier. Keeping the
  // record here writes that carrier within CT_RunTrackChange at the same
  // position.
  commentReference: true,
  moveFromRangeEnd: true,
  moveFromRangeStart: true,
  moveToRangeEnd: true,
  moveToRangeStart: true,
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
 * The bookmark and comment boundaries are admitted for the reason the
 * revision wrapper admits them: `CT_SdtContentRun` reaches
 * `EG_RangeMarkupElements` through
 * `EG_RunLevelElts`, so the marker is valid where Word wrote it, and lifting
 * it out changes what the bookmark covers. A range whose extent was the
 * control's content comes back covering the control plus whatever follows, so
 * a `REF` field or a link to it resolves to the wrong text.
 *
 * Point comment references are lifted from their required run carrier and
 * rebuilt as a run by the serializer; they are not emitted as direct
 * `<w:sdtContent>` children. Move-range boundaries stay where the source put
 * them. The editor has zero-width boundary nodes for them, so lifting one
 * beside the control would change which content the tracked move covers
 * without any representational need. Comment range boundaries are
 * paragraph-position markers in the editor, but `CT_SdtContentRun` also
 * admits them as direct children. Keeping them here preserves the authored
 * container when neighboring wrappers collapse.
 */
export const INLINE_SDT_CONTENT: {
  readonly inlineWrapper: true;
  readonly bookmarkEnd: true;
  readonly bookmarkStart: true;
  readonly complexField: true;
  readonly deletion: true;
  readonly hyperlink: true;
  readonly inlineSdt: true;
  readonly insertion: true;
  readonly mathEquation: true;
  readonly moveFrom: true;
  readonly moveTo: true;
  readonly preservedInline: true;
  readonly run: true;
  readonly simpleField: true;
  readonly commentRangeEnd: true;
  readonly commentRangeStart: true;
  readonly commentReference: true;
  readonly moveFromRangeEnd: true;
  readonly moveFromRangeStart: true;
  readonly moveToRangeEnd: true;
  readonly moveToRangeStart: true;
} = {
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
  commentRangeEnd: true,
  commentRangeStart: true,
  commentReference: true,
  moveFromRangeEnd: true,
  moveFromRangeStart: true,
  moveToRangeEnd: true,
  moveToRangeStart: true,
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
export const HYPERLINK_CONTENT: {
  readonly bookmarkEnd: true;
  readonly bookmarkStart: true;
  readonly commentRangeEnd: true;
  readonly commentRangeStart: true;
  readonly commentReference: true;
  readonly inlineWrapper: true;
  readonly preservedInline: true;
  readonly run: true;
  readonly complexField: true;
  readonly deletion: false;
  readonly hyperlink: false;
  readonly inlineSdt: false;
  readonly insertion: false;
  readonly mathEquation: false;
  readonly moveFrom: false;
  readonly moveFromRangeEnd: true;
  readonly moveFromRangeStart: true;
  readonly moveTo: false;
  readonly moveToRangeEnd: true;
  readonly moveToRangeStart: true;
  readonly simpleField: false;
} = {
  bookmarkEnd: true,
  bookmarkStart: true,
  commentRangeEnd: true,
  commentRangeStart: true,
  commentReference: true,
  inlineWrapper: true,
  preservedInline: true,
  run: true,
  complexField: true,
  deletion: false,
  hyperlink: false,
  inlineSdt: false,
  insertion: false,
  mathEquation: false,
  moveFrom: false,
  moveFromRangeEnd: true,
  moveFromRangeStart: true,
  moveTo: false,
  moveToRangeEnd: true,
  moveToRangeStart: true,
  simpleField: false,
} as const satisfies AdmissionMap<Hyperlink["children"][number]>;

/**
 * What the parser reads as a `w:fldSimple`'s own typed children
 * (CT_SimpleField).
 *
 * The cached result's runs and the link one may carry, the same four
 * transparent wrappers `EG_PContent` declares, the tracked changes around
 * them, and markup folio does not model. Every other child it keeps as the
 * markup it is. `w:fldData` is the field's own child rather than its
 * content, so it rides `dispatchChildren`'s sink like any other capture.
 *
 * The model's field holds more ({@link isSimpleFieldContent}): what a
 * tracked change inside it holds, which resolving the change leaves there.
 */
export const SIMPLE_FIELD_READ_CONTENT: {
  readonly hyperlink: true;
  readonly inlineWrapper: true;
  readonly preservedInline: true;
  readonly run: true;
  readonly bookmarkEnd: false;
  readonly bookmarkStart: false;
  readonly commentRangeEnd: false;
  readonly commentRangeStart: false;
  readonly commentReference: false;
  readonly complexField: false;
  readonly deletion: true;
  readonly inlineSdt: false;
  readonly insertion: true;
  readonly mathEquation: false;
  readonly moveFrom: true;
  readonly moveFromRangeEnd: false;
  readonly moveFromRangeStart: false;
  readonly moveTo: true;
  readonly moveToRangeEnd: false;
  readonly moveToRangeStart: false;
  readonly simpleField: false;
} = {
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
  deletion: true,
  inlineSdt: false,
  insertion: true,
  mathEquation: false,
  moveFrom: true,
  moveFromRangeEnd: false,
  moveFromRangeStart: false,
  moveTo: true,
  moveToRangeEnd: false,
  moveToRangeStart: false,
  simpleField: false,
} as const satisfies AdmissionMap<
  Run | Hyperlink | InlineWrapper | PreservedInline | TrackedRunChange
>;

/**
 * Every paragraph content kind is a tracked change child. The reader keeps a
 * tracked change whole on this: its resolution joins describe the one wrapper.
 * A model kind outside `TrackedRunContent` turns this type to `never`, so the
 * build fails until reading such a change is designed.
 */
export const EVERY_PARAGRAPH_CONTENT_IS_TRACKED_CHILD: (typeof TRACKED_CHANGE_WRAPPER_CONTENT)[ParagraphContent["type"]] extends true
  ? true
  : never = true;

/**
 * Whether a revision wrapper keeps this content inside itself.
 *
 * Generic in the input so a caller that has already narrowed its content
 * keeps that narrowing: filtering a list the bidirectional wrappers are
 * already flattened out of must not put one back into the result type.
 */
export const isTrackedChangeWrapperChild = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, TrackedRunContent> =>
  TRACKED_CHANGE_WRAPPER_CONTENT[content.type];

/** Whether an inline content control keeps this content inside `<w:sdtContent>`. */
export const isInlineSdtContent = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, InlineSdt["content"][number]> =>
  INLINE_SDT_CONTENT[content.type];

/** Whether a `w:hyperlink` keeps this content inside itself. */
export const isHyperlinkContent = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, Hyperlink["children"][number]> =>
  HYPERLINK_CONTENT[content.type];

/**
 * Whether a `w:fldSimple` keeps this content inside itself: what a tracked
 * change inside it holds, so resolving one leaves its content in the field.
 */
export const isSimpleFieldContent = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<Content, SimpleField["content"][number]> =>
  TRACKED_CHANGE_WRAPPER_CONTENT[content.type];

/**
 * Whether the parser reads this content as a `w:fldSimple`'s own typed child
 * ({@link SIMPLE_FIELD_READ_CONTENT}) rather than the markup it keeps.
 */
export const isSimpleFieldReadContent = <Content extends ParagraphContent>(
  content: Content,
): content is Extract<
  Content,
  Run | Hyperlink | InlineWrapper | PreservedInline | TrackedRunChange
> => SIMPLE_FIELD_READ_CONTENT[content.type];

/**
 * Whether a transparent wrapper holds, at any depth, content a link cannot
 * hold. The model types a wrapper's content as any paragraph content, so a
 * wrapper inside a link may hold an equation, a revision, a field or a
 * content control that the link itself may not.
 */
export const holdsUnlinkableContent = (item: ParagraphContent): boolean =>
  item.type === "inlineWrapper" &&
  item.content.some(
    (child) => !isHyperlinkContent(child) || holdsUnlinkableContent(child),
  );

/** A field's content once its linked pieces are rebuilt: each came out of it. */
const asFieldContent = (
  content: readonly ParagraphContent[],
): SimpleField["content"] => {
  const admitted = content.filter(isSimpleFieldContent);
  if (admitted.length !== content.length)
    panic("A field's linked pieces must stay field content");
  return admitted;
};

/**
 * `content`, read under `link`, with the link put back around what it can hold
 * and the rest left in the container that holds it: a wrapper, a revision, a
 * content control or a field owns the link's pieces inside it, as the
 * paragraph parser reads a wrapper under a link's revision.
 */
const linkedPieces = (
  link: Hyperlink,
  content: readonly ParagraphContent[],
): ParagraphContent[] => {
  const pieces: ParagraphContent[] = [];
  let linked: Hyperlink["children"][number][] = [];
  const flush = (): void => {
    if (linked.length === 0) return;
    pieces.push({ ...link, children: linked });
    linked = [];
  };
  for (const item of content) {
    if (isHyperlinkContent(item) && !holdsUnlinkableContent(item)) {
      linked.push(item);
      continue;
    }
    flush();
    switch (item.type) {
      case "inlineWrapper":
        pieces.push({ ...item, content: linkedPieces(link, item.content) });
        break;
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        pieces.push({
          ...item,
          content: linkedPieces(link, item.content).filter(
            isTrackedChangeWrapperChild,
          ),
        });
        break;
      case "inlineSdt":
        pieces.push({
          ...item,
          content: linkedPieces(link, item.content).filter(isInlineSdtContent),
        });
        break;
      case "simpleField":
        pieces.push({
          ...item,
          content: asFieldContent(linkedPieces(link, item.content)),
        });
        break;
      // An equation, a nested link and anything else with no runs of its own
      // under this link stands beside the link's pieces.
      default:
        pieces.push(item);
    }
  }
  flush();
  return pieces;
};

/**
 * Inline content with every link whose wrapper holds content the link cannot
 * hold split into pieces that content's containers own. The editor carries a
 * link as a mark on the leaves of its runs; an equation or a revision inside
 * one has no place under that mark, and was dropped. Returns `content` itself
 * when nothing in it needed splitting.
 */
export const withLinkedPiecesInWrappers = (
  content: readonly ParagraphContent[],
): readonly ParagraphContent[] => {
  let changed = false;
  const next: ParagraphContent[] = [];
  for (const item of content) {
    const normalized = linkedPiecesOf(item);
    if (normalized.length !== 1 || normalized[0] !== item) changed = true;
    next.push(...normalized);
  }
  return changed ? next : content;
};

const linkedPiecesOf = (
  item: ParagraphContent,
): readonly ParagraphContent[] => {
  switch (item.type) {
    case "hyperlink":
      if (!item.children.some(holdsUnlinkableContent)) return [item];
      // A container among the pieces may hold a further link of its own.
      return withLinkedPiecesInWrappers(linkedPieces(item, item.children));
    case "inlineWrapper": {
      const content = withLinkedPiecesInWrappers(item.content);
      return content === item.content
        ? [item]
        : [{ ...item, content: [...content] }];
    }
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo": {
      const content = withLinkedPiecesInWrappers(item.content);
      return content === item.content
        ? [item]
        : [{ ...item, content: content.filter(isTrackedChangeWrapperChild) }];
    }
    case "inlineSdt": {
      const content = withLinkedPiecesInWrappers(item.content);
      return content === item.content
        ? [item]
        : [{ ...item, content: content.filter(isInlineSdtContent) }];
    }
    case "simpleField": {
      const content = withLinkedPiecesInWrappers(item.content);
      return content === item.content
        ? [item]
        : [{ ...item, content: asFieldContent(content) }];
    }
    default:
      return [item];
  }
};
