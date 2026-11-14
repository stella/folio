import { panic, Result } from "better-result";

import { sanitizeExternalUrl } from "../markdown/href";
import type {
  Document,
  Hyperlink,
  Paragraph,
  ParagraphContent,
  Run,
  TrackedRunChange,
} from "../model/document";
import { applyDocumentOps } from "./apply";
import { type CompileIntent, nextCompoundMode } from "./compoundIntent";
import type { EditorIntentMode } from "./editorIntent";
import { structurallyEqual } from "./equality";
import { resolveGap } from "./gaps";
import { freshenIdentities } from "./identity";
import {
  countKeys,
  identityKeysIn,
  idKey,
  packageIdentityKeys,
  reservedIdentityKeysIn,
} from "./ids";
import {
  asParagraphContent,
  type Gap,
  type InlineNode,
  leafSpans,
  partitionContent,
} from "./leaves";
import { isTrackedChild, paragraphLength, withChildren } from "./offsets";
import { findStoryParagraph } from "./paragraphLocator";
import { selectedParagraphRuns } from "./plan";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { findStoryBody } from "./stories";
import { DOCUMENT_OP_TYPES, type DocumentOp, type TextPosition } from "./types";

export type HyperlinkEditorIntent =
  | {
      type: "setHyperlink";
      from: TextPosition;
      to: TextPosition;
      href: string;
      tooltip?: string;
    }
  | {
      type: "removeHyperlink";
      from: TextPosition;
      to: TextPosition;
      hyperlinkStyleId?: string;
    }
  | {
      type: "insertHyperlink";
      from: TextPosition;
      to: TextPosition;
      text: string;
      href: string;
      tooltip?: string;
    };

const refuse = (
  message: string,
  reason: DocumentOpRefusal["reason"] = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
) =>
  Result.err(
    new DocumentOpRefusal({
      message,
      reason,
      opType: DOCUMENT_OP_TYPES.REPLACE_INLINE,
    }),
  );

/**
 * Records a link edit can cut. A run's pending formatting change and a
 * tracked change are cut with it, and the halves after the first take fresh
 * ids; wrappers it cannot model (controls, opaque captures) are not cut.
 */
/** Content rebuilt inside a tracked change, which admits every kind it held. */
const trackedContent = (
  items: readonly ParagraphContent[],
): TrackedRunChange["content"] =>
  items.map((item) =>
    isTrackedChild(item)
      ? item
      : panic(
          "A link edit put a record into a tracked change that it cannot hold.",
        ),
  );

const HYPERLINK_STYLE_ID = "Hyperlink";

const styledRun = (
  run: Run,
  intent: Exclude<HyperlinkEditorIntent, { type: "insertHyperlink" }>,
): Run => {
  if (run.formatting === undefined) return run;
  const formatting = { ...run.formatting };
  if (intent.type === "setHyperlink") delete formatting.color;
  else if (
    formatting.styleId === HYPERLINK_STYLE_ID ||
    (intent.hyperlinkStyleId !== undefined &&
      formatting.styleId === intent.hyperlinkStyleId)
  )
    delete formatting.styleId;
  if (structurallyEqual(formatting, run.formatting)) return run;
  const next = { ...run };
  if (Object.keys(formatting).length === 0) delete next.formatting;
  else next.formatting = formatting;
  return next;
};

/** Flatten only the edited hyperlinks; unrelated run formatting and zero-width markers stay. */
const linkedChildren = (
  items: readonly ParagraphContent[],
  intent: Exclude<HyperlinkEditorIntent, { type: "insertHyperlink" }>,
): ParagraphContent[] => {
  const children: ParagraphContent[] = [];
  for (const item of items) {
    switch (item.type) {
      case "run":
        children.push(
          intent.type === "setHyperlink" ? styledRun(item, intent) : item,
        );
        break;
      case "simpleField":
      case "complexField":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "preservedInline":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "mathEquation":
        children.push(item);
        break;
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        children.push({
          ...item,
          content: trackedContent(linkedChildren(item.content, intent)),
        });
        break;
      // A content control or wrapper stays whole; the link edit goes inside.
      case "inlineSdt":
        children.push(withChildren(item, linkedChildren(item.content, intent)));
        break;
      case "inlineWrapper":
        children.push({
          ...item,
          content: linkedChildren(item.content, intent),
        });
        break;
      case "hyperlink":
        if (item.children.length === 0) {
          children.push(item);
          break;
        }
        for (const child of item.children) {
          if (child.type === "run") children.push(styledRun(child, intent));
          else children.push(child);
        }
        break;
      default: {
        const unhandled: never = item;
        return unhandled;
      }
    }
  }
  return children;
};

const hyperlink = (
  intent: Exclude<HyperlinkEditorIntent, { type: "removeHyperlink" }>,
  children: Hyperlink["children"],
): Hyperlink => {
  const result: Hyperlink = { type: "hyperlink", children };
  if (intent.href.startsWith("#")) result.anchor = intent.href.slice(1);
  else if (intent.href !== "") result.href = intent.href;
  if (intent.tooltip !== undefined) result.tooltip = intent.tooltip;
  return result;
};

/** Link text segments while structural markers and empty links retain their authored positions. */
const wrapLinkedChildren = (
  intent: Extract<HyperlinkEditorIntent, { type: "setHyperlink" }>,
  children: readonly ParagraphContent[],
): ParagraphContent[] => {
  const result: ParagraphContent[] = [];
  let link: Hyperlink | undefined;
  for (const child of children) {
    if (child.type === "run") {
      if (link === undefined) {
        link = hyperlink(intent, []);
        result.push(link);
      }
      link.children.push(child);
      continue;
    }
    link = undefined;
    // A change's own text is linked inside it.
    if (
      child.type === "insertion" ||
      child.type === "deletion" ||
      child.type === "moveFrom" ||
      child.type === "moveTo"
    ) {
      result.push({
        ...child,
        content: trackedContent(wrapLinkedChildren(intent, child.content)),
      });
      continue;
    }
    // A content control or wrapper keeps its place; its own text is linked
    // inside it.
    if (child.type === "inlineSdt") {
      result.push(
        withChildren(child, wrapLinkedChildren(intent, child.content)),
      );
      continue;
    }
    if (child.type === "inlineWrapper") {
      result.push({
        ...child,
        content: wrapLinkedChildren(intent, child.content),
      });
      continue;
    }
    // An equation, which a link cannot hold, stays between two pieces of it.
    result.push(child);
  }
  return result;
};

/**
 * The edited paragraphs with fresh ids where a cut left two records sharing
 * one: a run's pending formatting change or a tracked change cut at a link's
 * edge keeps its id on the first half, and the halves after it take the
 * mode's new ids, as any cut does.
 */
const freshenedParagraphs = (
  document: Document,
  edited: readonly { paragraph: Paragraph; content: ParagraphContent[] }[],
  mode: EditorIntentMode,
): Result<Paragraph[], DocumentOpRefusal> => {
  const before = edited.map(({ paragraph }) => paragraph);
  const freshened = freshenIdentities({
    before,
    after: edited.map(({ paragraph, content }) => ({ ...paragraph, content })),
    newIds: mode.newIds ?? {},
    usedElsewhere: () => {
      const counts = countKeys(packageIdentityKeys(document.package));
      for (const key of identityKeysIn(before))
        counts.set(key, (counts.get(key) ?? 1) - 1);
      return new Set(
        [...counts].flatMap(([key, count]) => (count > 0 ? [key] : [])),
      );
    },
    reserved: new Set(reservedIdentityKeysIn(document.package)),
  });
  switch (freshened.kind) {
    case "fresh":
      return Result.ok(freshened.paragraphs);
    case "needsIds":
      return refuse(
        `The hyperlink edit needs ${String(freshened.missing)} additional revision ids.`,
        DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
      );
    case "invalidId":
      return refuse(
        `Revision id ${String(freshened.id)} is not available.`,
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      );
    default: {
      const unhandled: never = freshened;
      return unhandled;
    }
  }
};

type CompileHyperlinkIntentOptions = {
  intent: HyperlinkEditorIntent;
  mode: EditorIntentMode;
  compileIntent: CompileIntent;
};

/** Run content that names a record of its own, which a copy would duplicate. */
const IDENTIFIED_RUN_CONTENT: ReadonlySet<string> = new Set([
  "footnoteRef",
  "endnoteRef",
  "drawing",
  "shape",
  "preservedXml",
]);

/**
 * The copy states a run's formatting as it now is; a pending formatting
 * change stays on the struck original.
 */
const withoutPropertyChanges = (run: Run): Run => {
  if (run.propertyChanges === undefined) return run;
  const { propertyChanges: _changes, ...rest } = run;
  return rest;
};

/** A change's own content in place of the change: the copy is the suggester's. */
const withoutTrackedWrappers = (
  items: readonly ParagraphContent[],
): ParagraphContent[] =>
  items.flatMap((item): ParagraphContent[] => {
    switch (item.type) {
      case "insertion":
      case "moveTo":
        return withoutTrackedWrappers(item.content);
      case "hyperlink":
        return [
          {
            ...item,
            children: item.children.map((child) =>
              child.type === "run" ? withoutPropertyChanges(child) : child,
            ),
          },
        ];
      case "run":
        return [withoutPropertyChanges(item)];
      default:
        return [item];
    }
  });

/**
 * Whether striking `segments` would leave a piece of a link holding only
 * zero-width markers. A tracked change cannot sit inside a link, so striking
 * part of one cuts it there and then; accepting would leave that piece a link
 * around no text, where the edit made directly leaves the bare markers.
 */
const leavesMarkerOnlyLinkPiece = (
  paragraph: Paragraph,
  segments: readonly LinkSegment[],
): boolean => {
  const pieces = new Map<InlineNode, string>();
  for (const { ancestors, before, after } of leafSpans(paragraph.content)) {
    // A link holds no link, so a leaf has at most one around it.
    const link = ancestors.find(({ type }) => type === "hyperlink");
    if (link === undefined) continue;
    let token = "z";
    if (after.offset > before.offset) {
      // A text leaf the selection cuts keeps text on the side it is cut.
      const struck = segments.find(
        ({ start, end }) =>
          start.offset < after.offset && end.offset > before.offset,
      );
      token =
        struck === undefined
          ? "u"
          : `${before.offset < struck.start.offset ? "u" : ""}s${
              after.offset > struck.end.offset ? "u" : ""
            }`;
    }
    pieces.set(link, (pieces.get(link) ?? "") + token);
  }
  return [...pieces.values()].some(
    (tokens) =>
      tokens.includes("s") &&
      tokens
        .split("s")
        .some((piece) => piece.includes("z") && !piece.includes("u")),
  );
};

type LinkSegment = { start: Gap; end: Gap; context: readonly InlineNode[] };

/**
 * The stretches of a paragraph's selected content a link edit changes, each
 * inside one tracked change (or none) and holding no zero-width marker: the
 * text a suggestion strikes and copies. Content already on its way out stays
 * as it is: accepting removes it, and rejecting brings back what it was.
 */
const linkSegments = (
  intent: Exclude<HyperlinkEditorIntent, { type: "insertHyperlink" }>,
  paragraph: Paragraph,
  start: Gap,
  end: Gap,
): LinkSegment[] => {
  const segments: LinkSegment[] = [];
  let open: LinkSegment | undefined;
  const close = () => {
    if (open !== undefined) segments.push(open);
    open = undefined;
  };
  for (const span of leafSpans(paragraph.content)) {
    const { ancestors, before, after } = span;
    const linked = ancestors.some(({ type }) => type === "hyperlink");
    const changes =
      linked ||
      (intent.type === "setHyperlink" &&
        ancestors.some(({ type }) => type === "run") &&
        ancestors.every(
          ({ type }) =>
            type === "run" || type === "insertion" || type === "moveTo",
        ));
    if (
      after.offset <= start.offset ||
      before.offset >= end.offset ||
      after.offset === before.offset ||
      !changes ||
      ancestors.some(({ type }) => type === "deletion" || type === "moveFrom")
    ) {
      close();
      continue;
    }
    const context = ancestors.filter(
      ({ type }) => type === "insertion" || type === "moveTo",
    );
    // Text the selection's ends fall inside is cut there.
    const from = before.offset < start.offset ? start : before;
    const to = after.offset > end.offset ? end : after;
    if (
      open !== undefined &&
      open.end.offset === from.offset &&
      open.context.length === context.length &&
      open.context.every((wrapper, index) => wrapper === context[index])
    ) {
      open.end = to;
      continue;
    }
    close();
    open = { start: from, end: to, context };
  }
  close();
  return segments;
};

/**
 * A suggested link change. The file records no change of a link, so each
 * stretch of changed text is struck and its changed copy inserted after it,
 * as a paste over it would be: rejecting brings the text back as it was, and
 * accepting leaves what the edit made directly leaves.
 */
const compileTrackedLinkEdit = (
  document: Document,
  {
    intent,
    mode,
    compileIntent,
    pieces,
  }: {
    intent: Exclude<HyperlinkEditorIntent, { type: "insertHyperlink" }>;
    mode: EditorIntentMode;
    compileIntent: CompileIntent;
    pieces: readonly { paragraph: Paragraph; start: Gap; end: Gap }[];
  },
): Result<
  { ops: DocumentOp[]; selection: TextPosition },
  DocumentOpRefusal
> => {
  const steps: {
    blockId: string;
    start: Gap;
    end: Gap;
    copy: ParagraphContent[];
  }[] = [];
  for (const { paragraph, start, end } of pieces) {
    const blockId = paragraph.paraId ?? "";
    const segments = linkSegments(intent, paragraph, start, end);
    if (leavesMarkerOnlyLinkPiece(paragraph, segments))
      return refuse(
        "A link suggestion cannot strike text beside a marker inside the same link.",
        DOCUMENT_OP_REFUSAL_REASONS.UNRECORDABLE_LINK_SUGGESTION,
      );
    for (const segment of segments) {
      const [, slice = []] = partitionContent(paragraph.content, [
        segment.start,
        segment.end,
      ]);
      // A copy would duplicate the content control or wrapper around it.
      if (
        leafSpans(asParagraphContent(slice)).some(({ ancestors }) =>
          ancestors.some(
            ({ type }) => type === "inlineSdt" || type === "inlineWrapper",
          ),
        )
      )
        return refuse(
          "A link suggestion cannot copy text inside a content control or a direction or smart tag wrapper.",
          DOCUMENT_OP_REFUSAL_REASONS.UNRECORDABLE_LINK_SUGGESTION,
        );
      const own = withoutTrackedWrappers(asParagraphContent(slice));
      const children = linkedChildren(own, intent);
      const copy =
        intent.type === "setHyperlink"
          ? wrapLinkedChildren(intent, children)
          : children;
      if (
        identityKeysIn(copy).length > 0 ||
        copy.some(({ type }) => type === "preservedInline") ||
        leafSpans(copy).some(({ node }) =>
          IDENTIFIED_RUN_CONTENT.has(node.type),
        )
      )
        return refuse(
          "A link suggestion cannot copy content that names a record of its own, such as a drawing or a note reference.",
          DOCUMENT_OP_REFUSAL_REASONS.UNRECORDABLE_LINK_SUGGESTION,
        );
      steps.push({ blockId, start: segment.start, end: segment.end, copy });
    }
  }
  const ops: DocumentOp[] = [];
  let current = document;
  let stepMode = mode;
  let previous: Document | undefined;
  // Last first, so the positions of the stretches before stay as they were.
  for (const step of steps.toReversed()) {
    if (previous !== undefined) {
      const next = nextCompoundMode(stepMode, previous, current);
      if (next.isErr()) return Result.err(next.error);
      stepMode = next.value;
    }
    const position = (gap: Gap): TextPosition => ({
      story: intent.from.story,
      blockId: step.blockId,
      offset: gap.offset,
      zeroWidthBefore: gap.zeroWidthBefore,
    });
    const compiled = compileIntent(
      current,
      {
        type: "replaceFragment",
        from: position(step.start),
        to: position(step.end),
        paragraphs: [{ type: "paragraph", content: step.copy }],
        openStart: 1,
        openEnd: 1,
      },
      stepMode,
    );
    if (compiled.isErr()) return Result.err(compiled.error);
    const applied = applyDocumentOps(current, compiled.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    previous = current;
    current = applied.value.document;
    ops.push(...compiled.value.ops);
  }
  // The copies in the last paragraph all stand before the selection's end,
  // so it moves by what that paragraph grew. A stretch this author already
  // suggested inserting changes in place, adding no copy.
  const lengthIn = (version: Document) => {
    const body = findStoryBody(version, intent.to.story);
    const located =
      body === undefined
        ? undefined
        : findStoryParagraph(body, intent.to.blockId);
    return located === undefined
      ? panic("A link suggestion lost the paragraph its selection ends in.")
      : paragraphLength(located.paragraph);
  };
  return Result.ok({
    ops,
    selection: {
      ...intent.to,
      offset: intent.to.offset + lengthIn(current) - lengthIn(document),
    },
  });
};

/** A content control or a direction or smart tag wrapper. */
type WholeContainer = Extract<
  ParagraphContent,
  { type: "inlineSdt" | "inlineWrapper" }
>;

const isWholeContainer = (item: ParagraphContent): item is WholeContainer =>
  item.type === "inlineSdt" || item.type === "inlineWrapper";

/**
 * The index of the content control or wrapper the gap passes through, which
 * cutting the content there splits in two; `undefined` when it passes
 * through none.
 */
const cutContainer = (
  items: readonly ParagraphContent[],
  gap: Gap,
): number | undefined => {
  const [left = [], right = []] = partitionContent(items, [gap]);
  if (left.length + right.length <= items.length) return undefined;
  const index = left.length - 1;
  const item = items[index];
  return item !== undefined && isWholeContainer(item) ? index : undefined;
};

/** Two pieces of one container cut apart, or two that were never one. */
const samePieces = (
  left: ParagraphContent,
  right: ParagraphContent,
): boolean => {
  if (left.type === "inlineSdt" && right.type === "inlineSdt")
    return (
      left.properties.id !== undefined &&
      left.properties.id === right.properties.id
    );
  if (left.type === "inlineWrapper" && right.type === "inlineWrapper")
    return structurallyEqual(
      { ...left, content: [] },
      { ...right, content: [] },
    );
  return false;
};

/** The pieces of one container cut apart, whole again. */
const joinPieces = (
  left: WholeContainer,
  right: ParagraphContent,
): ParagraphContent => {
  if (left.type === "inlineSdt" && right.type === "inlineSdt")
    return withChildren(left, joinContent(left.content, right.content));
  if (left.type === "inlineWrapper" && right.type === "inlineWrapper")
    return { ...left, content: joinContent(left.content, right.content) };
  return panic("A cut container's pieces differ in kind.");
};

/** Content of two pieces, where a container nested in both is joined too. */
const joinContent = (
  left: readonly ParagraphContent[],
  right: readonly ParagraphContent[],
): ParagraphContent[] => {
  const last = left.at(-1);
  const first = right.at(0);
  if (
    last === undefined ||
    first === undefined ||
    !isWholeContainer(last) ||
    !samePieces(last, first)
  )
    return [...left, ...right];
  return [...left.slice(0, -1), joinPieces(last, first), ...right.slice(1)];
};

/**
 * The edited paragraph content, where a content control or wrapper the
 * range's ends cut stays one record, as it was: only its own content
 * changes.
 */
const joinCutContainers = ({
  before,
  replacement,
  after,
  startCut,
  endCut,
}: {
  before: readonly ParagraphContent[];
  replacement: readonly ParagraphContent[];
  after: readonly ParagraphContent[];
  startCut: number | undefined;
  endCut: number | undefined;
}): ParagraphContent[] => {
  let middle = [...replacement];
  let head = [...before];
  let tail = [...after];
  if (startCut !== undefined) {
    const left = head.at(-1);
    const right = middle.at(0);
    if (left === undefined || right === undefined || !isWholeContainer(left))
      return panic("A cut container lost a piece.");
    head = head.slice(0, -1);
    middle = [joinPieces(left, right), ...middle.slice(1)];
  }
  if (endCut !== undefined) {
    const left = middle.at(-1);
    const right = tail.at(0);
    if (left === undefined || right === undefined || !isWholeContainer(left))
      return panic("A cut container lost a piece.");
    tail = tail.slice(1);
    middle = [...middle.slice(0, -1), joinPieces(left, right)];
  }
  return [...head, ...middle, ...tail];
};

/** Wrapper edits compile to exact-inverse content ops, never to a projection snapshot. */
export const compileHyperlinkIntent = (
  document: Document,
  { intent: sourceIntent, mode, compileIntent }: CompileHyperlinkIntentOptions,
): Result<
  { ops: DocumentOp[]; selection: TextPosition },
  DocumentOpRefusal
> => {
  let intent = sourceIntent;
  if (
    intent.type !== "removeHyperlink" &&
    intent.href !== "" &&
    !intent.href.startsWith("#")
  ) {
    const href = sanitizeExternalUrl(intent.href);
    if (href === undefined)
      return refuse(
        "Hyperlink targets must use a supported URL scheme.",
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      );
    intent = { ...intent, href };
  }
  const selected = selectedParagraphRuns(document, intent.from, intent.to);
  if (selected.isErr()) return Result.err(selected.error);
  const locations = selected.value.flat();
  const pieces = [];
  for (const { paragraph } of locations) {
    const blockId = paragraph.paraId ?? "";
    const from =
      idKey(blockId) === idKey(intent.from.blockId)
        ? intent.from
        : { story: intent.from.story, blockId, offset: 0 };
    const to =
      idKey(blockId) === idKey(intent.to.blockId)
        ? intent.to
        : {
            story: intent.to.story,
            blockId,
            offset: paragraphLength(paragraph),
          };
    const start = resolveGap({
      paragraph,
      position: from,
      fallback: "insertion",
    });
    if (typeof start !== "object")
      return refuse("Invalid hyperlink range start.", start);
    const end = resolveGap({ paragraph, position: to, fallback: "insertion" });
    if (typeof end !== "object")
      return refuse("Invalid hyperlink range end.", end);
    if (
      start.offset > end.offset ||
      (start.offset === end.offset &&
        start.zeroWidthBefore > end.zeroWidthBefore)
    )
      return refuse(
        "Hyperlink range endpoints are reversed.",
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
      );
    const [before = [], middle = [], after = []] = partitionContent(
      paragraph.content,
      [start, end],
    );
    pieces.push({
      paragraph,
      from,
      to,
      start,
      end,
      startCut: cutContainer(paragraph.content, start),
      endCut: cutContainer(paragraph.content, end),
      before: asParagraphContent(before),
      middle: asParagraphContent(middle),
      after: asParagraphContent(after),
    });
  }
  if (intent.type === "insertHyperlink") {
    if (intent.text === "")
      return refuse(
        "A hyperlink insertion needs text.",
        DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
      );
    const content = [
      hyperlink(intent, [
        { type: "run", content: [{ type: "text", text: intent.text }] },
      ]),
    ];
    // A suggested link is a tracked paste of it over the selection.
    if (mode.type === "suggesting")
      return compileIntent(
        document,
        {
          type: "replaceFragment",
          from: intent.from,
          to: intent.to,
          paragraphs: [{ type: "paragraph", content }],
          openStart: 1,
          openEnd: 1,
        },
        mode,
      );
    const deletion = compileIntent(
      document,
      { type: "replaceText", from: intent.from, to: intent.to, text: "" },
      mode,
    );
    if (deletion.isErr()) return deletion;
    const applied = applyDocumentOps(document, deletion.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    const at = deletion.value.selection;
    const target = selectedParagraphRuns(applied.value.document, at, at);
    if (target.isErr()) return Result.err(target.error);
    const paragraph = target.value.at(0)?.at(0)?.paragraph;
    if (!paragraph) return refuse("Hyperlink insertion lost its paragraph.");
    const gap = resolveGap({ paragraph, position: at, fallback: "insertion" });
    if (typeof gap !== "object")
      return refuse("Invalid hyperlink insertion gap.", gap);
    const [before = [], after = []] = partitionContent(paragraph.content, [
      gap,
    ]);
    const replacement = asParagraphContent(before).concat(
      content,
      asParagraphContent(after),
    );
    const freshened = freshenedParagraphs(
      applied.value.document,
      [{ paragraph, content: replacement }],
      mode,
    );
    if (freshened.isErr()) return Result.err(freshened.error);
    return Result.ok({
      ops: deletion.value.ops.concat([
        {
          type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
          story: at.story,
          blockId: at.blockId,
          expected: paragraph.content,
          content: freshened.value[0]?.content ?? replacement,
        },
      ]),
      selection: {
        ...at,
        offset: at.offset + intent.text.length,
        zeroWidthBefore: 0,
      },
    });
  }
  if (mode.type === "suggesting")
    return compileTrackedLinkEdit(document, {
      intent,
      mode,
      compileIntent,
      pieces: pieces.filter(({ from, to }) => from.offset !== to.offset),
    });
  const edited: { paragraph: Paragraph; content: ParagraphContent[] }[] = [];
  for (const piece of pieces) {
    if (piece.from.offset === piece.to.offset) continue;
    const children = linkedChildren(piece.middle, intent);
    const replacement: ParagraphContent[] =
      intent.type === "setHyperlink"
        ? wrapLinkedChildren(intent, children)
        : children;
    if (structurallyEqual(piece.middle, replacement)) continue;
    edited.push({
      paragraph: piece.paragraph,
      content: joinCutContainers({
        before: piece.before,
        replacement,
        after: piece.after,
        startCut: piece.startCut,
        endCut: piece.endCut,
      }),
    });
  }
  const freshened = freshenedParagraphs(document, edited, mode);
  if (freshened.isErr()) return Result.err(freshened.error);
  const ops: DocumentOp[] = freshened.value.map((paragraph, index) => ({
    type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
    story: intent.from.story,
    blockId: paragraph.paraId ?? "",
    expected:
      edited[index]?.paragraph.content ??
      panic("A freshened paragraph has no edited source."),
    content: paragraph.content,
  }));
  return Result.ok({ ops, selection: intent.to });
};
