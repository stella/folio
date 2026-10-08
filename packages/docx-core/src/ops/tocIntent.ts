import { Result, panic } from "better-result";

import type { Document, Paragraph, ParagraphContent } from "../model/document";
import type { HeadingOutlineLevel } from "../model/outlineLevel";
import { applyDocumentOp } from "./apply";
import { storyParagraphs, storyBody } from "./blocks";
import { resolveGap } from "./gaps";
import { idKey, packageBookmarkCensus, packageParagraphIds } from "./ids";
import { paragraphLength, paragraphLogicalText } from "./offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  SPLIT_HALVES,
  type DocumentOp,
  type TextPosition,
} from "./types";
import type { EditorIntentMode } from "./editorIntent";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import { sameStory } from "./stories";

export type TocHeading = {
  blockId: string;
  text: string;
  /** Zero-based OOXML outline level, from 0 through 8. */
  level: HeadingOutlineLevel;
  styleId?: string;
};

export type GenerateTOCIntent = {
  type: "generateTOC";
  at: TextPosition;
  title: string;
  headings: readonly TocHeading[];
  titleStyleId?: string;
  tabPosition: number;
};

const refuse = (
  message: string,
  reason: DocumentOpRefusal["reason"] = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
) =>
  Result.err(new DocumentOpRefusal({ message, reason, opType: DOCUMENT_OP_TYPES.REPLACE_INLINE }));

const isPendingOrUnsupported = (content: readonly ParagraphContent[]): boolean =>
  content.some((item) => {
    if (item.type === "bookmarkStart" || item.type === "bookmarkEnd") return false;
    if (item.type === "run") return (item.propertyChanges?.length ?? 0) > 0;
    if (item.type === "hyperlink") return isPendingOrUnsupported(item.children);
    return true;
  });

const takeBookmarkId = (
  census: { ids: ReadonlySet<number> },
  reserved: Set<number>,
): number | undefined => {
  for (let id = 1; Number.isSafeInteger(id); id += 1) {
    if (census.ids.has(id) || reserved.has(id)) continue;
    reserved.add(id);
    return id;
  }
  return undefined;
};

type TakeParagraphIdsOptions = {
  document: Document;
  count: number;
  reserved: ReadonlySet<string> | undefined;
};
const takeParagraphIds = ({
  document,
  count,
  reserved,
}: TakeParagraphIdsOptions): string[] | undefined => {
  const occupied = new Set(
    [...packageParagraphIds(document.package), ...(reserved ?? [])].map(idKey),
  );
  const ids: string[] = [];
  for (let id = 1; id < 0x8000_0000 && ids.length < count; id += 1) {
    const candidate = id.toString(16).padStart(8, "0").toUpperCase();
    if (occupied.has(candidate)) continue;
    occupied.add(candidate);
    ids.push(candidate);
  }
  return ids.length === count ? ids : undefined;
};

const takeBookmarkName = (
  census: { names: ReadonlySet<string> },
  reserved: Set<string>,
): string => {
  for (let suffix = 1; ; suffix += 1) {
    const name = `_Toc${suffix}`;
    if (census.names.has(name) || reserved.has(name)) continue;
    reserved.add(name);
    return name;
  }
};

const paragraphFor = (document: Document, position: TextPosition): Paragraph | undefined =>
  storyParagraphs(storyBody(document, position.story)).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(position.blockId),
  )?.paragraph;

const bookmarkAtStart = (paragraph: Paragraph, id: number, name: string): ParagraphContent[] => [
  { type: "bookmarkStart", id, name },
  ...paragraph.content,
  { type: "bookmarkEnd", id },
];

/** Map original endpoints through the bookmark wrappers emitted by TOC compilation. */
export const mapTocBookmarkPosition = (
  position: TextPosition,
  ops: readonly DocumentOp[],
): TextPosition => {
  if (position.offset !== 0) return position;
  let prepended = 0;
  for (const op of ops) {
    if (
      op.type !== DOCUMENT_OP_TYPES.REPLACE_INLINE ||
      !sameStory(op.story, position.story) ||
      idKey(op.blockId) !== idKey(position.blockId)
    )
      continue;
    // bookmarkAtStart adds one opening marker before every original gap at offset zero.
    prepended += 1;
  }
  return prepended === 0
    ? position
    : { ...position, zeroWidthBefore: (position.zeroWidthBefore ?? 0) + prepended };
};

const paragraphBookmarks = (
  content: readonly ParagraphContent[],
): { starts: Map<number, string>; ends: Set<number> } => {
  const starts = new Map<number, string>();
  const ends = new Set<number>();
  for (const item of content) {
    if (item.type === "bookmarkStart") starts.set(item.id, item.name);
    else if (item.type === "bookmarkEnd") ends.add(item.id);
    else if (item.type === "hyperlink") {
      const nested = paragraphBookmarks(item.children);
      for (const [id, name] of nested.starts) starts.set(id, name);
      for (const id of nested.ends) ends.add(id);
    }
  }
  return { starts, ends };
};

const tocParagraphs = (
  intent: GenerateTOCIntent,
  links: readonly { heading: TocHeading; name: string }[],
): Paragraph[] => {
  const paragraphs: Paragraph[] = [
    {
      type: "paragraph",
      formatting: {
        ...(intent.titleStyleId === undefined ? {} : { styleId: intent.titleStyleId }),
        alignment: "center",
      },
      content: [
        {
          type: "run",
          formatting: { bold: true },
          content: [{ type: "text", text: intent.title }],
        },
      ],
    },
  ];
  for (const { heading, name } of links) {
    paragraphs.push({
      type: "paragraph",
      formatting: {
        ...(heading.styleId === undefined ? {} : { styleId: heading.styleId }),
        ...(heading.level === 0 ? {} : { indentLeft: heading.level * 720 }),
        tabs: [{ position: intent.tabPosition, alignment: "right", leader: "dot" }],
      },
      content: [
        {
          type: "hyperlink",
          anchor: name,
          children: [{ type: "run", content: [{ type: "text", text: heading.text }] }],
        },
        { type: "run", content: [{ type: "tab" }] },
        {
          type: "complexField",
          instruction: `PAGEREF ${name} \\h`,
          fieldType: "PAGEREF",
          fieldCode: [
            { type: "run", content: [{ type: "instrText", text: `PAGEREF ${name} \\h` }] },
          ],
          fieldResult: [{ type: "run", content: [{ type: "text", text: "1" }] }],
          dirty: true,
        },
      ],
    });
  }
  return paragraphs;
};

/** These paragraphs are newly constructed for this compile, so assigning their ids is local. */
const identifyTocParagraphs = (paragraphs: Paragraph[], ids: readonly string[]): Paragraph[] => {
  for (const [index, paragraph] of paragraphs.entries())
    paragraph.paraId =
      ids.at(index) ?? panic("TOC allocation must cover every generated paragraph.");
  return paragraphs;
};

export const compileGenerateTOCIntent = (
  document: Document,
  intent: GenerateTOCIntent,
  mode: EditorIntentMode,
): Result<{ ops: DocumentOp[]; selection: TextPosition }, DocumentOpRefusal> => {
  if (intent.headings.length === 0) return Result.ok({ ops: [], selection: intent.at });
  if (mode.type === "suggesting")
    return refuse("TOC suggestions require serializable wrapper review provenance.");
  if (intent.at.story !== OP_STORIES.MAIN)
    return refuse("TOC generation is supported in the main document story only.");
  if (
    intent.title === "" ||
    hasIllegalXmlCharacters(intent.title) ||
    intent.headings.some(({ text }) => hasIllegalXmlCharacters(text))
  )
    return refuse(
      "TOC text is empty or contains characters that XML cannot represent.",
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
    );
  if (!Number.isSafeInteger(intent.tabPosition) || intent.tabPosition < 0)
    return refuse(
      "The TOC tab position must be a nonnegative integer.",
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
    );
  const target = paragraphFor(document, intent.at);
  if (target === undefined)
    return refuse(
      "The TOC insertion paragraph does not exist.",
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
    );
  if (target.paraId === undefined)
    return refuse(
      "The TOC insertion paragraph has no stable identity.",
      DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID,
    );
  const initialGap = resolveGap({ paragraph: target, position: intent.at, fallback: "insertion" });
  if (typeof initialGap !== "object")
    return refuse("The TOC insertion position is invalid.", initialGap);
  const targetLength = paragraphLength(target);

  const locations = storyParagraphs(storyBody(document, intent.at.story));
  const links: { heading: TocHeading; name: string }[] = [];
  const bookmarkIds: number[] = [];
  const seenHeadings = new Set<string>();
  const census = packageBookmarkCensus(document.package);
  const reservedBookmarkIds = new Set<number>();
  const reservedBookmarkNames = new Set<string>();
  for (const heading of intent.headings) {
    if (!Number.isInteger(heading.level) || heading.level < 0 || heading.level > 8)
      return refuse(
        "TOC heading levels must be between zero and eight.",
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
      );
    const location = locations.find(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(heading.blockId),
    );
    if (location === undefined)
      return refuse(
        "A TOC heading paragraph does not exist.",
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      );
    if (
      paragraphLogicalText(location.paragraph).trim() !== heading.text.trim() ||
      location.paragraph.paraId === undefined ||
      isPendingOrUnsupported(location.paragraph.content) ||
      (location.paragraph.propertyChanges?.length ?? 0) > 0 ||
      location.paragraph.pPrMark !== undefined ||
      location.paragraph.reviewCarrier !== undefined
    )
      return refuse("A TOC heading is stale or contains pending review wrappers.");
    const headingKey = idKey(heading.blockId);
    if (seenHeadings.has(headingKey)) return refuse("A TOC heading is listed more than once.");
    seenHeadings.add(headingKey);
    const local = paragraphBookmarks(location.paragraph.content);
    const existing = [...local.starts].find(
      ([id, name]) =>
        name.startsWith("_Toc") &&
        local.ends.has(id) &&
        census.pairs.some((pair) => pair.id === id && pair.name === name),
    );
    if (existing !== undefined) {
      bookmarkIds.push(existing[0]);
      links.push({ heading, name: existing[1] });
      continue;
    }
    const id = takeBookmarkId(census, reservedBookmarkIds);
    if (id === undefined) return refuse("The bookmark identity space is exhausted.");
    bookmarkIds.push(id);
    links.push({ heading, name: takeBookmarkName(census, reservedBookmarkNames) });
  }

  const ops: DocumentOp[] = [];
  let insertionAt: TextPosition = { ...intent.at, ...initialGap };
  for (const [index, { heading, name }] of links.entries()) {
    const paragraph = locations.find(
      ({ paragraph: candidate }) => idKey(candidate.paraId ?? "") === idKey(heading.blockId),
    )?.paragraph;
    if (paragraph === undefined) return refuse("A TOC heading paragraph disappeared.");
    const id = bookmarkIds[index];
    if (id === undefined) return refuse("A TOC bookmark id was not allocated.");
    if (census.pairs.some((pair) => pair.id === id && pair.name === name)) continue;
    const op = {
      type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
      story: intent.at.story,
      blockId: paragraph.paraId ?? "",
      expected: paragraph.content,
      content: bookmarkAtStart(paragraph, id, name),
    } as const;
    ops.push(op);
    const applied = applyDocumentOp(document, op);
    if (applied.isErr()) return Result.err(applied.error);
    document = applied.value.document;
    insertionAt = mapTocBookmarkPosition(insertionAt, [op]);
  }

  const currentTarget = paragraphFor(document, intent.at);
  if (currentTarget === undefined) return refuse("The TOC insertion paragraph disappeared.");
  const gap = resolveGap({
    paragraph: currentTarget,
    position: insertionAt,
    fallback: "insertion",
  });
  if (typeof gap !== "object")
    return refuse("The TOC insertion position is invalid after bookmark updates.", gap);
  if (
    gap.offset !== 0 &&
    gap.offset !== paragraphLength(currentTarget) &&
    (isPendingOrUnsupported(currentTarget.content) ||
      (currentTarget.propertyChanges?.length ?? 0) > 0 ||
      currentTarget.pPrMark !== undefined ||
      currentTarget.reviewCarrier !== undefined)
  )
    return refuse("TOC insertion cannot split pending review wrappers.");
  const toc = tocParagraphs(intent, links);
  const story = intent.at.story;
  if (gap.offset === 0 || gap.offset === targetLength) {
    const ids = takeParagraphIds({ document, count: toc.length, reserved: mode.reservedBlockIds });
    if (ids === undefined) return refuse("The paragraph identity space is exhausted.");
    ops.push({
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story,
      at:
        gap.offset === 0
          ? { type: "before", blockId: intent.at.blockId }
          : { type: "after", blockId: intent.at.blockId },
      blocks: identifyTocParagraphs(toc, ids),
    });
    return Result.ok({
      ops,
      selection: { ...insertionAt, ...gap },
    });
  }

  const newBlockId = takeParagraphIds({ document, count: 1, reserved: mode.reservedBlockIds })?.at(
    0,
  );
  if (newBlockId === undefined) return refuse("The paragraph identity space is exhausted.");
  const split = {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: { ...insertionAt, ...gap },
    newBlockId,
    newHalf: SPLIT_HALVES.FIRST,
    newParagraph: target.formatting === undefined ? {} : { formatting: target.formatting },
  } as const;
  ops.push(split);
  const afterSplit = applyDocumentOp(document, split);
  if (afterSplit.isErr()) return Result.err(afterSplit.error);
  document = afterSplit.value.document;
  const tocIds = takeParagraphIds({ document, count: toc.length, reserved: mode.reservedBlockIds });
  if (tocIds === undefined) return refuse("The paragraph identity space is exhausted.");
  const blocks = identifyTocParagraphs(toc, tocIds);
  ops.push({
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story,
    at: { type: "before", blockId: intent.at.blockId },
    blocks,
  });
  // The retained half starts with the markers that followed the split gap; stay before them.
  return Result.ok({
    ops,
    selection: { story, blockId: intent.at.blockId, offset: 0, zeroWidthBefore: 0 },
  });
};
