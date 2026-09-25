/**
 * Synthetic documents and operations for the operation properties.
 *
 * The documents are built from the model types directly, with the fields a
 * parsed package carries that an edit must not lose: unmodelled attributes,
 * captured property markup, tracked changes and property changes, range
 * markers, fields, content controls, tables, section breaks and the parser's
 * per-section view. Every paragraph has a unique `paraId`.
 */

import fc from "fast-check";

import type {
  BlockContent,
  Document,
  DocumentBody,
  Hyperlink,
  InlineSdt,
  InlineWrapper,
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  PreservedAttribute,
  Run,
  RunContent,
  Section,
  SectionProperties,
  TextFormatting,
  TrackedRunContent,
} from "../../model/document";
import { storyParagraphs } from "../blocks";
import { paragraphIdsIn } from "../ids";
import { deleteBetween } from "../inline";
import { runGaps, zeroWidthLeavesAt } from "../leaves";
import { paragraphLength, paragraphLogicalText } from "../offsets";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpType,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  type ParagraphPropsPatch,
  type RunPropsPatch,
  type SplitParagraphFields,
  type TextPosition,
} from "../types";

const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Includes a surrogate pair, so positions inside one are generated too.
const textArbitrary = fc
  .array(fc.constantFrom("a", "b", " ", "é", "ß", "😀"), { maxLength: 5 })
  .map((parts) => parts.join(""));

const nonEmptyTextArbitrary = fc
  .array(fc.constantFrom("x", "y", "ü", "😀"), { minLength: 1, maxLength: 3 })
  .map((parts) => parts.join(""));

const preservedAttributesArbitrary: fc.Arbitrary<PreservedAttribute[]> = fc.array(
  fc.record({
    namespace: fc.constant(W_NAMESPACE),
    name: fc.constantFrom("rsidR", "rsidRPr", "rsidDel"),
    value: fc.constantFrom("00A1B2C3", "00FF0011"),
  }),
  { minLength: 1, maxLength: 2 },
);

export const textFormattingArbitrary: fc.Arbitrary<TextFormatting> = fc.record(
  {
    bold: fc.boolean(),
    italic: fc.boolean(),
    fontSize: fc.integer({ min: 16, max: 28 }),
    styleId: fc.constantFrom("Emphasis", "Strong"),
    preserved: fc.constant({ children: [{ index: 38, xml: "<w:webHidden/>" }] }),
  },
  { requiredKeys: [] },
);

const paragraphFormattingArbitrary: fc.Arbitrary<ParagraphFormatting> = fc.record(
  {
    alignment: fc.constantFrom("start", "center", "end"),
    spaceBefore: fc.integer({ min: 0, max: 240 }),
    keepNext: fc.boolean(),
    styleId: fc.constantFrom("Heading1", "BodyText"),
    runProperties: textFormattingArbitrary,
    preserved: fc.constant({ children: [{ index: 3, xml: "<w:suppressOverlap/>" }] }),
  },
  { requiredKeys: [] },
);

const ATOM_RUN_CONTENT: readonly RunContent[] = [
  { type: "tab" },
  { type: "break", breakType: "textWrapping" },
  { type: "symbol", font: "Symbol", char: "F0B7" },
  { type: "footnoteRef", id: 2 },
  { type: "preservedXml", xml: "<w:ruby/>", text: "r" },
  { type: "renderedPageBreak" },
];

const runContentArbitrary: fc.Arbitrary<RunContent> = fc.oneof(
  { weight: 6, arbitrary: textArbitrary.map((text): RunContent => ({ type: "text", text })) },
  { weight: 2, arbitrary: fc.constantFrom(...ATOM_RUN_CONTENT) },
);

const trackedInfoArbitrary = fc.record(
  {
    id: fc.integer({ min: 1, max: 999 }),
    author: fc.constantFrom("A", "B"),
    date: fc.constant("2026-01-02T03:04:05Z"),
  },
  { requiredKeys: ["id", "author"] },
);

export const runArbitrary: fc.Arbitrary<Run> = fc.record(
  {
    type: fc.constant("run" as const),
    content: fc.array(runContentArbitrary, { maxLength: 3 }),
    formatting: textFormattingArbitrary,
    preservedAttributes: preservedAttributesArbitrary,
    propertyChanges: trackedInfoArbitrary.map((info) => [
      {
        type: "runPropertyChange" as const,
        info: { ...info, rsid: "00C0FFEE" },
        previousFormatting: { italic: true },
      },
    ]),
  },
  { requiredKeys: ["type", "content"] },
);

const markerId = fc.integer({ min: 1, max: 40 });

const bookmarkStartArbitrary = markerId.map(
  (id): ParagraphContent => ({ type: "bookmarkStart", id, name: `_Ref${id}` }),
);
const bookmarkEndArbitrary = markerId.map((id): ParagraphContent => ({ type: "bookmarkEnd", id }));

const markerArbitrary: fc.Arbitrary<ParagraphContent> = fc.oneof(
  bookmarkStartArbitrary,
  bookmarkEndArbitrary,
  markerId.map((id): ParagraphContent => ({ type: "commentRangeStart", id })),
  markerId.map((id): ParagraphContent => ({ type: "commentRangeEnd", id })),
  markerId.map(
    (id): ParagraphContent => ({ type: "moveToRangeStart", id, name: `move${id}`, author: "A" }),
  ),
  markerId.map((id): ParagraphContent => ({ type: "moveToRangeEnd", id })),
);

const ATOMS: readonly ParagraphContent[] = [
  { type: "commentReference", id: 7 },
  { type: "mathEquation", display: "inline", ommlXml: "<m:oMath/>", plainText: "x" },
  {
    type: "simpleField",
    instruction: "PAGE",
    fieldType: "PAGE",
    content: [{ type: "run", content: [{ type: "text", text: "3" }] }],
  },
  {
    type: "complexField",
    instruction: "DATE",
    fieldType: "DATE",
    fieldCode: [{ type: "run", content: [{ type: "instrText", text: " DATE " }] }],
    fieldResult: [{ type: "run", content: [{ type: "text", text: "today" }] }],
  },
  { type: "preservedInline", xml: "<w:proofErr w:type='spellStart'/>", text: "" },
  { type: "preservedInline", xml: "<w:customXml/>", text: "c" },
];

const atomArbitrary = fc.constantFrom(...ATOMS);

const runItem = runArbitrary.map((run): ParagraphContent => run);

const wrapperArbitrary: fc.Arbitrary<InlineWrapper> = fc
  .array(fc.oneof({ weight: 4, arbitrary: runItem }, markerArbitrary, atomArbitrary), {
    minLength: 1,
    maxLength: 3,
  })
  .map((content) => ({ type: "inlineWrapper", kind: "bidi", control: "embedding", content }));

type HyperlinkChild = Hyperlink["children"][number];

const hyperlinkArbitrary: fc.Arbitrary<Hyperlink> = fc
  .array(
    fc.oneof(
      { weight: 4, arbitrary: runArbitrary.map((run): HyperlinkChild => run) },
      markerId.map((id): HyperlinkChild => ({ type: "bookmarkStart", id, name: "_Link" })),
      wrapperArbitrary,
    ),
    { minLength: 1, maxLength: 3 },
  )
  .map((children) => ({ type: "hyperlink", rId: "rId9", href: "https://example.org", children }));

const trackedContentArbitrary: fc.Arbitrary<TrackedRunContent[]> = fc.array(
  fc.oneof(
    { weight: 4, arbitrary: runArbitrary.map((run): TrackedRunContent => run) },
    hyperlinkArbitrary,
    fc.constant<TrackedRunContent>({ type: "bookmarkEnd", id: 3 }),
    fc.constant<TrackedRunContent>({
      type: "mathEquation",
      display: "inline",
      ommlXml: "<m:oMath/>",
    }),
  ),
  { minLength: 1, maxLength: 3 },
);

const trackedArbitrary: fc.Arbitrary<ParagraphContent> = fc
  .tuple(
    fc.constantFrom("insertion", "deletion", "moveFrom", "moveTo"),
    trackedInfoArbitrary,
    trackedContentArbitrary,
  )
  .map(([type, info, content]): ParagraphContent => ({ type, info, content }));

const inlineSdtArbitrary: fc.Arbitrary<InlineSdt> = fc
  .array(
    fc.oneof(
      { weight: 4, arbitrary: runArbitrary.map((run): InlineSdt["content"][number] => run) },
      hyperlinkArbitrary,
    ),
    { minLength: 1, maxLength: 3 },
  )
  .map((content) => ({
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 42, tag: "clause" },
    content,
  }));

// Two runs alike in everything but their text, as a producer splitting one writes them.
const alikeRunsArbitrary = fc
  .tuple(runArbitrary, fc.array(runContentArbitrary, { minLength: 1, maxLength: 2 }))
  .map(([run, content]): ParagraphContent[] => [run, Object.assign({}, run, { content })]);

const paragraphContentArbitrary: fc.Arbitrary<ParagraphContent[]> = fc
  .array(
    fc.oneof(
      { weight: 8, arbitrary: runItem.map((run) => [run]) },
      { weight: 2, arbitrary: alikeRunsArbitrary },
      { weight: 2, arbitrary: markerArbitrary.map((marker) => [marker]) },
      { weight: 1, arbitrary: atomArbitrary.map((atom) => [atom]) },
      { weight: 1, arbitrary: hyperlinkArbitrary.map((link): ParagraphContent[] => [link]) },
      { weight: 1, arbitrary: trackedArbitrary.map((change) => [change]) },
      { weight: 1, arbitrary: wrapperArbitrary.map((wrapper): ParagraphContent[] => [wrapper]) },
      { weight: 1, arbitrary: inlineSdtArbitrary.map((control): ParagraphContent[] => [control]) },
    ),
    { maxLength: 5 },
  )
  .map((groups) => groups.flat());

const sectionPropertiesArbitrary: fc.Arbitrary<SectionProperties> = fc.record(
  {
    pageWidth: fc.constantFrom(11906, 12240),
    sectionStart: fc.constantFrom("nextPage", "continuous"),
    preservedAttributes: preservedAttributesArbitrary,
  },
  { requiredKeys: ["pageWidth"] },
);

/** A paragraph without its id; {@link assignParagraphIds} names it. */
const paragraphFields = {
  type: fc.constant("paragraph" as const),
  content: paragraphContentArbitrary,
  textId: fc.constant("77777777"),
  formatting: paragraphFormattingArbitrary,
  preservedAttributes: preservedAttributesArbitrary,
  propertyChanges: trackedInfoArbitrary.map((info) => [
    {
      type: "paragraphPropertyChange" as const,
      info,
      previousFormatting: { alignment: "end" as const },
    },
  ]),
  pPrMark: trackedInfoArbitrary.map((info) => ({ kind: "ins" as const, info })),
  renderedPageBreakBefore: fc.constant(true),
  listRendering: fc.constant({ marker: "1.", level: 0, numId: 1, isBullet: false }),
};

const paragraphArbitrary: fc.Arbitrary<Paragraph> = fc.oneof(
  { weight: 6, arbitrary: fc.record(paragraphFields, { requiredKeys: ["type", "content"] }) },
  // A section break is rarer than the other fields, as in real documents.
  {
    weight: 1,
    arbitrary: fc.record(
      { ...paragraphFields, sectionProperties: sectionPropertiesArbitrary },
      { requiredKeys: ["type", "content", "sectionProperties"] },
    ),
  },
);

const nestedParagraphs = fc.array(paragraphArbitrary, { minLength: 1, maxLength: 2 });

const blockArbitrary: fc.Arbitrary<BlockContent> = fc.oneof(
  { weight: 8, arbitrary: paragraphArbitrary },
  {
    weight: 1,
    arbitrary: fc
      .array(fc.array(nestedParagraphs, { minLength: 1, maxLength: 2 }), {
        minLength: 1,
        maxLength: 2,
      })
      .map(
        (rows): BlockContent => ({
          type: "table",
          columnWidths: [2400, 2400],
          rows: rows.map((cells) => ({
            type: "tableRow",
            cells: cells.map((content) => ({ type: "tableCell", content })),
            preservedAttributes: [{ name: "rsidTr", value: "00ABCDEF" }],
          })),
        }),
      ),
  },
  {
    weight: 1,
    arbitrary: nestedParagraphs.map(
      (content): BlockContent => ({
        type: "blockSdt",
        properties: { sdtType: "group", id: 9 },
        content,
      }),
    ),
  },
  {
    weight: 1,
    arbitrary: fc.constant<BlockContent>({ type: "preservedBlock", xml: "<w:altChunk/>" }),
  },
  {
    weight: 1,
    arbitrary: fc.constant<BlockContent>({ type: "bookmarkStart", id: 90, name: "_GoBack" }),
  },
);

const toHexId = (value: number): string => value.toString(16).toUpperCase().padStart(8, "0");

/** Name every paragraph in the block tree, in document order, from `0x00000001`. */
const assignParagraphIds = (blocks: readonly BlockContent[]): BlockContent[] => {
  let next = 1;
  const name = (list: readonly BlockContent[]): BlockContent[] =>
    list.map((block): BlockContent => {
      switch (block.type) {
        case "paragraph":
          return { ...block, paraId: toHexId(next++) };
        case "table":
          return {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({ ...cell, content: name(cell.content) })),
            })),
          };
        case "blockSdt":
          return { ...block, content: name(block.content) };
        default:
          return block;
      }
    });
  return name(blocks);
};

/** The parser's per-section view: a section ends at each paragraph carrying one. */
const buildSections = (
  content: readonly BlockContent[],
  finalSectionProperties: SectionProperties,
): Section[] => {
  const sections: Section[] = [];
  let current: BlockContent[] = [];
  for (const block of content) {
    current.push(block);
    if (block.type === "paragraph" && block.sectionProperties !== undefined) {
      sections.push({ properties: block.sectionProperties, content: current });
      current = [];
    }
  }
  sections.push({ properties: finalSectionProperties, content: current });
  return sections;
};

export const documentArbitrary: fc.Arbitrary<Document> = fc
  .tuple(paragraphArbitrary, fc.array(blockArbitrary, { maxLength: 5 }), fc.nat())
  .map(([paragraph, blocks, at]): Document => {
    const withParagraph = [...blocks];
    withParagraph.splice(at % (blocks.length + 1), 0, paragraph);
    const content = assignParagraphIds(withParagraph);
    const finalSectionProperties: SectionProperties = { pageWidth: 12240, pageHeight: 15840 };
    const body: DocumentBody = {
      content,
      sections: buildSections(content, finalSectionProperties),
      finalSectionProperties,
      comments: [
        {
          id: 7,
          author: "A",
          content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
        },
      ],
    };
    return {
      package: {
        document: body,
        settings: { defaultTabStop: 720 },
        properties: { title: "synthetic" },
      },
      warnings: ["kept"],
    };
  });

/** Random numbers an operation is drawn from once the document it targets is known. */
export type OpSeed = {
  kind: number;
  block: number;
  first: number;
  second: number;
  third: number;
  /** When set, positions state `zeroWidthBefore`, drawn from these. */
  zeroWidth: { first: number; second: number } | undefined;
  depth: number;
  text: string;
  formatting: TextFormatting;
  inherit: boolean;
  runPatch: RunPropsPatch;
  paragraphPatch: ParagraphPropsPatch;
  newParagraph: SplitParagraphFields | undefined;
  content: ParagraphContent[];
  fresh: number;
};

const runPatchArbitrary: fc.Arbitrary<RunPropsPatch> = fc.record(
  {
    bold: fc.option(fc.boolean(), { nil: null }),
    italic: fc.option(fc.boolean(), { nil: null }),
    fontSize: fc.option(fc.integer({ min: 16, max: 28 }), { nil: null }),
    preserved: fc.constant(null),
  },
  { requiredKeys: [] },
);

const paragraphPatchArbitrary: fc.Arbitrary<ParagraphPropsPatch> = fc.record(
  {
    alignment: fc.option(fc.constantFrom("start" as const, "center" as const), { nil: null }),
    keepNext: fc.option(fc.boolean(), { nil: null }),
    styleId: fc.option(fc.constant("Heading2"), { nil: null }),
  },
  { requiredKeys: [] },
);

const newParagraphArbitrary: fc.Arbitrary<SplitParagraphFields> = fc.record(
  {
    formatting: paragraphFormattingArbitrary,
    textId: fc.constant("66666666"),
    preservedAttributes: preservedAttributesArbitrary,
  },
  { requiredKeys: [] },
);

const insertableContentArbitrary: fc.Arbitrary<ParagraphContent[]> = fc.array(
  fc.oneof(
    { weight: 4, arbitrary: runItem },
    { weight: 1, arbitrary: markerArbitrary },
    { weight: 1, arbitrary: atomArbitrary },
    { weight: 1, arbitrary: hyperlinkArbitrary },
  ),
  { minLength: 1, maxLength: 3 },
);

export const opSeedArbitrary: fc.Arbitrary<OpSeed> = fc.record({
  kind: fc.nat(),
  block: fc.nat(),
  first: fc.nat(),
  second: fc.nat(),
  third: fc.nat(),
  zeroWidth: fc.option(fc.record({ first: fc.nat(), second: fc.nat() }), {
    nil: undefined,
    freq: 2,
  }),
  depth: fc.integer({ min: 1, max: 4 }),
  text: nonEmptyTextArbitrary,
  formatting: textFormattingArbitrary,
  inherit: fc.boolean(),
  runPatch: runPatchArbitrary,
  paragraphPatch: paragraphPatchArbitrary,
  newParagraph: fc.option(newParagraphArbitrary, { nil: undefined }),
  content: insertableContentArbitrary,
  fresh: fc.integer({ min: 0x10_00_00_00, max: 0x7f_ff_ff_00 }),
});

const OP_KINDS = [
  DOCUMENT_OP_TYPES.INSERT_TEXT,
  DOCUMENT_OP_TYPES.INSERT_CONTENT,
  DOCUMENT_OP_TYPES.DELETE_RANGE,
  DOCUMENT_OP_TYPES.SPLIT_INLINE,
  DOCUMENT_OP_TYPES.JOIN_INLINE,
  DOCUMENT_OP_TYPES.SET_RUN_PROPS,
  DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
  DOCUMENT_OP_TYPES.SPLIT_BLOCK,
  DOCUMENT_OP_TYPES.JOIN_BLOCKS,
  DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
] as const;

/** The operation kinds {@link opFor} draws. */
export const GENERATED_OP_KINDS: readonly DocumentOpType[] = OP_KINDS;

/**
 * An operation against `document`, drawn from `seed`. Positions fall inside
 * the paragraph they name, so most operations apply; the ones that do not
 * (a position inside a surrogate pair, a split through a tracked change, a
 * join across a section break, a slice whose open ends do not fit) exercise
 * refusals.
 */
export const opFor = (document: Document, seed: OpSeed): DocumentOp => {
  const paragraphs = storyParagraphs(document.package.document);
  const target = paragraphs[seed.block % paragraphs.length];
  if (target === undefined) {
    throw new Error("A synthetic document always has a paragraph.");
  }
  const { paragraph } = target;
  const blockId = paragraph.paraId ?? "";
  const length = paragraphLength(paragraph);
  const position = (offset: number, zeroWidth: number | undefined): TextPosition => {
    const base = { story: OP_STORIES.MAIN, blockId, offset };
    if (zeroWidth === undefined) {
      return base;
    }
    const available = zeroWidthLeavesAt(paragraph.content, offset).length;
    return { ...base, zeroWidthBefore: zeroWidth % (available + 1) };
  };
  const first = seed.first % (length + 1);
  const second = seed.second % (length + 1);
  const from = position(Math.min(first, second), seed.zeroWidth?.first);
  const to = position(Math.max(first, second), seed.zeroWidth?.second);
  const at = position(seed.third % (length + 1), seed.zeroWidth?.first);
  const kind = OP_KINDS[seed.kind % OP_KINDS.length] ?? DOCUMENT_OP_TYPES.INSERT_TEXT;
  switch (kind) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
      return {
        type: kind,
        at: position(seed.third % (length + 1), undefined),
        text: seed.text,
        runProps: seed.inherit ? INHERIT_RUN_PROPS : seed.formatting,
      };
    case DOCUMENT_OP_TYPES.INSERT_CONTENT: {
      // Half the time a slice cut from the paragraph itself, open ends and all.
      const text = paragraphLogicalText(paragraph);
      const splitsPair = (offset: number) =>
        /[\uD800-\uDBFF]/u.test(text.charAt(offset - 1)) &&
        /[\uDC00-\uDFFF]/u.test(text.charAt(offset));
      if (
        seed.inherit &&
        from.offset < to.offset &&
        !splitsPair(from.offset) &&
        !splitsPair(to.offset)
      ) {
        const { removed } = deleteBetween(
          paragraph.content,
          { offset: from.offset, zeroWidthBefore: from.zeroWidthBefore ?? 0 },
          { offset: to.offset, zeroWidthBefore: to.zeroWidthBefore ?? 0 },
        );
        return { type: kind, at, slice: removed };
      }
      return { type: kind, at, slice: { content: seed.content, openStart: 0, openEnd: 0 } };
    }
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return { type: kind, from, to };
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
      return { type: kind, at, depth: seed.depth };
    case DOCUMENT_OP_TYPES.JOIN_INLINE: {
      // Half the time the end of a run, where an alike run may follow.
      const ends = runGaps(paragraph.content);
      const end = ends[seed.third % Math.max(1, ends.length)];
      if (seed.inherit && end !== undefined) {
        return {
          type: kind,
          at: { story: OP_STORIES.MAIN, blockId, ...end.after },
          depth: (seed.depth % 2) + 1,
        };
      }
      return { type: kind, at, depth: seed.depth };
    }
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return { type: kind, from, to, patch: seed.runPatch };
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return { type: kind, story: OP_STORIES.MAIN, blockId, patch: seed.paragraphPatch };
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK: {
      const used = new Set(paragraphIdsIn(document.package));
      let fresh = seed.fresh;
      while (used.has(toHexId(fresh))) fresh += 1;
      return seed.newParagraph === undefined
        ? { type: kind, at, newBlockId: toHexId(fresh) }
        : { type: kind, at, newBlockId: toHexId(fresh), newParagraph: seed.newParagraph };
    }
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS: {
      const used = new Set(paragraphIdsIn(document.package));
      let fresh = seed.fresh;
      while (used.has(toHexId(fresh))) fresh += 1;
      const replacement: Paragraph = {
        ...paragraph,
        content: seed.content,
        paraId: seed.depth % 2 === 0 ? blockId : toHexId(fresh),
      };
      const follower = paragraphs.find(
        (candidate) =>
          candidate.index === target.index + 1 &&
          JSON.stringify(candidate.list) === JSON.stringify(target.list),
      );
      // Sometimes two paragraphs become one, which may move a section break.
      const expected =
        seed.inherit && follower !== undefined ? [paragraph, follower.paragraph] : [paragraph];
      return { type: kind, story: OP_STORIES.MAIN, expected, blocks: [replacement] };
    }
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS: {
      const followerOf = (location: (typeof paragraphs)[number]) =>
        paragraphs.find(
          (candidate) =>
            candidate.index === location.index + 1 &&
            JSON.stringify(candidate.list) === JSON.stringify(location.list),
        );
      const joinable = paragraphs.filter((location) => followerOf(location) !== undefined);
      const leading = joinable[seed.block % Math.max(1, joinable.length)];
      // One join in ten names a paragraph that does not follow, to exercise the refusal.
      const trailing =
        leading === undefined || seed.first % 10 === 0
          ? paragraphs[seed.second % paragraphs.length]
          : followerOf(leading);
      return {
        type: kind,
        story: OP_STORIES.MAIN,
        blockId: (leading ?? target).paragraph.paraId ?? "",
        nextBlockId: trailing?.paragraph.paraId ?? "",
        depth: seed.depth % 3,
      };
    }
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};
