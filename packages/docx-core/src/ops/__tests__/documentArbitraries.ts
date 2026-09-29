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
import { normalizeForOps } from "../contract";
import { IDENTITY_SPACES, packageIdentityKeys, paragraphIdsIn } from "../ids";
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
  type RevisionStamp,
  type RunPropsPatch,
  SPLIT_HALVES,
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
  // As the parser does: a trailing empty section only when there is no other.
  if (current.length > 0 || sections.length === 0) {
    sections.push({ properties: finalSectionProperties, content: current });
  }
  return sections;
};

/**
 * The same plain data with every revision id and content-control id unique,
 * as the seed contract requires: tracked changes and property changes by
 * their `info`, content controls by their properties.
 */
const withUniqueRecordIds = (
  value: unknown,
  next: { revision: number; control: number },
): unknown => {
  if (Array.isArray(value)) {
    return value.map((item) => withUniqueRecordIds(item, next));
  }
  if (typeof value !== "object" || value === null) {
    return value;
  }
  const out: Record<string, unknown> = {};
  for (const [key, field] of Object.entries(value)) out[key] = withUniqueRecordIds(field, next);
  const info = out["info"];
  if (
    typeof info === "object" &&
    info !== null &&
    typeof Reflect.get(info, "author") === "string"
  ) {
    out["info"] = Object.assign({}, info, { id: next.revision++ });
  }
  if (typeof out["sdtType"] === "string" && typeof out["id"] === "number") {
    out["id"] = next.control++;
  }
  return out;
};

const REVIEW_MARK_KINDS = ["ins", "del", "moveFrom", "moveTo"] as const;

/** First revision id of the deletions {@link reviewValid} nests: above every generated one. */
const NESTED_REVISION_IDS = 5000;

/**
 * Blocks as a review leaves them: marks of every kind, but none on a
 * paragraph that ends its container (the story body or a table cell) or ends
 * a section, and some tracked insertions holding a deletion of their first
 * run (another author deleted part of an insertion).
 */
const reviewValid = (
  blocks: readonly BlockContent[],
  choices: readonly number[],
): BlockContent[] => {
  let drawn = 0;
  const choose = (): number => choices[drawn++ % Math.max(1, choices.length)] ?? 0;
  let nestedId = NESTED_REVISION_IDS;
  const nestDeletion = (item: ParagraphContent): ParagraphContent => {
    if ((item.type !== "insertion" && item.type !== "moveTo") || choose() % 2 !== 0) {
      return item;
    }
    const [first, ...rest] = item.content;
    if (first?.type !== "run") {
      return item;
    }
    const deletion: TrackedRunContent = {
      type: "deletion",
      info: { id: nestedId++, author: "B", date: "2026-01-02T03:04:05Z" },
      content: [first],
    };
    return { ...item, content: [deletion, ...rest] };
  };
  const visit = (list: readonly BlockContent[], endsContainer: boolean): BlockContent[] =>
    list.map((block, index): BlockContent => {
      const last = index === list.length - 1;
      switch (block.type) {
        case "paragraph": {
          const paragraph: Paragraph = { ...block, content: block.content.map(nestDeletion) };
          const mark = paragraph.pPrMark;
          if (mark === undefined) return paragraph;
          paragraph.pPrMark = { ...mark, kind: REVIEW_MARK_KINDS[choose() % 4] ?? "ins" };
          if ((last && endsContainer) || paragraph.sectionProperties !== undefined) {
            delete paragraph.pPrMark;
          }
          return paragraph;
        }
        case "table":
          return {
            ...block,
            rows: block.rows.map((row) => ({
              ...row,
              cells: row.cells.map((cell) => ({ ...cell, content: visit(cell.content, true) })),
            })),
          };
        case "blockSdt":
          return { ...block, content: visit(block.content, last && endsContainer) };
        default:
          return block;
      }
    });
  return visit(blocks, true);
};

/** Which documents a generator draws. */
type DocumentMode = "any" | "review";

const documentArbitraryIn = (mode: DocumentMode): fc.Arbitrary<Document> =>
  fc
    .tuple(
      paragraphArbitrary,
      fc.array(blockArbitrary, { maxLength: 5 }),
      fc.nat(),
      fc.boolean(),
      fc.array(fc.nat(), { minLength: 8, maxLength: 8 }),
    )
    .map(([paragraph, blocks, at, shareSections, choices]): Document => {
      const withParagraph = [...blocks];
      withParagraph.splice(at % (blocks.length + 1), 0, paragraph);
      // SAFETY: the renumbering keeps the shape of the plain data it is given.
      const named = withUniqueRecordIds(assignParagraphIds(withParagraph), {
        revision: 1,
        control: 1,
      }) as BlockContent[];
      const reviewed = mode === "review" ? reviewValid(named, choices) : named;
      return documentFrom(reviewed, shareSections);
    });

/**
 * A synthetic document around blocks. Half share their records between the
 * body and its section view, as the parser builds them; half hold an
 * independently built copy in the view, as a document read back from JSON
 * does. Paragraphs in a comment and a note share the id space.
 */
const documentFrom = (blocks: BlockContent[], shareSections: boolean): Document => {
  const content = normalizeForOps({ package: { document: { content: blocks } } }).package.document
    .content;
  const finalSectionProperties: SectionProperties = { pageWidth: 12240, pageHeight: 15840 };
  const sections = buildSections(content, finalSectionProperties);
  const body: DocumentBody = {
    content,
    sections: shareSections ? sections : independentCopy(sections),
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
      footnotes: [
        {
          type: "footnote",
          id: 2,
          content: [{ type: "paragraph", paraId: "7FFFFFF1", content: [] }],
        },
      ],
      settings: { defaultTabStop: 720 },
      properties: { title: "synthetic" },
    },
    warnings: ["kept"],
  };
};

/** Synthetic documents meeting the seed contract, marks and tracked changes of every shape. */
export const documentArbitrary: fc.Arbitrary<Document> = documentArbitraryIn("any");

/**
 * Review-valid synthetic documents: as {@link documentArbitrary}, with marks
 * of every kind but none on a paragraph that ends its container or a
 * section, and deletions nested in insertions.
 */
export const reviewDocumentArbitrary: fc.Arbitrary<Document> = documentArbitraryIn("review");

/** A structurally equal copy that shares no record with its source. */
export const independentCopy = <Value>(value: Value): Value => {
  // SAFETY: JSON round-trips the plain data these fixtures are made of.
  const copy = JSON.parse(JSON.stringify(value)) as Value;
  return copy;
};

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
  // Normalized as a seeded document is; an empty slice then exercises the refusal.
  content: insertableContentArbitrary.map((content): ParagraphContent[] => {
    const [paragraph] = normalizeForOps({
      package: { document: { content: [{ type: "paragraph", content }] } },
    }).package.document.content;
    return paragraph?.type === "paragraph" ? paragraph.content : [];
  }),
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
  // One operation in five names no new ids, so a cut through an identified
  // record is refused for want of them.
  const ids =
    seed.depth % 5 === 0
      ? {}
      : {
          newIds: {
            revision: Array.from(
              { length: 6 },
              (_, index) => 10_000 + (seed.fresh % 100_000) * 8 + index,
            ),
            control: Array.from(
              { length: 3 },
              (_, index) => 10_000 + (seed.fresh % 100_000) * 4 + index,
            ),
          },
        };
  switch (kind) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
      return {
        type: kind,
        at: position(seed.third % (length + 1), undefined),
        text: seed.text,
        runProps: seed.inherit ? INHERIT_RUN_PROPS : seed.formatting,
        ...ids,
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
        return { type: kind, at, slice: removed, ...ids };
      }
      return {
        type: kind,
        at,
        slice: { content: seed.content, openStart: 0, openEnd: 0 },
        ...ids,
      };
    }
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return { type: kind, from, to };
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
      return { type: kind, at, depth: seed.depth, ...ids };
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
      return { type: kind, from, to, patch: seed.runPatch, ...ids };
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return { type: kind, story: OP_STORIES.MAIN, blockId, patch: seed.paragraphPatch };
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK: {
      const used = new Set(paragraphIdsIn(document.package));
      let fresh = seed.fresh;
      while (used.has(toHexId(fresh))) fresh += 1;
      // One split in three names the new half instead of taking the default.
      const halves = [SPLIT_HALVES.FIRST, SPLIT_HALVES.SECOND] as const;
      const newHalf = seed.depth % 3 === 0 ? { newHalf: halves[seed.first % 2] } : {};
      return seed.newParagraph === undefined
        ? { type: kind, at, newBlockId: toHexId(fresh), ...newHalf, ...ids }
        : {
            type: kind,
            at,
            newBlockId: toHexId(fresh),
            newParagraph: seed.newParagraph,
            ...newHalf,
            ...ids,
          };
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
        // One join in three keeps the first paragraph instead of the second.
        ...(seed.third % 3 === 0 ? { survivor: SPLIT_HALVES.FIRST } : {}),
      };
    }
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};

const TRACKED_OP_KINDS = [
  DOCUMENT_OP_TYPES.INSERT_TEXT,
  DOCUMENT_OP_TYPES.INSERT_CONTENT,
  DOCUMENT_OP_TYPES.DELETE_RANGE,
  DOCUMENT_OP_TYPES.SET_RUN_PROPS,
  DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
  DOCUMENT_OP_TYPES.SPLIT_BLOCK,
  DOCUMENT_OP_TYPES.JOIN_BLOCKS,
] as const;

/** The operation kinds {@link trackedOpFor} draws. */
export const GENERATED_TRACKED_OP_KINDS: readonly DocumentOpType[] = TRACKED_OP_KINDS;

/** A revision id no record in the document carries. */
export const unusedRevisionId = (document: Document): number => {
  const prefix = `${IDENTITY_SPACES.REVISION}:`;
  let largest = 0;
  for (const key of packageIdentityKeys(document.package)) {
    if (key.startsWith(prefix)) largest = Math.max(largest, Number(key.slice(prefix.length)));
  }
  return largest + 1;
};

const twoDigits = (value: number): string => String(value % 60).padStart(2, "0");

/**
 * The stamp of the `index`th tracked operation of a run: a date no generated
 * record and no other operation of the run carries, so no operation merges
 * into a change that was there before it.
 */
export const stampFor = (document: Document, seed: OpSeed, index: number): RevisionStamp => {
  const stamp: RevisionStamp = {
    id: unusedRevisionId(document),
    author: seed.inherit ? "A" : "C",
    date: `2026-03-04T05:${twoDigits(Math.floor(index / 60))}:${twoDigits(index)}Z`,
  };
  if (seed.fresh % 3 === 0) {
    stamp.initials = stamp.author;
  }
  return stamp;
};

/** A paragraph with no mark change and the one directly after it, when there is such a pair. */
const unmarkedJoin = (
  document: Document,
  seed: OpSeed,
): { blockId: string; nextBlockId: string } | undefined => {
  const paragraphs = storyParagraphs(document.package.document);
  const pairs = paragraphs.flatMap((location) => {
    const next = paragraphs.find(
      (candidate) =>
        candidate.index === location.index + 1 &&
        JSON.stringify(candidate.list) === JSON.stringify(location.list),
    );
    // Half the time also one whose property change the join's own would not meet.
    return next === undefined ||
      location.paragraph.pPrMark !== undefined ||
      location.paragraph.sectionProperties !== undefined ||
      (seed.inherit && (next.paragraph.propertyChanges?.length ?? 0) > 0)
      ? []
      : [{ blockId: location.paragraph.paraId ?? "", nextBlockId: next.paragraph.paraId ?? "" }];
  });
  return pairs[seed.block % Math.max(1, pairs.length)];
};

/**
 * A tracked operation against `document`: one of {@link opFor}'s, of a kind
 * that takes a stamp, carrying the stamp for the `index`th operation of a run
 * and new ids for the records it creates past the first (one operation in
 * five names none, exercising the refusal).
 */
export const trackedOpFor = (document: Document, seed: OpSeed, index = 0): DocumentOp => {
  const kind =
    TRACKED_OP_KINDS[seed.kind % TRACKED_OP_KINDS.length] ?? DOCUMENT_OP_TYPES.INSERT_TEXT;
  const op = opFor(document, { ...seed, kind: OP_KINDS.indexOf(kind) });
  const revision = stampFor(document, seed, index);
  const pool = (base: number, length: number) =>
    Array.from({ length }, (_, offset) => base + (seed.fresh % 100_000) * 8 + offset);
  const ids =
    seed.depth % 5 === 0 ? {} : { newIds: { revision: pool(20_000, 8), control: pool(10_000, 3) } };
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return { ...op, ...ids, revision };
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS: {
      // Mostly a paragraph whose mark carries no change: a tracked join refuses the others.
      const unmarked = seed.first % 4 === 0 ? undefined : unmarkedJoin(document, seed);
      // A tracked join always leaves the second paragraph.
      const join = { ...op, ...unmarked, ...ids, revision };
      Reflect.deleteProperty(join, "survivor");
      return join;
    }
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return { ...op, revision };
    default:
      return op;
  }
};
