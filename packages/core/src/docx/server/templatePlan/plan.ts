/**
 * A template plan checked against its source and template and resolved into
 * what the output must hold: per output paragraph the expected text and
 * emphasis, derived from the source blocks alone. The writer emits it and the
 * verifier compares the written package against it.
 */
import { Result } from "better-result";

import { parseFieldInstruction } from "../../fieldParser";
import { computeListRendering } from "../../numberingParser";
import { NAMESPACES, type XmlElement } from "../../xmlParser";
import type { AnalyzedBlock, AnalyzedDocument } from "./blocks";
import {
  bookmarkIdsOf,
  type BookmarkWrap,
  mintRefBookmarkName,
  refFieldInstruction,
} from "./builders";
import {
  emphasisDeviation,
  emphasisSignature,
  paragraphBaseline,
  readDirectEmphasis,
  withCharacterStyle,
  withOverride,
  withSignatureHighlight,
} from "./emphasis";
import { type ParagraphLabel, styleNumbering } from "./labels";
import { PLAN_HIGHLIGHT_COLORS, planShapeIssue } from "./planShape";
import type { StoryOrigin } from "./resources";
import { DELETED_PARAGRAPH_MARK } from "./revisions";
import type {
  AtomKind,
  AtomLink,
  InlineAtom,
  ParagraphRead,
  StoryField,
} from "./story";
import {
  type TemplatePlan,
  type TemplatePlanAlignment,
  TemplatePlanError,
  type TemplatePlanErrorCode,
  type TemplatePlanHighlight,
  type TemplatePlanIndent,
  type TemplatePlanMerge,
  type TemplatePlanNumbering,
  type TemplatePlanOmitReason,
  type TemplatePlanReference,
  type TemplatePlanSection,
  type TemplatePlanSourceItem,
} from "./types";
import {
  elementChildren,
  isW,
  maxOf,
  walkElements,
  wAttr,
  wChild,
  wChildVal,
  wName,
} from "./xml";

export type ExpectedPiece = {
  kind: AtomKind;
  text: string;
  /** The emphasis the piece states over its paragraph's style. */
  signature: string;
  link: AtomLink | null;
  /** The source atom, for what only the source can supply: notes, objects, symbols. */
  atom: InlineAtom | null;
  /** The part the atom was read from: the main document or a notes part. */
  origin: PieceOrigin;
};

export type PieceOrigin = StoryOrigin;

export type ExpectedSegment =
  | { kind: "piece"; piece: ExpectedPiece }
  | {
      kind: "reference";
      reference: TemplatePlanReference;
      pieces: ExpectedPiece[];
      replacedText: string;
    };

export type NumberingMint = {
  key: string;
  numId: number;
  /** The template list the instance copies its level overrides from. */
  baseNumId: number;
  abstractNumId: number;
  /** The number each restarted level begins with. */
  starts: Map<number, number>;
};

export type ResolvedNumbering =
  | { kind: "style" }
  | { kind: "none" }
  | { kind: "direct"; numId: number; ilvl: number };

export type ResolvedPart = {
  block: AnalyzedBlock;
  strippedPrefix: string;
  /** The lead a slot dropped off the part and the template paragraph that
   *  shows the field, or null. */
  droppedLead: { text: string; shownByTemplateBlockId: string } | null;
  /** The automatic label written as text before the part, or null. */
  labelAsText: string | null;
};

export type ResolvedParagraph = {
  parts: ResolvedPart[];
  /** A paragraph style of the template; null leaves the paragraph unstyled. */
  styleId: string | null;
  numbering: ResolvedNumbering;
  segments: ExpectedSegment[];
  text: string;
  /** One signature per UTF-16 unit of `text`. */
  signatures: string[];
  inlineObjects: number;
  placement: "body" | "nested" | "slot";
  /** Whether the paragraph starts on a new page. */
  pageBreakBefore: boolean;
  /** Whether the paragraph is kept with the one after it. */
  keepNext: boolean;
  /** The paragraph's own indent, or null. */
  indent: TemplatePlanIndent | null;
  /** The paragraph's own left tab stops, or null. */
  tabStops: readonly number[] | null;
  /** The paragraph's own alignment, or null. */
  alignment: TemplatePlanAlignment | null;
  /** The section the paragraph is the last of, or null. */
  endsSection: TemplatePlanSection | null;
};

export type ResolvedItem =
  | { kind: "template"; block: AnalyzedBlock }
  /**
   * An empty paragraph the writer adds, in no plan: between two tables that
   * would otherwise stand next to each other and read as one, and after a
   * table that would end the body.
   */
  | { kind: "separator" }
  | { kind: "paragraph"; paragraph: ResolvedParagraph }
  | {
      kind: "table";
      block: AnalyzedBlock;
      styleId: string | null;
      /** The resolved paragraph of each nested `w:p`. */
      cells: Map<XmlElement, ResolvedParagraph>;
    }
  | {
      kind: "slot";
      templateBlock: AnalyzedBlock;
      paragraph: ResolvedParagraph;
    };

export type ResolvedSlot = {
  templateBlock: AnalyzedBlock;
  paragraph: ResolvedParagraph;
};

const fieldCharOf = (run: XmlElement): string | null =>
  wName(run) === "r" ? wAttr(wChild(run, "fldChar"), "fldCharType") : null;

/**
 * The run of a paragraph that holds the separator of the field the
 * paragraph begins and leaves open: what follows it is the start of a result
 * that goes on in later paragraphs. A field the paragraph both begins and
 * ends (a page number in front of a table of contents) is not that field.
 * Null when no one field is left open with its result begun, in runs of the
 * paragraph itself.
 */
export const fieldSeparatorRun = (paragraph: XmlElement): XmlElement | null => {
  const open: { separator: XmlElement | null }[] = [];
  for (const child of elementChildren(paragraph)) {
    const mark = fieldCharOf(child);
    if (mark === "begin") {
      open.push({ separator: null });
    } else if (mark === "end") {
      open.pop();
    } else if (mark === "separate") {
      const top = open.at(-1);
      if (top !== undefined && top.separator === null) {
        top.separator = child;
      }
    }
  }
  const [left, ...others] = open;
  return left === undefined || others.length > 0 ? null : left.separator;
};

const holdsText = (element: XmlElement): boolean => {
  let text = false;
  walkElements(element, (inner) => {
    text ||= wName(inner) === "t";
  });
  return text;
};

/**
 * Whether a paragraph ends the field left open in front of it and holds no
 * more behind that end: the end of the field's result and nothing else. The
 * end is the field's own, the first one that closes no field the paragraph
 * began; text behind it stays, whatever field follows that text.
 */
const endsOpenField = (paragraph: XmlElement): boolean => {
  const children = elementChildren(paragraph);
  let depth = 0;
  for (const [position, child] of children.entries()) {
    const mark = fieldCharOf(child);
    if (mark === "begin") {
      depth += 1;
    } else if (mark === "end") {
      if (depth === 0) {
        return !children.slice(position + 1).some(holdsText);
      }
      depth -= 1;
    }
  }
  return false;
};

/**
 * The kept tables of contents that show entries over several paragraphs:
 * the paragraph each begins in, the length of the text that paragraph holds
 * in front of the field's result, and the further paragraphs of its
 * entries. A field that does not begin in runs of its own paragraph, that
 * ends in a paragraph with text behind it, or whose paragraphs are not all
 * kept top-level blocks, is not among them and stays as the template has it.
 */
const contentsWithEntries = (
  template: AnalyzedDocument,
  keptIds: ReadonlySet<string>,
): { first: AnalyzedBlock; textKept: number; rest: AnalyzedBlock[] }[] => {
  const found: {
    first: AnalyzedBlock;
    textKept: number;
    rest: AnalyzedBlock[];
  }[] = [];
  for (const field of template.story.state.fields) {
    if (
      field.kind !== "TOC" ||
      !field.closed ||
      field.nestedInCode ||
      field.endParagraphOrdinal <= field.paragraphOrdinal
    ) {
      continue;
    }
    const [first, ...rest] = template.story.paragraphs.slice(
      field.paragraphOrdinal,
      field.endParagraphOrdinal + 1,
    );
    const last = rest.at(-1);
    if (
      first !== undefined &&
      last !== undefined &&
      [first, ...rest].every(
        ({ block, parentId }) => parentId === null && keptIds.has(block.id),
      ) &&
      field.resultStart >= 0 &&
      fieldSeparatorRun(first.element) !== null &&
      endsOpenField(last.element)
    ) {
      found.push({ first, textKept: field.resultStart, rest });
    }
  }
  return found;
};

export type ResolvedPlan = {
  items: ResolvedItem[];
  /** Slots that sit inside a kept template table. */
  nestedSlots: ResolvedSlot[];
  /** Paragraphs of kept template tables written without their text. */
  emptied: AnalyzedBlock[];
  /** The output paragraph each placed source block ends up in. */
  byBlockId: Map<string, ResolvedParagraph>;
  omitted: Map<string, TemplatePlanOmitReason>;
  mints: NumberingMint[];
  /** The bookmark of every paragraph a reference targets. */
  bookmarks: Map<ResolvedParagraph, BookmarkWrap>;
  references: {
    reference: TemplatePlanReference;
    paragraph: ResolvedParagraph;
    target: ResolvedParagraph;
    replacedText: string;
    instruction: string;
  }[];
  freezeKinds: ReadonlySet<string>;
  keptTemplateBlocks: AnalyzedBlock[];
  /** The paragraphs kept tables of contents begin in, written without the
   *  entries they cache: each with the length of the text it keeps in front
   *  of the field's result. */
  emptiedContents: Map<AnalyzedBlock, number>;
};

type Refusal = {
  code: TemplatePlanErrorCode;
  message: string;
  blockId?: string;
};

const refuse = ({ code, message, blockId }: Refusal): TemplatePlanError =>
  new TemplatePlanError({
    code,
    message,
    ...(blockId === undefined ? {} : { blockId }),
  });

/** Where a label may end: at white space or a line break. */
const LABEL_END = /^[ \t\u00A0\n]$/u;
const MAX_LABEL_LENGTH = 40;
const ROMAN_NUMERAL =
  /^(?=.)M{0,3}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3})$/iu;
const LABEL_WORD = /^(?:\p{L}[\p{L}\p{M}]*\.?|§{1,2}|№|#)$/u;
const LABEL_DESIGNATOR =
  /^[([]?(?<core>\d+(?:[.\-\u2013]\d+)*[a-z]?|\p{L}+)[)\]]?[.:)]*$/u;
const LABEL_BULLET = /^[-\u2013\u2014\u2022\u00B7\u25AA\u25E6*]$/u;
/** A number of several parts typed with blanks between them: `1. 1.`. */
const LABEL_SPACED_NUMBER = /^\d{1,3}\.(?:[ \u00A0]+\d{1,3}\.){1,3}$/u;
const MAX_LABEL_WORDS = 3;

/**
 * Whether `label` reads as a manual list label: up to three leading words
 * ("Article", "§", "Annex No.") and one designator, a number, a roman numeral
 * or a letter with its brackets and punctuation; or a single bullet; or a
 * number whose parts stand apart (`1. 1.` for `1.1.`). This is the only text
 * a plan may have removed.
 */
export const isManualLabel = (label: string): boolean => {
  if (label === "" || label.length > MAX_LABEL_LENGTH) {
    return false;
  }
  if (LABEL_BULLET.test(label) || LABEL_SPACED_NUMBER.test(label)) {
    return true;
  }
  const tokens = label.split(/[ \u00A0]/u);
  const designator = tokens.at(-1) ?? "";
  const words = tokens.slice(0, -1);
  const core = LABEL_DESIGNATOR.exec(designator)?.groups?.["core"];
  return (
    core !== undefined &&
    words.length <= MAX_LABEL_WORDS &&
    words.every((word) => LABEL_WORD.test(word)) &&
    (/^\d/u.test(core) || core.length === 1 || ROMAN_NUMERAL.test(core))
  );
};

const NOTES_OF_REFERENCE: Partial<Record<AtomKind, "footnotes" | "endnotes">> =
  {
    footnoteReference: "footnotes",
    endnoteReference: "endnotes",
  };

const LABEL_SPACE = /^[ \t ]+/u;
/** The blank behind a label when a line break behind it goes too. */
const LABEL_SPACE_AND_BREAK = /^[ \t ]*(?:\n[ \t ]*)?/u;
const LEADING_SPACE = /^[ \t ]*/u;

/** Whether a block holds no text and no object. */
export const isBlankBlock = ({ block }: AnalyzedBlock): boolean =>
  block.text.trim() === "" && block.inlineObjects === 0;

/** Whether a block, or a paragraph inside it, renders a list label. */
const showsLabel = ({ block }: AnalyzedBlock): boolean =>
  block.label !== null ||
  (block.nested ?? []).some(({ label }) => label !== null);

export type SignatureContext = Pick<
  AnalyzedDocument,
  "styles" | "defaultStyleId" | "defaultEmphasis"
>;

export type ParagraphPiecesOptions = {
  read: ParagraphRead;
  /** Styles of the document the paragraph was read from. */
  context: SignatureContext;
  origin: PieceOrigin;
  /** The fields of the story the paragraph belongs to. */
  fields: readonly StoryField[];
};

const HYPERLINK_FIELD = "HYPERLINK";
const LOCATION_SWITCH = "l";

/** The link a HYPERLINK field makes: its address and the place in it, if any. */
export const hyperlinkFieldLink = ({
  instruction,
  kind,
}: StoryField): AtomLink | null => {
  if (kind !== HYPERLINK_FIELD) {
    return null;
  }
  const { argument, switches } = parseFieldInstruction(instruction);
  // Without an address the field points into its own document.
  if (argument === undefined || argument === "") {
    return null;
  }
  const location = switches.find(
    ({ switch: name }) => name.toLowerCase() === LOCATION_SWITCH,
  )?.value;
  return { kind: "url", url: argument, anchor: location ?? null };
};

const NOTE_REFERENCE_KINDS: ReadonlySet<AtomKind> = new Set([
  "footnoteReference",
  "endnoteReference",
]);
/** What a stripped label may consist of. */
const STRIPPABLE_KINDS: ReadonlySet<AtomKind> = new Set(["text", "tab"]);
/** The same when a line break behind the label goes with it. */
const STRIPPABLE_WITH_BREAK: ReadonlySet<AtomKind> = new Set([
  ...STRIPPABLE_KINDS,
  "lineBreak",
]);
/**
 * Character styles that only give a link its look. The template supplies
 * that look; any other character style inside a link is emphasis.
 */
const LINK_STYLE_NAMES: ReadonlySet<string> = new Set([
  "hyperlink",
  "followedhyperlink",
  "internet link",
  "visited internet link",
]);

/** The pieces of one source paragraph, each with the emphasis it states. */
export const paragraphPieces = ({
  read,
  context: { styles, defaultStyleId, defaultEmphasis },
  origin,
  fields,
}: ParagraphPiecesOptions): ExpectedPiece[] => {
  const baseline = paragraphBaseline(
    styles.get(read.styleId ?? defaultStyleId ?? ""),
    defaultEmphasis,
  );
  return read.atoms.map((atom) => {
    // A HYPERLINK field is a link like any other once its result is kept.
    const fieldLink = atom.fields
      .map((index) => fields[index])
      .map((field) => (field === undefined ? null : hyperlinkFieldLink(field)))
      .find((link) => link !== null);
    const link: AtomLink | null = atom.link ?? fieldLink ?? null;
    let signature = "";
    if (!NOTE_REFERENCE_KINDS.has(atom.kind)) {
      const characterStyle = styles.get(
        wChildVal(atom.runProperties, "rStyle") ?? "",
      );
      const isLinkLook =
        link !== null &&
        LINK_STYLE_NAMES.has(characterStyle?.name?.toLowerCase() ?? "");
      signature = emphasisSignature(
        emphasisDeviation(
          withOverride(
            isLinkLook
              ? baseline
              : withCharacterStyle(baseline, characterStyle),
            readDirectEmphasis(atom.runProperties),
          ),
          baseline,
        ),
      );
    }
    return {
      kind: atom.kind,
      text: atom.text,
      signature,
      link,
      atom,
      origin,
    };
  });
};

type PositionedPiece = { piece: ExpectedPiece; start: number; end: number };

const isHighSurrogate = (unit: number): boolean =>
  unit >= 0xd800 && unit <= 0xdbff;
const isLowSurrogate = (unit: number): boolean =>
  unit >= 0xdc00 && unit <= 0xdfff;

/** Whether `offset` falls between the two units of one character. */
const splitsCharacter = (text: string, offset: number): boolean =>
  offset > 0 &&
  offset < text.length &&
  isHighSurrogate(text.charCodeAt(offset - 1)) &&
  isLowSurrogate(text.charCodeAt(offset));

/**
 * Split text pieces at the offsets in `cuts`. A cut inside a character, or
 * inside a piece that is not text, is not made; the caller finds it missing
 * among the boundaries.
 */
const splitPieces = (
  pieces: readonly ExpectedPiece[],
  cuts: readonly number[],
): PositionedPiece[] => {
  const out: PositionedPiece[] = [];
  let offset = 0;
  for (const piece of pieces) {
    const start = offset;
    const end = start + piece.text.length;
    offset = end;
    const inside =
      piece.kind === "text"
        ? cuts.filter(
            (cut) =>
              cut > start &&
              cut < end &&
              !splitsCharacter(piece.text, cut - start),
          )
        : [];
    if (inside.length === 0) {
      out.push({ piece, start, end });
      continue;
    }
    let cursor = start;
    for (const cut of [...new Set(inside)].toSorted((a, b) => a - b)) {
      out.push({
        piece: {
          ...piece,
          text: piece.text.slice(cursor - start, cut - start),
        },
        start: cursor,
        end: cut,
      });
      cursor = cut;
    }
    out.push({
      piece: { ...piece, text: piece.text.slice(cursor - start) },
      start: cursor,
      end,
    });
  }
  return out;
};

type Span = { start: number; end: number };

const isSpanWithin = (
  { start, end }: Span,
  from: number,
  length: number,
): boolean =>
  Number.isInteger(start) &&
  Number.isInteger(end) &&
  start >= from &&
  start < end &&
  end <= length;

const spansOverlap = (spans: readonly Span[]): boolean => {
  const ordered = spans.toSorted((a, b) => a.start - b.start);
  return ordered.some((span, index) => {
    const previous = ordered[index - 1];
    return previous !== undefined && span.start < previous.end;
  });
};

type PartSegmentsOptions = {
  block: AnalyzedBlock;
  pieces: ExpectedPiece[];
  strippedLength: number;
  /** Whether what is stripped may hold a line break. */
  strippedBreak: boolean;
  references: readonly TemplatePlanReference[];
  highlights: readonly TemplatePlanHighlight[];
};

const partSegments = ({
  block,
  pieces,
  strippedLength,
  strippedBreak,
  references,
  highlights,
}: PartSegmentsOptions): Result<ExpectedSegment[], TemplatePlanError> => {
  const strippable = strippedBreak ? STRIPPABLE_WITH_BREAK : STRIPPABLE_KINDS;
  const blockId = block.block.id;
  const length = block.block.text.length;
  const badReference = (message: string): Result<never, TemplatePlanError> =>
    Result.err(refuse({ code: "reference-span-invalid", message, blockId }));
  const badHighlight = (message: string): Result<never, TemplatePlanError> =>
    Result.err(refuse({ code: "highlight-span-invalid", message, blockId }));
  for (const reference of references) {
    if (!isSpanWithin(reference, strippedLength, length)) {
      return badReference(
        `Reference span ${reference.start}-${reference.end} lies outside block ${blockId}.`,
      );
    }
  }
  for (const highlight of highlights) {
    if (
      !isSpanWithin(highlight, strippedLength, length) ||
      !PLAN_HIGHLIGHT_COLORS.has(highlight.color)
    ) {
      return badHighlight(
        `Highlight ${highlight.start}-${highlight.end} (${highlight.color}) does not fit block ${blockId}.`,
      );
    }
  }
  if (spansOverlap(references)) {
    return badReference(`Reference spans overlap in block ${blockId}.`);
  }
  if (spansOverlap(highlights)) {
    return badHighlight(`Highlight spans overlap in block ${blockId}.`);
  }
  const ordered = references.toSorted((a, b) => a.start - b.start);

  const split = splitPieces(pieces, [
    strippedLength,
    ...references.flatMap(({ start, end }) => [start, end]),
    ...highlights.flatMap(({ start, end }) => [start, end]),
  ]);
  // Every offset the plan names must fall between two pieces: not inside a
  // character, a symbol or a mark that cannot be cut.
  const boundaries = new Set([0, ...split.map(({ end }) => end)]);
  if (!boundaries.has(strippedLength)) {
    return Result.err(
      refuse({
        code: "strip-label-mismatch",
        message: `The label to strip from block ${blockId} ends inside a character or mark.`,
        blockId,
      }),
    );
  }
  for (const { start, end } of references) {
    if (!boundaries.has(start) || !boundaries.has(end)) {
      return badReference(
        `Reference span ${start}-${end} in block ${blockId} cuts a character or mark.`,
      );
    }
  }
  for (const { start, end } of highlights) {
    if (!boundaries.has(start) || !boundaries.has(end)) {
      return badHighlight(
        `Highlight ${start}-${end} in block ${blockId} cuts a character or mark.`,
      );
    }
  }

  // A tracked change is carried as it stands. A label that is stripped or
  // a reference that becomes a field would take the change with it, or cut
  // it in two, so neither may reach one.
  for (const { piece, start, end } of split) {
    if (piece.atom?.revision === undefined) {
      continue;
    }
    if (strippedLength > 0 && start < strippedLength) {
      return Result.err(
        refuse({
          code: "strip-label-mismatch",
          message: `The label to strip from block ${blockId} holds a tracked change.`,
          blockId,
        }),
      );
    }
    // Touching counts: a deletion at the edge of a cited span has no side
    // of the field to stand on.
    const cited = references.find(
      (reference) => start <= reference.end && end >= reference.start,
    );
    if (cited !== undefined) {
      return badReference(
        `Reference span ${cited.start}-${cited.end} in block ${blockId} holds a tracked change.`,
      );
    }
  }

  const positioned: PositionedPiece[] = [];
  for (const entry of split) {
    const empty = entry.start === entry.end;
    if (entry.end <= strippedLength && !empty) {
      // A label is text and the space after it; anything else stays put.
      if (!strippable.has(entry.piece.kind)) {
        return Result.err(
          refuse({
            code: "strip-label-mismatch",
            message: `The label to strip from block ${blockId} covers more than text.`,
            blockId,
          }),
        );
      }
      continue;
    }
    const highlight = highlights.find(
      ({ start, end }) => entry.start >= start && entry.end <= end && !empty,
    );
    // A note reference mark carries no emphasis of its own.
    if (
      highlight !== undefined &&
      !NOTE_REFERENCE_KINDS.has(entry.piece.kind)
    ) {
      entry.piece = {
        ...entry.piece,
        signature: withSignatureHighlight(
          entry.piece.signature,
          highlight.color,
        ),
      };
    }
    positioned.push(entry);
  }

  const segments: ExpectedSegment[] = [];
  let open:
    { reference: TemplatePlanReference; pieces: ExpectedPiece[] } | undefined;
  const close = (): Result<undefined, TemplatePlanError> => {
    if (open === undefined) {
      return Result.ok(undefined);
    }
    const replacedText = open.pieces.map(({ text }) => text).join("");
    const first = open.pieces.at(0);
    const { reference } = open;
    if (first === undefined || replacedText === reference.expectedText) {
      segments.push({
        kind: "reference",
        reference,
        replacedText,
        pieces: open.pieces.map((piece) => ({ ...piece, link: null })),
      });
    } else if (
      open.pieces.some(({ signature }) => signature !== first.signature)
    ) {
      // New text has no positions to carry mixed emphasis over to.
      return badReference(
        `Reference span ${reference.start}-${reference.end} in block ${blockId} changes text that is emphasised in part.`,
      );
    } else {
      segments.push({
        kind: "reference",
        reference,
        replacedText,
        pieces: [
          {
            kind: "text",
            text: reference.expectedText,
            signature: first.signature,
            link: null,
            atom: null,
            origin: first.origin,
          },
        ],
      });
    }
    open = undefined;
    return Result.ok(undefined);
  };
  for (const { piece, start, end } of positioned) {
    const empty = start === end;
    const reference = ordered.find((candidate) =>
      empty
        ? start > candidate.start && start < candidate.end
        : start >= candidate.start && end <= candidate.end,
    );
    if (reference === undefined) {
      const closed = close();
      if (Result.isError(closed)) {
        return closed;
      }
      segments.push({ kind: "piece", piece });
      continue;
    }
    if (piece.kind !== "text") {
      return badReference(
        `Reference span ${reference.start}-${reference.end} in block ${blockId} covers more than text.`,
      );
    }
    if (empty) {
      continue;
    }
    if (open !== undefined && open.reference !== reference) {
      const closed = close();
      if (Result.isError(closed)) {
        return closed;
      }
    }
    open ??= { reference, pieces: [] };
    open.pieces.push(piece);
  }
  const closed = close();
  return Result.isError(closed) ? closed : Result.ok(segments);
};

const SEPARATOR_PIECES = {
  lineBreak: {
    kind: "lineBreak",
    text: "\n",
    signature: "",
    link: null,
    atom: null,
    origin: "document",
  },
  space: {
    kind: "text",
    text: " ",
    signature: "",
    link: null,
    atom: null,
    origin: "document",
  },
  none: null,
} as const satisfies Record<TemplatePlanMerge, ExpectedPiece | null>;

/** The pieces of a segment list in order, reference results included. */
export const segmentPieces = (
  segments: readonly ExpectedSegment[],
): ExpectedPiece[] =>
  segments.flatMap((segment) =>
    segment.kind === "piece" ? [segment.piece] : segment.pieces,
  );

/** The expected text and per-unit signatures of a piece list. */
export const piecesContent = (
  pieces: readonly ExpectedPiece[],
): { text: string; signatures: string[] } => {
  const signatures: string[] = [];
  let text = "";
  for (const piece of pieces) {
    text += piece.text;
    for (let unit = 0; unit < piece.text.length; unit += 1) {
      signatures.push(piece.signature);
    }
  }
  return { text, signatures };
};

const isTableItem = (item: ResolvedItem): boolean =>
  item.kind === "table" ||
  (item.kind === "template" && item.block.block.kind === "table");

/** Puts a separator between two neighbouring tables and after a last one. */
const partTables = (items: ResolvedItem[]): void => {
  for (let index = items.length - 1; index >= 0; index -= 1) {
    const item = items[index];
    const next = items[index + 1];
    if (
      item !== undefined &&
      isTableItem(item) &&
      (next === undefined || isTableItem(next))
    ) {
      items.splice(index + 1, 0, { kind: "separator" });
    }
  }
};

type PlannedPart = {
  block: AnalyzedBlock;
  stripLabel: string | undefined;
  /** A slot's lead to drop off the part: a field's name the template
   *  shows itself. */
  dropLead?: { text: string; shownByTemplateBlockId: string };
  /** Whether a line break behind the stripped label goes with it. */
  stripLabelBreak?: boolean;
  labelAsText: boolean;
};

const plainPiece = (kind: "text" | "tab", text: string): ExpectedPiece => ({
  kind,
  text,
  signature: "",
  link: null,
  atom: null,
  origin: "document",
});

/** A label written as text: the label, then what its list puts after it. */
const labelPieces = ({
  label,
  suffix,
}: NonNullable<ParagraphLabel["asText"]>): ExpectedPiece[] => {
  const pieces = [plainPiece("text", label)];
  if (suffix === "tab") {
    pieces.push(plainPiece("tab", "\t"));
  } else if (suffix === "space") {
    pieces.push(plainPiece("text", " "));
  }
  return pieces;
};

type ParagraphRequest = {
  parts: PlannedPart[];
  separators: TemplatePlanMerge[];
  styleId: string | undefined;
  numbering: TemplatePlanNumbering | null | undefined;
  placement: ResolvedParagraph["placement"];
  /** Where a `"new"` list without a group gets its own key. */
  ordinal: number;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  indent?: TemplatePlanIndent | undefined;
  tabStops?: readonly number[] | undefined;
  alignment?: TemplatePlanAlignment | undefined;
  endsSection?: TemplatePlanSection | undefined;
};

const SHA256 = /^[0-9a-f]{64}$/u;
const TRANSITIONAL_WORDPROCESSINGML: string = NAMESPACES.w;

const groupByBlock = <Entry extends { blockId: string }>(
  entries: readonly Entry[],
): Map<string, Entry[]> => {
  const groups = new Map<string, Entry[]>();
  for (const entry of entries) {
    const group = groups.get(entry.blockId);
    if (group === undefined) {
      groups.set(entry.blockId, [entry]);
    } else {
      group.push(entry);
    }
  }
  return groups;
};

/**
 * Check `plan` against the analysed source and template and resolve it.
 * Every refusal the engine documents is decided here, before anything is written.
 */
export const resolvePlan = (
  source: AnalyzedDocument,
  template: AnalyzedDocument,
  plan: TemplatePlan,
): Result<ResolvedPlan, TemplatePlanError> => {
  const invalid = (message: string): Result<never, TemplatePlanError> =>
    Result.err(refuse({ code: "plan-invalid", message }));
  const unsupported = (
    message: string,
    blockId?: string,
  ): Result<never, TemplatePlanError> =>
    Result.err(
      refuse({
        code: "unsupported-content",
        message,
        ...(blockId === undefined ? {} : { blockId }),
      }),
    );
  const shapeIssue = planShapeIssue(plan);
  if (shapeIssue !== null) {
    return invalid(shapeIssue);
  }
  if (
    !SHA256.test(plan.sourceSha256) ||
    plan.sourceSha256 !== source.manifest.sha256
  ) {
    return Result.err(
      refuse({
        code: "sha-mismatch",
        message: "The plan was made for another source document.",
      }),
    );
  }
  if (plan.templateSha256 !== template.manifest.sha256) {
    return Result.err(
      refuse({
        code: "sha-mismatch",
        message: "The plan was made for another template.",
      }),
    );
  }
  // Markup is written in the Transitional spelling the template declares;
  // a Strict package on either side would end up mixed.
  for (const [name, analyzed] of [
    ["source", source],
    ["template", template],
  ] as const) {
    if (analyzed.document.root.namespaceUri !== TRANSITIONAL_WORDPROCESSINGML) {
      return unsupported(
        `The ${name} is not a Transitional WordprocessingML package.`,
      );
    }
  }
  // A source read for a plan that keeps its tracked changes carries them
  // as tagged pieces; any other source has to be free of them.
  if (
    source.manifest.trackedChanges > 0 &&
    source.manifest.trackedChangesKept === undefined
  ) {
    return Result.err(
      refuse({
        code: "tracked-changes",
        message: `The source holds ${source.manifest.trackedChanges} tracked change(s); resolve them first.`,
      }),
    );
  }
  // Text after a field that never closes reads as field code, so nothing
  // read from such a story can be trusted.
  for (const state of [
    source.story.state,
    source.footnotes.state,
    source.endnotes.state,
  ]) {
    if (state.open.length > 0 || state.strayFieldEnds > 0) {
      return unsupported(
        "The source holds complex field marks that do not pair up.",
      );
    }
  }
  const unreadBlock = source.story.unreadBlocks.at(0);
  if (unreadBlock !== undefined) {
    return unsupported(
      `The source holds block-level ${unreadBlock} content that cannot be moved onto a template.`,
    );
  }

  const references = plan.references ?? [];
  const highlights = plan.highlights ?? [];
  const freezeKinds = plan.fields?.freezeKinds ?? [];
  const patches = plan.numberingPatches ?? [];

  const templateNumbering = template.numbering;
  const usedNumIds = (templateNumbering?.definitions.nums ?? []).map(
    ({ numId }) => numId,
  );
  let nextNumId = maxOf(usedNumIds, 0) + 1;
  const mints = new Map<string, NumberingMint>();

  const paragraphStyle = (
    styleId: string | undefined,
    blockId: string,
  ): Result<string | null, TemplatePlanError> => {
    if (styleId === undefined) {
      return Result.ok(null);
    }
    return template.styles.get(styleId)?.type === "paragraph"
      ? Result.ok(styleId)
      : Result.err(
          refuse({
            code: "style-unknown",
            message: `The template has no paragraph style ${styleId}.`,
            blockId,
          }),
        );
  };

  type NumberingRequest = {
    numbering: TemplatePlanNumbering | null | undefined;
    styleId: string | null;
    blockId: string;
    ordinal: number;
  };
  const resolveNumbering = ({
    numbering,
    styleId,
    blockId,
    ordinal,
  }: NumberingRequest): Result<ResolvedNumbering, TemplatePlanError> => {
    if (numbering === undefined) {
      return Result.ok({ kind: "style" });
    }
    if (numbering === null) {
      return Result.ok({ kind: "none" });
    }
    const unknown = (message: string): Result<never, TemplatePlanError> =>
      Result.err(refuse({ code: "numbering-unknown", message, blockId }));
    const { numId, ilvl, restartFrom, baseNumId, group } = numbering;
    if (templateNumbering === null) {
      return unknown("The template has no numbering definitions.");
    }
    const mintsInstance =
      numId === "new" || restartFrom !== undefined || group !== undefined;
    if (!mintsInstance) {
      return computeListRendering({ numId, ilvl }, templateNumbering) === null
        ? unknown(`The template has no list ${numId} with level ${ilvl}.`)
        : Result.ok({ kind: "direct", numId, ilvl });
    }
    let base: number | undefined = numId === "new" ? baseNumId : numId;
    if (base === undefined) {
      const bound = styleNumbering(
        styleId ?? template.defaultStyleId,
        template.styles,
        templateNumbering,
      );
      base = bound?.kind === "reference" ? bound.numId : undefined;
    }
    const abstractNumId =
      base === undefined ? null : templateNumbering.getAbstractNumId(base);
    // The level as the list renders it, from the one reader of list levels.
    const baseLevel =
      base === undefined
        ? null
        : computeListRendering({ numId: base, ilvl }, templateNumbering);
    if (base === undefined || abstractNumId === null || baseLevel === null) {
      return unknown(
        `Block ${blockId} starts a list the template does not define.`,
      );
    }
    // A group's key and an ungrouped item's key never meet.
    const key = group === undefined ? `item:${ordinal}` : `group:${group}`;
    let mint = mints.get(key);
    if (mint === undefined) {
      // A patched start is the start the new instance restarts from.
      const patchedStart = patches.find(
        (patch) => patch.abstractNumId === abstractNumId && patch.ilvl === ilvl,
      )?.start;
      mint = {
        key,
        numId: nextNumId,
        baseNumId: base,
        abstractNumId,
        starts: new Map([
          [
            ilvl,
            restartFrom ?? patchedStart ?? baseLevel.levelStarts?.at(ilvl) ?? 1,
          ],
        ]),
      };
      nextNumId += 1;
      mints.set(key, mint);
      return Result.ok({ kind: "direct", numId: mint.numId, ilvl });
    }
    if (mint.baseNumId !== base) {
      return invalid(`List group ${key} spans two lists.`);
    }
    if (restartFrom !== undefined) {
      const known = mint.starts.get(ilvl);
      if (known !== undefined && known !== restartFrom) {
        return invalid(
          `List group ${key} restarts level ${ilvl} from two different numbers.`,
        );
      }
      mint.starts.set(ilvl, restartFrom);
    }
    return Result.ok({ kind: "direct", numId: mint.numId, ilvl });
  };

  const referencesByBlock = groupByBlock(references);
  const highlightsByBlock = groupByBlock(highlights);
  const byBlockId = new Map<string, ResolvedParagraph>();
  const placed = new Set<string>();

  const place = (
    block: AnalyzedBlock,
  ): Result<undefined, TemplatePlanError> => {
    const { id } = block.block;
    if (placed.has(id)) {
      return Result.err(
        refuse({
          code: "block-placed-twice",
          message: `Source block ${id} is placed more than once.`,
          blockId: id,
        }),
      );
    }
    placed.add(id);
    return Result.ok(undefined);
  };

  const resolveParagraph = ({
    parts,
    separators,
    styleId,
    numbering,
    placement,
    ordinal,
    pageBreakBefore = false,
    keepNext = false,
    indent,
    tabStops,
    alignment,
    endsSection,
  }: ParagraphRequest): Result<ResolvedParagraph, TemplatePlanError> => {
    const first = parts.at(0);
    if (first === undefined) {
      return invalid("A plan item names no source block.");
    }
    const firstId = first.block.block.id;
    const style = paragraphStyle(styleId, firstId);
    if (Result.isError(style)) {
      return style;
    }
    const resolvedNumbering = resolveNumbering({
      numbering,
      styleId: style.value,
      blockId: firstId,
      ordinal,
    });
    if (Result.isError(resolvedNumbering)) {
      return resolvedNumbering;
    }
    const segments: ExpectedSegment[] = [];
    const resolvedParts: ResolvedPart[] = [];
    for (const [index, planned] of parts.entries()) {
      const { block, stripLabel, stripLabelBreak } = planned;
      const { id, text } = block.block;
      if (block.read === null) {
        return invalid(
          `Block ${id} is a table and cannot be joined or slotted.`,
        );
      }
      // A tracked paragraph mark is the mark of a paragraph of its own. A
      // block that ends in an inserted one joins no block after it, wherever
      // it stands in a chain: joined, the mark would stand nowhere. And a
      // slot is one paragraph of the template, which a block that holds a
      // tracked mark of either kind does not fit.
      const endsInInsertedMark = block.read.atoms.some(
        ({ revision }) => revision?.mark === true,
      );
      const holdsDeletedMark = block.read.atoms.some(
        ({ revision }) => revision?.deleted === DELETED_PARAGRAPH_MARK,
      );
      if (
        (endsInInsertedMark && index < parts.length - 1) ||
        (placement === "slot" && (endsInInsertedMark || holdsDeletedMark))
      ) {
        return Result.err(
          refuse({
            code: "merge-invalid",
            message:
              placement === "slot"
                ? `Block ${id} holds a tracked paragraph mark and cannot fill a paragraph of the template.`
                : `Block ${id} ends in a tracked paragraph mark and cannot join the next block.`,
            blockId: id,
          }),
        );
      }
      const placedOnce = place(block);
      if (Result.isError(placedOnce)) {
        return placedOnce;
      }
      const unread = block.read.unsupported.at(0);
      if (unread !== undefined) {
        return unsupported(
          `Block ${id} holds ${unread} content that cannot be moved onto a template.`,
          id,
        );
      }
      let strippedPrefix = "";
      if (stripLabel !== undefined) {
        const lead = LEADING_SPACE.exec(text)?.[0] ?? "";
        const labelEnd = lead.length + stripLabel.length;
        // Only a label goes: a short designator the engine recognises, that
        // the block starts with and that ends where a word would.
        if (
          !isManualLabel(stripLabel) ||
          !text.startsWith(stripLabel, lead.length) ||
          !(labelEnd === text.length || LABEL_END.test(text.charAt(labelEnd)))
        ) {
          return Result.err(
            refuse({
              code: "strip-label-mismatch",
              message: `Block ${id} does not start with the label to strip.`,
              blockId: id,
            }),
          );
        }
        const blank =
          stripLabelBreak === true ? LABEL_SPACE_AND_BREAK : LABEL_SPACE;
        strippedPrefix = text.slice(
          0,
          labelEnd + (blank.exec(text.slice(labelEnd))?.[0].length ?? 0),
        );
      }
      // A slot's lead: the block starts with it, and something follows.
      let droppedLead: ResolvedPart["droppedLead"] = null;
      if (planned.dropLead !== undefined) {
        const lead = LEADING_SPACE.exec(text)?.[0] ?? "";
        const leadEnd = lead.length + planned.dropLead.text.length;
        const dropped = text.slice(
          0,
          leadEnd + (LABEL_SPACE.exec(text.slice(leadEnd))?.[0].length ?? 0),
        );
        if (
          planned.dropLead.text.trim() === "" ||
          !text.startsWith(planned.dropLead.text, lead.length) ||
          text.slice(dropped.length).trim() === ""
        ) {
          return Result.err(
            refuse({
              code: "strip-label-mismatch",
              message: `Block ${id} does not start with the lead to drop, or holds nothing behind it.`,
              blockId: id,
            }),
          );
        }
        droppedLead = {
          text: dropped,
          shownByTemplateBlockId: planned.dropLead.shownByTemplateBlockId,
        };
      }
      const part = partSegments({
        block,
        pieces: paragraphPieces({
          read: block.read,
          context: source,
          origin: "document",
          fields: source.story.state.fields,
        }),
        strippedLength:
          droppedLead === null
            ? strippedPrefix.length
            : droppedLead.text.length,
        strippedBreak: stripLabelBreak === true,
        references: referencesByBlock.get(id) ?? [],
        highlights: highlightsByBlock.get(id) ?? [],
      });
      if (Result.isError(part)) {
        return part;
      }
      if (index > 0) {
        const separator = SEPARATOR_PIECES[separators[index - 1] ?? "none"];
        if (separator !== null) {
          segments.push({ kind: "piece", piece: separator });
        }
      }
      const written = planned.labelAsText
        ? (block.label?.asText ?? null)
        : null;
      if (planned.labelAsText) {
        // The label is the source's own, and only one that is the same
        // text everywhere; the paragraph it leads renders no other.
        if (written === null) {
          return Result.err(
            refuse({
              code: "label-not-writable",
              message: `Block ${id} shows no automatic label that can be written as text.`,
              blockId: id,
            }),
          );
        }
        if (
          stripLabel !== undefined ||
          (index === 0 && resolvedNumbering.value.kind !== "none")
        ) {
          return invalid(
            `Block ${id} takes its label as text and can neither lose a typed label nor join a list.`,
          );
        }
        segments.push(
          ...labelPieces(written).map((piece): ExpectedSegment => ({
            kind: "piece",
            piece,
          })),
        );
      }
      segments.push(...part.value);
      resolvedParts.push({
        block,
        strippedPrefix,
        droppedLead,
        labelAsText: written?.label ?? null,
      });
    }
    const pieces = segmentPieces(segments);
    for (const piece of pieces) {
      const notes = NOTES_OF_REFERENCE[piece.kind];
      if (notes === undefined) {
        continue;
      }
      const noteId = wAttr(piece.atom?.element, "id") ?? "";
      const note = source[notes].byId.get(noteId);
      if (note === undefined) {
        return unsupported(
          `Block ${firstId} refers to note ${noteId}, which the source does not hold.`,
          firstId,
        );
      }
      const unread =
        note.paragraphs.flatMap((read) => read.unsupported).at(0) ??
        (note.hasOtherBlocks ? "table" : undefined);
      if (unread !== undefined) {
        return unsupported(
          `Note ${noteId} holds ${unread} content that cannot be moved onto a template.`,
          firstId,
        );
      }
    }
    const paragraph: ResolvedParagraph = {
      parts: resolvedParts,
      styleId: style.value,
      numbering: resolvedNumbering.value,
      segments,
      ...piecesContent(pieces),
      inlineObjects: pieces.filter(({ kind }) => kind === "object").length,
      placement,
      pageBreakBefore,
      keepNext,
      indent: indent ?? null,
      tabStops: tabStops ?? null,
      alignment: alignment ?? null,
      endsSection: endsSection ?? null,
    };
    for (const { block } of resolvedParts) {
      byBlockId.set(block.block.id, paragraph);
    }
    return Result.ok(paragraph);
  };

  const sourceTopLevel = (
    id: string,
  ): Result<AnalyzedBlock, TemplatePlanError> => {
    const block = source.story.byId.get(id);
    if (block === undefined || block.parentId !== null) {
      return Result.err(
        refuse({
          code: "unknown-block",
          message:
            block === undefined
              ? `The source has no block ${id}.`
              : `Block ${id} sits in a table and is placed with it.`,
          blockId: id,
        }),
      );
    }
    return Result.ok(block);
  };

  const resolveTable = (
    block: AnalyzedBlock,
    item: TemplatePlanSourceItem,
    tableOrdinal: number,
  ): Result<ResolvedItem, TemplatePlanError> => {
    const { id } = block.block;
    if (
      item.mergeNext !== undefined ||
      item.stripLabel !== undefined ||
      item.stripLabelBreak !== undefined ||
      item.labelAsText !== undefined ||
      item.pageBreakBefore !== undefined ||
      item.keepNext !== undefined ||
      item.indent !== undefined ||
      item.tabStops !== undefined ||
      item.alignment !== undefined ||
      item.endsSection !== undefined ||
      item.numbering !== undefined
    ) {
      return invalid(`Table ${id} takes a style and nested placements only.`);
    }
    if (
      item.styleId !== undefined &&
      template.styles.get(item.styleId)?.type !== "table"
    ) {
      return Result.err(
        refuse({
          code: "style-unknown",
          message: `The template has no table style ${item.styleId}.`,
          blockId: id,
        }),
      );
    }
    const placedOnce = place(block);
    if (Result.isError(placedOnce)) {
      return placedOnce;
    }
    const overrides = new Map(
      (item.nested ?? []).map((nested) => [nested.blockId, nested]),
    );
    const cells = new Map<XmlElement, ResolvedParagraph>();
    const nestedBlocks = source.story.order.filter(
      (candidate) => candidate.topLevelId === id && candidate !== block,
    );
    for (const nested of nestedBlocks) {
      if (nested.read === null) {
        const placedTable = place(nested);
        if (Result.isError(placedTable)) {
          return placedTable;
        }
        continue;
      }
      const override = overrides.get(nested.block.id);
      overrides.delete(nested.block.id);
      const paragraph = resolveParagraph({
        parts: [
          {
            block: nested,
            stripLabel: override?.stripLabel,
            labelAsText: override?.labelAsText === true,
          },
        ],
        separators: [],
        styleId: override?.styleId ?? item.cellStyleId,
        numbering: override?.numbering,
        placement: "nested",
        ordinal: tableOrdinal * 1_000_000 + nested.block.index,
        keepNext: override?.keepNext === true,
        indent: override?.indent,
        tabStops: override?.tabStops,
        alignment: override?.alignment,
      });
      if (Result.isError(paragraph)) {
        return paragraph;
      }
      cells.set(nested.element, paragraph.value);
    }
    const stray = [...overrides.keys()].at(0);
    if (stray !== undefined) {
      return Result.err(
        refuse({
          code: "unknown-block",
          message: `Table ${id} holds no paragraph ${stray}.`,
          blockId: stray,
        }),
      );
    }
    return Result.ok({
      kind: "table",
      block,
      styleId: item.styleId ?? null,
      cells,
    });
  };

  const items: ResolvedItem[] = [];
  const nestedSlots: ResolvedSlot[] = [];
  const keptTemplateBlocks: AnalyzedBlock[] = [];
  const keptTemplateIds = new Set<string>();
  const slottedTemplateIds = new Set<string>();
  const pendingNestedSlots: ResolvedSlot[] = [];
  /** The template paragraphs slots name as showing a dropped lead's field. */
  const leadShowers: { templateBlockId: string; slotId: string }[] = [];
  let ordinal = 0;

  for (let index = 0; index < plan.body.length; index += 1) {
    const item = plan.body[index];
    ordinal += 1;
    if (item === undefined) {
      continue;
    }
    switch (item.kind) {
      case "template": {
        for (const id of item.templateBlockIds) {
          const block = template.story.byId.get(id);
          if (block === undefined || block.parentId !== null) {
            return Result.err(
              refuse({
                code: "unknown-block",
                message: `The template has no top-level block ${id}.`,
                blockId: id,
              }),
            );
          }
          if (keptTemplateIds.has(id) || slottedTemplateIds.has(id)) {
            return Result.err(
              refuse({
                code: "block-placed-twice",
                message: `Template block ${id} is placed more than once.`,
                blockId: id,
              }),
            );
          }
          keptTemplateIds.add(id);
          keptTemplateBlocks.push(block);
          items.push({ kind: "template", block });
        }
        break;
      }
      case "source": {
        const block = sourceTopLevel(item.blockId);
        if (Result.isError(block)) {
          return block;
        }
        if (block.value.block.hash !== item.hash) {
          return Result.err(
            refuse({
              code: "hash-mismatch",
              message: `Block ${item.blockId} does not have the hash the plan pins.`,
              blockId: item.blockId,
            }),
          );
        }
        if (block.value.read === null) {
          const table = resolveTable(block.value, item, ordinal);
          if (Result.isError(table)) {
            return table;
          }
          items.push(table.value);
          break;
        }
        const parts: PlannedPart[] = [
          {
            block: block.value,
            stripLabel: item.stripLabel,
            stripLabelBreak: item.stripLabelBreak === true,
            labelAsText: item.labelAsText === true,
          },
        ];
        const separators: TemplatePlanMerge[] = [];
        let merge = item.mergeNext;
        while (merge !== undefined) {
          const next = plan.body[index + 1];
          const nextBlock =
            next?.kind === "source" ? sourceTopLevel(next.blockId) : undefined;
          if (
            next?.kind !== "source" ||
            nextBlock === undefined ||
            Result.isError(nextBlock) ||
            nextBlock.value.read === null
          ) {
            return Result.err(
              refuse({
                code: "merge-invalid",
                message: `Block ${item.blockId} can only join a source paragraph placed right after it.`,
                blockId: item.blockId,
              }),
            );
          }
          if (nextBlock.value.block.hash !== next.hash) {
            return Result.err(
              refuse({
                code: "hash-mismatch",
                message: `Block ${next.blockId} does not have the hash the plan pins.`,
                blockId: next.blockId,
              }),
            );
          }
          // The joined paragraph has one style and one list: the first item's.
          if (
            next.styleId !== undefined ||
            next.numbering !== undefined ||
            next.cellStyleId !== undefined ||
            next.nested !== undefined ||
            next.pageBreakBefore !== undefined ||
            next.keepNext !== undefined ||
            next.indent !== undefined ||
            next.tabStops !== undefined ||
            next.alignment !== undefined ||
            next.endsSection !== undefined
          ) {
            return Result.err(
              refuse({
                code: "merge-invalid",
                message: `Block ${next.blockId} joins the paragraph before it and cannot carry a style, a list, a page start, a format or a section end of its own.`,
                blockId: next.blockId,
              }),
            );
          }
          separators.push(merge);
          parts.push({
            block: nextBlock.value,
            stripLabel: next.stripLabel,
            stripLabelBreak: next.stripLabelBreak === true,
            labelAsText: next.labelAsText === true,
          });
          merge = next.mergeNext;
          index += 1;
        }
        const paragraph = resolveParagraph({
          parts,
          separators,
          styleId: item.styleId,
          numbering: item.numbering,
          placement: "body",
          ordinal,
          pageBreakBefore: item.pageBreakBefore === true,
          keepNext: item.keepNext === true,
          indent: item.indent,
          tabStops: item.tabStops,
          alignment: item.alignment,
          endsSection: item.endsSection,
        });
        if (Result.isError(paragraph)) {
          return paragraph;
        }
        items.push({ kind: "paragraph", paragraph: paragraph.value });
        break;
      }
      case "slot": {
        const templateBlock = template.story.byId.get(item.templateBlockId);
        if (
          templateBlock === undefined ||
          templateBlock.read === null ||
          !Array.isArray(item.sourceBlockIds)
        ) {
          return Result.err(
            refuse({
              code: "unknown-block",
              message: `The template has no paragraph ${item.templateBlockId} to fill.`,
              blockId: item.templateBlockId,
            }),
          );
        }
        if (
          keptTemplateIds.has(item.templateBlockId) ||
          slottedTemplateIds.has(item.templateBlockId)
        ) {
          return Result.err(
            refuse({
              code: "block-placed-twice",
              message: `Template block ${item.templateBlockId} is placed more than once.`,
              blockId: item.templateBlockId,
            }),
          );
        }
        slottedTemplateIds.add(item.templateBlockId);
        if (
          item.dropLead !== undefined &&
          (typeof item.dropLead.text !== "string" ||
            typeof item.dropLead.shownByTemplateBlockId !== "string")
        ) {
          return invalid(
            `Slot ${item.templateBlockId} names a lead to drop without its text or the paragraph that shows it.`,
          );
        }
        const parts: PlannedPart[] = [];
        for (const [position, id] of item.sourceBlockIds.entries()) {
          const block = sourceTopLevel(id);
          if (Result.isError(block)) {
            return block;
          }
          parts.push({
            block: block.value,
            stripLabel: undefined,
            labelAsText: false,
            ...(position === 0 && item.dropLead !== undefined
              ? { dropLead: item.dropLead }
              : {}),
          });
        }
        if (item.dropLead !== undefined) {
          leadShowers.push({
            templateBlockId: item.dropLead.shownByTemplateBlockId,
            slotId: item.templateBlockId,
          });
        }
        const paragraph = resolveParagraph({
          parts,
          separators: parts.slice(1).map(() => "lineBreak"),
          styleId: undefined,
          numbering: undefined,
          placement: "slot",
          ordinal,
        });
        if (Result.isError(paragraph)) {
          return paragraph;
        }
        const slot = { templateBlock, paragraph: paragraph.value };
        if (templateBlock.parentId === null) {
          items.push({ kind: "slot", ...slot });
        } else {
          pendingNestedSlots.push(slot);
        }
        break;
      }
      default:
        return invalid(`Body item ${index} has an unknown kind.`);
    }
  }

  for (const slot of pendingNestedSlots) {
    if (!keptTemplateIds.has(slot.templateBlock.topLevelId)) {
      return Result.err(
        refuse({
          code: "unknown-block",
          message: `Slot ${slot.templateBlock.block.id} sits in a template table the plan does not keep.`,
          blockId: slot.templateBlock.block.id,
        }),
      );
    }
    nestedSlots.push(slot);
  }
  const emptied: AnalyzedBlock[] = [];
  for (const id of plan.emptyTemplateBlockIds ?? []) {
    const block = template.story.byId.get(id);
    if (
      block === undefined ||
      block.read === null ||
      block.parentId === null ||
      !keptTemplateIds.has(block.topLevelId) ||
      slottedTemplateIds.has(id) ||
      emptied.includes(block)
    ) {
      return Result.err(
        refuse({
          code: "unknown-block",
          message: `Template block ${id} is no paragraph of a kept template table that the plan may empty once.`,
          blockId: id,
        }),
      );
    }
    emptied.push(block);
  }

  // The paragraph that shows a dropped lead's field is one the output has
  // as the template has it: kept, with text, and neither filled nor emptied.
  for (const { templateBlockId, slotId } of leadShowers) {
    const shower = template.story.byId.get(templateBlockId);
    if (
      shower === undefined ||
      shower.read === null ||
      shower.block.text.trim() === "" ||
      !keptTemplateIds.has(shower.topLevelId) ||
      slottedTemplateIds.has(templateBlockId) ||
      emptied.includes(shower)
    ) {
      return Result.err(
        refuse({
          code: "unknown-block",
          message: `Template block ${templateBlockId} is no kept paragraph with text that could show the field of slot ${slotId}.`,
          blockId: templateBlockId,
        }),
      );
    }
  }

  partTables(items);
  const splitsSections =
    plan.lastSection !== undefined ||
    items.some(
      (item) =>
        item.kind === "paragraph" && item.paragraph.endsSection !== null,
    );
  if (splitsSections && wChild(template.body, "sectPr") === null) {
    return invalid(
      "The template has no last section to build the plan's sections from.",
    );
  }

  // A template field that runs over several blocks is kept whole, in place
  // and with nothing in between, or not at all: anything placed inside it
  // would become part of the field's result.
  const keptPosition = new Map<string, number>();
  for (const [position, item] of items.entries()) {
    if (item.kind === "template") {
      keptPosition.set(item.block.block.id, position);
    }
  }
  const templateOrder = template.story.topLevel.map(({ block }) => block.id);
  const slots = [
    ...nestedSlots,
    ...items.flatMap((item) => (item.kind === "slot" ? [item] : [])),
    ...emptied.map((templateBlock) => ({ templateBlock })),
  ];
  for (const field of template.story.state.fields) {
    const lastOrdinal = field.closed
      ? field.endParagraphOrdinal
      : Number.POSITIVE_INFINITY;
    if (lastOrdinal === field.paragraphOrdinal) {
      continue;
    }
    const slotted = slots.find(({ templateBlock }) => {
      const slotOrdinal = templateBlock.read?.ordinal ?? -1;
      return (
        slotOrdinal >= field.paragraphOrdinal && slotOrdinal <= lastOrdinal
      );
    });
    if (slotted !== undefined) {
      return Result.err(
        refuse({
          code: "plan-invalid",
          message: `Slot ${slotted.templateBlock.block.id} lies inside a ${field.kind} field that runs over several paragraphs.`,
          blockId: slotted.templateBlock.block.id,
        }),
      );
    }
    const from = templateOrder.indexOf(
      template.story.paragraphs[field.paragraphOrdinal]?.topLevelId ?? "",
    );
    const to = field.closed
      ? templateOrder.indexOf(
          template.story.paragraphs[field.endParagraphOrdinal]?.topLevelId ??
            "",
        )
      : templateOrder.length - 1;
    if (from === -1 || to <= from) {
      continue;
    }
    const span = templateOrder.slice(from, to + 1);
    if (!span.some((id) => keptPosition.has(id))) {
      continue;
    }
    const first = keptPosition.get(span.at(0) ?? "");
    if (
      first === undefined ||
      span.some((id, offset) => keptPosition.get(id) !== first + offset)
    ) {
      return Result.err(
        refuse({
          code: "plan-invalid",
          message: `Template blocks ${span.at(0) ?? ""} to ${span.at(-1) ?? ""} hold one ${field.kind} field; keep all of them together and in order, or none.`,
          blockId: span.at(0) ?? "",
        }),
      );
    }
  }

  // A kept table of contents keeps no entries: they are a cached result
  // that names the headings and pages of the template's sample text. The
  // field stays in the paragraph it begins in; the further paragraphs of
  // its entries are not placed.
  // A plan that freezes such a field keeps its result as text instead, so
  // nothing of it is a field to empty.
  const emptiedContents = new Map<AnalyzedBlock, number>();
  for (const { first, textKept, rest } of freezeKinds.includes("TOC")
    ? []
    : contentsWithEntries(template, keptTemplateIds)) {
    emptiedContents.set(first, textKept);
    const gone = new Set(rest);
    for (let position = items.length - 1; position >= 0; position -= 1) {
      const item = items[position];
      if (item?.kind === "template" && gone.has(item.block)) {
        items.splice(position, 1);
      }
    }
    for (
      let position = keptTemplateBlocks.length - 1;
      position >= 0;
      position -= 1
    ) {
      const block = keptTemplateBlocks[position];
      if (block !== undefined && gone.has(block)) {
        keptTemplateBlocks.splice(position, 1);
      }
    }
  }

  const omitted = new Map<string, TemplatePlanOmitReason>();
  for (const entry of plan.omit) {
    const block = sourceTopLevel(entry.blockId);
    if (Result.isError(block)) {
      return block;
    }
    if (placed.has(entry.blockId) || omitted.has(entry.blockId)) {
      return Result.err(
        refuse({
          code: "block-placed-twice",
          message: `Block ${entry.blockId} is both placed and omitted, or omitted twice.`,
          blockId: entry.blockId,
        }),
      );
    }
    if (entry.reason === "empty" && !isBlankBlock(block.value)) {
      return invalid(`Block ${entry.blockId} is omitted as empty and is not.`);
    }
    omitted.set(entry.blockId, entry.reason);
  }
  for (const block of source.story.topLevel) {
    const { id } = block.block;
    if (placed.has(id) || omitted.has(id)) {
      continue;
    }
    // Only a block that shows nothing at all may go unmentioned.
    if (!isBlankBlock(block) || showsLabel(block)) {
      return Result.err(
        refuse({
          code: "block-unplaced",
          message: `Source block ${id} is neither placed nor omitted.`,
          blockId: id,
        }),
      );
    }
    omitted.set(id, "empty");
  }

  for (const [blockId] of [...referencesByBlock, ...highlightsByBlock]) {
    if (!byBlockId.has(blockId)) {
      return Result.err(
        refuse({
          code: referencesByBlock.has(blockId)
            ? "reference-span-invalid"
            : "highlight-span-invalid",
          message: `Block ${blockId} is not placed, so nothing can be marked in it.`,
          blockId,
        }),
      );
    }
  }

  const takenNames = new Set<string>();
  const takenIds = bookmarkIdsOf(template.document.root);
  walkElements(template.document.root, (element) => {
    if (isW(element, "bookmarkStart")) {
      takenNames.add(wAttr(element, "name") ?? "");
    }
  });
  let nextBookmarkId = maxOf(takenIds, 0) + 1;
  const bookmarks = new Map<ResolvedParagraph, BookmarkWrap>();
  const resolvedReferences: ResolvedPlan["references"] = [];
  for (const reference of references) {
    const paragraph = byBlockId.get(reference.blockId);
    const target = byBlockId.get(reference.targetBlockId);
    if (paragraph === undefined) {
      return invalid(`Block ${reference.blockId} is not placed.`);
    }
    if (target === undefined) {
      return Result.err(
        refuse({
          code: "reference-target-unplaced",
          message: `Reference target ${reference.targetBlockId} is not a placed paragraph.`,
          blockId: reference.blockId,
        }),
      );
    }
    if (paragraph === target) {
      // A field inside its own bookmark refers to itself.
      return Result.err(
        refuse({
          code: "reference-span-invalid",
          message: `A reference in ${reference.blockId} targets the paragraph it sits in.`,
          blockId: reference.blockId,
        }),
      );
    }
    let bookmark = bookmarks.get(target);
    if (bookmark === undefined) {
      bookmark = {
        id: nextBookmarkId,
        name: mintRefBookmarkName(
          `${source.manifest.sha256}:${target.parts.at(0)?.block.block.id ?? ""}`,
          takenNames,
        ),
      };
      nextBookmarkId += 1;
      bookmarks.set(target, bookmark);
    }
    const instruction = refFieldInstruction(
      bookmark.name,
      reference.instruction,
    );
    if (instruction === null) {
      return invalid(
        `A reference in ${reference.blockId} carries more than field switches.`,
      );
    }
    const segment = paragraph.segments.find(
      (candidate) =>
        candidate.kind === "reference" && candidate.reference === reference,
    );
    resolvedReferences.push({
      reference,
      paragraph,
      target,
      replacedText: segment?.kind === "reference" ? segment.replacedText : "",
      instruction,
    });
  }

  for (const patch of patches) {
    const abstract = templateNumbering?.getAbstract(patch.abstractNumId);
    if (
      abstract === null ||
      abstract === undefined ||
      !abstract.levels.some(({ ilvl }) => ilvl === patch.ilvl)
    ) {
      return Result.err(
        refuse({
          code: "numbering-unknown",
          message: `The template has no level ${patch.ilvl} in list definition ${patch.abstractNumId}.`,
        }),
      );
    }
  }

  return Result.ok({
    items,
    nestedSlots,
    emptied,
    byBlockId,
    omitted,
    mints: [...mints.values()],
    bookmarks,
    references: resolvedReferences,
    freezeKinds: new Set(freezeKinds),
    keptTemplateBlocks,
    emptiedContents,
  });
};
