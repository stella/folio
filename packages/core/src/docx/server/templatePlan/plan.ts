/**
 * A template plan checked against its source and template and resolved into
 * what the output must hold: per output paragraph the expected text and
 * emphasis, derived from the source blocks alone. The writer emits it and the
 * verifier compares the written package against it.
 */
import { Result } from "better-result";

import { FOLIO_HIGHLIGHT_COLOR_VALUES } from "../../../document-operations";
import { parseFieldInstruction } from "../../fieldParser";
import type { XmlElement } from "../../xmlParser";
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
import { styleNumbering } from "./labels";
import type {
  AtomKind,
  AtomLink,
  InlineAtom,
  ParagraphRead,
  StoryField,
} from "./story";
import {
  type TemplatePlan,
  TemplatePlanError,
  type TemplatePlanErrorCode,
  type TemplatePlanHighlight,
  type TemplatePlanMerge,
  type TemplatePlanNumbering,
  type TemplatePlanOmitReason,
  type TemplatePlanReference,
  type TemplatePlanSourceItem,
} from "./types";
import { isW, maxOf, walkElements, wAttr, wChildVal } from "./xml";

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

export type PieceOrigin = "document" | "footnotes" | "endnotes";

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
  abstractNumId: number;
  ilvl: number;
  start: number;
};

export type ResolvedNumbering =
  | { kind: "style" }
  | { kind: "none" }
  | { kind: "direct"; numId: number; ilvl: number };

export type ResolvedPart = {
  block: AnalyzedBlock;
  strippedPrefix: string;
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
};

export type ResolvedItem =
  | { kind: "template"; block: AnalyzedBlock }
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

export type ResolvedPlan = {
  items: ResolvedItem[];
  /** Slots that sit inside a kept template table. */
  nestedSlots: ResolvedSlot[];
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

const OMIT_REASONS: ReadonlySet<string> = new Set([
  "empty",
  "sourceToc",
  "pageArtifact",
  "duplicateLabel",
]);
const MERGES: ReadonlySet<string> = new Set(["lineBreak", "space", "none"]);
const SUFFIXES: ReadonlySet<string> = new Set(["tab", "space", "nothing"]);
const HIGHLIGHT_COLORS: ReadonlySet<string> = new Set(
  FOLIO_HIGHLIGHT_COLOR_VALUES,
);
const LABEL_SPACE = /^[ \t ]+/u;
const LEADING_SPACE = /^[ \t ]*/u;

/** Whether a block carries nothing a reader would miss. */
export const isEmptyBlock = ({ block }: AnalyzedBlock): boolean =>
  block.text.trim() === "" && block.inlineObjects === 0;

export type SignatureContext = Pick<
  AnalyzedDocument,
  "styles" | "defaultStyleId"
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

/** The external address a HYPERLINK field names, when it names one. */
const hyperlinkFieldUrl = ({
  instruction,
  kind,
}: StoryField): string | null => {
  if (kind !== HYPERLINK_FIELD) {
    return null;
  }
  const { argument, switches } = parseFieldInstruction(instruction);
  return argument === undefined ||
    argument === "" ||
    switches.some(({ switch: name }) => name === "l")
    ? null
    : argument;
};

const NOTE_REFERENCE_KINDS: ReadonlySet<AtomKind> = new Set([
  "footnoteReference",
  "endnoteReference",
]);

/** The pieces of one source paragraph, each with the emphasis it states. */
export const paragraphPieces = ({
  read,
  context: { styles, defaultStyleId },
  origin,
  fields,
}: ParagraphPiecesOptions): ExpectedPiece[] => {
  const baseline = paragraphBaseline(
    styles.get(read.styleId ?? defaultStyleId ?? ""),
  );
  return read.atoms.map((atom) => {
    // A HYPERLINK field is a link like any other once its result is kept.
    const fieldUrl = atom.fields
      .map((index) => fields[index])
      .map((field) => (field === undefined ? null : hyperlinkFieldUrl(field)))
      .find((url) => url !== null);
    const link: AtomLink | null =
      atom.link ??
      (fieldUrl === undefined || fieldUrl === null
        ? null
        : { kind: "url", url: fieldUrl });
    let signature = "";
    if (!NOTE_REFERENCE_KINDS.has(atom.kind)) {
      const characterStyleId = wChildVal(atom.runProperties, "rStyle");
      // Inside a link the character style is the link's look, which the
      // template supplies; what the run states itself still counts.
      const styled =
        link === null && characterStyleId !== null
          ? withCharacterStyle(baseline, styles.get(characterStyleId))
          : baseline;
      signature = emphasisSignature(
        emphasisDeviation(
          withOverride(styled, readDirectEmphasis(atom.runProperties)),
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

/** Split text pieces so every offset in `cuts` falls on a piece boundary. */
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
        ? cuts.filter((cut) => cut > start && cut < end)
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

type PartSegmentsOptions = {
  block: AnalyzedBlock;
  pieces: ExpectedPiece[];
  strippedLength: number;
  references: readonly TemplatePlanReference[];
  highlights: readonly TemplatePlanHighlight[];
};

const partSegments = ({
  block,
  pieces,
  strippedLength,
  references,
  highlights,
}: PartSegmentsOptions): Result<ExpectedSegment[], TemplatePlanError> => {
  const blockId = block.block.id;
  const length = block.block.text.length;
  for (const { start, end } of references) {
    if (
      !Number.isInteger(start) ||
      !Number.isInteger(end) ||
      start < strippedLength ||
      start >= end ||
      end > length
    ) {
      return Result.err(
        refuse({
          code: "reference-span-invalid",
          message: `Reference span ${start}-${end} lies outside block ${blockId}.`,
          blockId,
        }),
      );
    }
  }
  for (const { start, end, color } of highlights) {
    if (
      !Number.isInteger(start) ||
      !Number.isInteger(end) ||
      start < strippedLength ||
      start >= end ||
      end > length ||
      !HIGHLIGHT_COLORS.has(color)
    ) {
      return Result.err(
        refuse({
          code: "highlight-span-invalid",
          message: `Highlight ${start}-${end} (${color}) does not fit block ${blockId}.`,
          blockId,
        }),
      );
    }
  }
  const ordered = references.toSorted((a, b) => a.start - b.start);
  for (let index = 1; index < ordered.length; index += 1) {
    const previous = ordered[index - 1];
    const current = ordered[index];
    if (
      previous !== undefined &&
      current !== undefined &&
      current.start < previous.end
    ) {
      return Result.err(
        refuse({
          code: "reference-span-invalid",
          message: `Reference spans overlap in block ${blockId}.`,
          blockId,
        }),
      );
    }
  }

  const positioned: PositionedPiece[] = [];
  for (const entry of splitPieces(pieces, [
    strippedLength,
    ...references.flatMap(({ start, end }) => [start, end]),
    ...highlights.flatMap(({ start, end }) => [start, end]),
  ])) {
    // What the label strip removes; a piece without text stays where it was.
    if (entry.end <= strippedLength && entry.start !== entry.end) {
      continue;
    }
    const highlight = highlights.find(
      ({ start, end }) =>
        entry.start >= start && entry.end <= end && entry.end > entry.start,
    );
    if (highlight !== undefined) {
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
  const close = (): void => {
    if (open === undefined) {
      return;
    }
    const replacedText = open.pieces.map(({ text }) => text).join("");
    const first = open.pieces.at(0);
    const { reference } = open;
    segments.push({
      kind: "reference",
      reference,
      replacedText,
      pieces:
        replacedText === reference.expectedText || first === undefined
          ? open.pieces.map((piece) => ({ ...piece, link: null }))
          : [
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
    open = undefined;
  };
  for (const { piece, start, end } of positioned) {
    const empty = start === end;
    const reference = ordered.find((candidate) =>
      empty
        ? start > candidate.start && start < candidate.end
        : start >= candidate.start && end <= candidate.end,
    );
    if (reference === undefined) {
      close();
      segments.push({ kind: "piece", piece });
      continue;
    }
    if (piece.kind !== "text") {
      return Result.err(
        refuse({
          code: "reference-span-invalid",
          message: `Reference span ${reference.start}-${reference.end} in block ${blockId} covers more than text.`,
          blockId,
        }),
      );
    }
    if (empty) {
      continue;
    }
    if (open !== undefined && open.reference !== reference) {
      close();
    }
    open ??= { reference, pieces: [] };
    open.pieces.push(piece);
  }
  close();
  return Result.ok(segments);
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

type PlannedPart = {
  block: AnalyzedBlock;
  stripLabel: string | undefined;
};

type ParagraphRequest = {
  parts: PlannedPart[];
  separators: TemplatePlanMerge[];
  styleId: string | undefined;
  numbering: TemplatePlanNumbering | null | undefined;
  placement: ResolvedParagraph["placement"];
  /** Where a `"new"` list without a group gets its own key. */
  ordinal: number;
};

const SHA256 = /^[0-9a-f]{64}$/u;

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
  // A plan arrives as parsed JSON; its shape is checked where a wrong one
  // would otherwise fail somewhere unrelated.
  const raw: Partial<TemplatePlan> | null | undefined = plan;
  if (
    raw === null ||
    raw === undefined ||
    raw.version !== 1 ||
    !Array.isArray(raw.body) ||
    !Array.isArray(raw.omit) ||
    typeof raw.sourceSha256 !== "string" ||
    typeof raw.templateSha256 !== "string"
  ) {
    return invalid("The plan is not a version 1 template plan.");
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
  if (source.manifest.trackedChanges > 0) {
    return Result.err(
      refuse({
        code: "tracked-changes",
        message: `The source holds ${source.manifest.trackedChanges} tracked change(s); resolve them first.`,
      }),
    );
  }

  const references = plan.references ?? [];
  const highlights = plan.highlights ?? [];
  const freezeKinds = plan.fields?.freezeKinds ?? [];
  if (
    !Array.isArray(references) ||
    !Array.isArray(highlights) ||
    !Array.isArray(freezeKinds) ||
    freezeKinds.some((kind) => !/^[A-Z][A-Z0-9]*$/u.test(kind))
  ) {
    return invalid(
      "The plan's references, highlights or field kinds are malformed.",
    );
  }

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
    if (
      !Number.isInteger(ilvl) ||
      ilvl < 0 ||
      (restartFrom !== undefined && !Number.isInteger(restartFrom))
    ) {
      return invalid(`Block ${blockId} names a malformed list level.`);
    }
    if (templateNumbering === null) {
      return unknown("The template has no numbering definitions.");
    }
    const mintsInstance =
      numId === "new" || restartFrom !== undefined || group !== undefined;
    if (!mintsInstance) {
      return templateNumbering.getLevel(numId, ilvl) === null
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
    if (
      base === undefined ||
      abstractNumId === null ||
      templateNumbering.getLevel(base, ilvl) === null
    ) {
      return unknown(
        `Block ${blockId} starts a list the template does not define.`,
      );
    }
    const key = group ?? `#${ordinal}`;
    let mint = mints.get(key);
    if (mint === undefined) {
      mint = {
        key,
        numId: nextNumId,
        abstractNumId,
        ilvl,
        start:
          restartFrom ?? templateNumbering.getLevel(base, ilvl)?.start ?? 1,
      };
      nextNumId += 1;
      mints.set(key, mint);
    } else if (mint.abstractNumId !== abstractNumId) {
      return invalid(`List group ${key} spans two list definitions.`);
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
    for (const [index, { block, stripLabel }] of parts.entries()) {
      const { id, text } = block.block;
      if (block.read === null) {
        return invalid(
          `Block ${id} is a table and cannot be joined or slotted.`,
        );
      }
      const placedOnce = place(block);
      if (Result.isError(placedOnce)) {
        return placedOnce;
      }
      let strippedPrefix = "";
      if (stripLabel !== undefined) {
        const lead = LEADING_SPACE.exec(text)?.[0] ?? "";
        if (stripLabel === "" || !text.startsWith(stripLabel, lead.length)) {
          return Result.err(
            refuse({
              code: "strip-label-mismatch",
              message: `Block ${id} does not start with the label to strip.`,
              blockId: id,
            }),
          );
        }
        const labelEnd = lead.length + stripLabel.length;
        strippedPrefix = text.slice(
          0,
          labelEnd + (LABEL_SPACE.exec(text.slice(labelEnd))?.[0].length ?? 0),
        );
      }
      const part = partSegments({
        block,
        pieces: paragraphPieces({
          read: block.read,
          context: source,
          origin: "document",
          fields: source.story.state.fields,
        }),
        strippedLength: strippedPrefix.length,
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
      segments.push(...part.value);
      resolvedParts.push({ block, strippedPrefix });
    }
    const pieces = segmentPieces(segments);
    const paragraph: ResolvedParagraph = {
      parts: resolvedParts,
      styleId: style.value,
      numbering: resolvedNumbering.value,
      segments,
      ...piecesContent(pieces),
      inlineObjects: pieces.filter(({ kind }) => kind === "object").length,
      placement,
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
        parts: [{ block: nested, stripLabel: override?.stripLabel }],
        separators: [],
        styleId: override?.styleId ?? item.cellStyleId,
        numbering: override?.numbering,
        placement: "nested",
        ordinal: tableOrdinal * 1_000_000 + nested.block.index,
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
  let ordinal = 0;

  for (let index = 0; index < plan.body.length; index += 1) {
    const item = plan.body[index];
    ordinal += 1;
    if (item === undefined || typeof item !== "object") {
      return invalid(`Body item ${index} is malformed.`);
    }
    switch (item.kind) {
      case "template": {
        if (!Array.isArray(item.templateBlockIds)) {
          return invalid(`Body item ${index} names no template blocks.`);
        }
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
          { block: block.value, stripLabel: item.stripLabel },
        ];
        const separators: TemplatePlanMerge[] = [];
        let merge = item.mergeNext;
        while (merge !== undefined) {
          if (!MERGES.has(merge)) {
            return invalid(`Body item ${index} names an unknown merge.`);
          }
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
          separators.push(merge);
          parts.push({ block: nextBlock.value, stripLabel: next.stripLabel });
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
        const parts: PlannedPart[] = [];
        for (const id of item.sourceBlockIds) {
          const block = sourceTopLevel(id);
          if (Result.isError(block)) {
            return block;
          }
          parts.push({ block: block.value, stripLabel: undefined });
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

  const omitted = new Map<string, TemplatePlanOmitReason>();
  for (const entry of plan.omit) {
    const block = sourceTopLevel(entry.blockId);
    if (Result.isError(block)) {
      return block;
    }
    if (!OMIT_REASONS.has(entry.reason)) {
      return invalid(
        `Block ${entry.blockId} is omitted for an unknown reason.`,
      );
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
    if (entry.reason === "empty" && !isEmptyBlock(block.value)) {
      return invalid(`Block ${entry.blockId} is omitted as empty and is not.`);
    }
    omitted.set(entry.blockId, entry.reason);
  }
  for (const block of source.story.topLevel) {
    const { id } = block.block;
    if (placed.has(id) || omitted.has(id)) {
      continue;
    }
    if (!isEmptyBlock(block)) {
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
    if (
      typeof reference.expectedText !== "string" ||
      reference.expectedText === "" ||
      typeof reference.instruction !== "string"
    ) {
      return invalid(`A reference in ${reference.blockId} has no cached text.`);
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

  for (const patch of plan.numberingPatches ?? []) {
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
    if (
      (patch.numFmt !== undefined && !/^[A-Za-z]+$/u.test(patch.numFmt)) ||
      (patch.suff !== undefined && !SUFFIXES.has(patch.suff)) ||
      (patch.start !== undefined && !Number.isInteger(patch.start)) ||
      (patch.lvlText !== undefined && typeof patch.lvlText !== "string")
    ) {
      return invalid("A numbering patch carries a malformed value.");
    }
  }

  return Result.ok({
    items,
    nestedSlots,
    byBlockId,
    omitted,
    mints: [...mints.values()],
    bookmarks,
    references: resolvedReferences,
    freezeKinds: new Set(freezeKinds),
    keptTemplateBlocks,
  });
};
