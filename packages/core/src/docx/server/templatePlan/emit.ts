/**
 * The markup of what a resolved plan places: paragraphs, tables and notes
 * rebuilt from source content in the template's own spelling. Text, tabs,
 * breaks, note references, links and stated emphasis travel; fonts, sizes,
 * colours and direct paragraph formatting stay behind.
 */
import { panic, Result } from "better-result";

import { resolveNamespaceUri } from "@stll/docx-core/xml/namespaces";

import type { RelationshipMap } from "../../../types/document";
import { parseRelationships, RELATIONSHIP_TYPES } from "../../relsParser";
import { captureVerbatimXml } from "../../verbatimCapture";
import {
  cloneElement,
  getLocalName,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "../../xmlParser";
import { appendToXmlRoot } from "../../xmlRootInsertion";
import type { AnalyzedDocument, AnalyzedNote } from "./blocks";
import {
  bookmarkEndXml,
  bookmarkStartXml,
  type BookmarkWrap,
  complexFieldXml,
  wrapParagraphContentInBookmark,
} from "./builders";
import { runPropertiesXml } from "./emphasis";
import {
  type ExpectedPiece,
  type ExpectedSegment,
  paragraphPieces,
  type PieceOrigin,
  type ResolvedParagraph,
  type SignatureContext,
} from "./plan";
import {
  DELETED_PARAGRAPH_MARK,
  deletedRowMarkOf,
  insertedRowMarkOf,
  keptDeletedRowOf,
  rowTextOf,
} from "./revisions";
import type { AtomLink, AtomRevision, StoryField } from "./story";
import { TemplatePlanError } from "./types";
import {
  elementChildren,
  escapeXmlAttribute,
  escapeXmlText,
  MATH_NAMESPACE,
  type PartPrefixes,
  wAttr,
  wChild,
  wChildren,
  wName,
} from "./xml";

const PACKAGE_RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";

export type RelationshipsEditor = {
  path: string;
  add: (relationship: {
    type: string;
    target: string;
    external: boolean;
  }) => string;
  /** The part's markup with the added relationships, or null when none were added. */
  xml: () => string | null;
};

/**
 * An editor over a relationships part that may not exist yet. Null when the
 * part exists and its root is not a relationships root.
 */
export const createRelationshipsEditor = (
  path: string,
  existing: string | null,
): RelationshipsEditor | null => {
  const source =
    existing ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"></Relationships>`;
  const known = parseRelationships(source);
  // Ids are strings: every one in use is reserved as written, and the
  // counter only follows suffixes it can count on exactly.
  const reserved = new Set(known.keys());
  let next = 1;
  for (const id of reserved) {
    const number = /^rId(?<number>\d{1,15})$/u.exec(id)?.groups?.["number"];
    if (number !== undefined) {
      next = Math.max(next, Number(number) + 1);
    }
  }
  const added = new Map<string, string>();
  const markup: string[] = [];
  const root = parseXmlDocument(source);
  if (
    root?.name === undefined ||
    getLocalName(root.name) !== "Relationships" ||
    root.namespaceUri !== PACKAGE_RELATIONSHIPS_NAMESPACE
  ) {
    return null;
  }
  const rootName = root.name;
  const colon = rootName.indexOf(":");
  const elementName =
    colon === -1 ? "Relationship" : `${rootName.slice(0, colon)}:Relationship`;
  return {
    path,
    add: ({ type, target, external }) => {
      const key = `${type}\u0000${target}\u0000${external ? "1" : "0"}`;
      const reused = added.get(key);
      if (reused !== undefined) {
        return reused;
      }
      let id = `rId${next}`;
      while (reserved.has(id)) {
        next += 1;
        id = `rId${next}`;
      }
      next += 1;
      reserved.add(id);
      added.set(key, id);
      markup.push(
        `<${elementName} Id="${id}" Type="${escapeXmlAttribute(type)}" Target="${escapeXmlAttribute(target)}"${external ? ' TargetMode="External"' : ""}/>`,
      );
      return id;
    },
    xml: () =>
      markup.length === 0
        ? null
        : appendToXmlRoot(source, { rootName, markup: markup.join("") }),
  };
};

export type TemplateCharacterStyles = {
  hyperlink: string | null;
  footnoteReference: string | null;
  endnoteReference: string | null;
  footnoteText: string | null;
  endnoteText: string | null;
};

const styleNamed = (
  { styles }: Pick<AnalyzedDocument, "styles">,
  type: "paragraph" | "character",
  name: string,
): string | null =>
  [...styles.values()].find(
    (style) => style.type === type && style.name?.toLowerCase() === name,
  )?.styleId ?? null;

/** The template styles conventional marks and note text are written with. */
export const templateCharacterStyles = (
  template: Pick<AnalyzedDocument, "styles">,
): TemplateCharacterStyles => ({
  hyperlink: styleNamed(template, "character", "hyperlink"),
  footnoteReference: styleNamed(template, "character", "footnote reference"),
  endnoteReference: styleNamed(template, "character", "endnote reference"),
  footnoteText: styleNamed(template, "paragraph", "footnote text"),
  endnoteText: styleNamed(template, "paragraph", "endnote text"),
});

export type NoteKind = "footnote" | "endnote";

/** A media part of the source, read ahead so the emit stays synchronous. */
export type SourceMedia = { bytes: Uint8Array; extension: string };

export type EmitShared = {
  styles: TemplateCharacterStyles;
  /** The source relationships of each part pieces come from. */
  sourceRelationships: Record<PieceOrigin, RelationshipMap>;
  /** Source media by the package path it is stored at. */
  sourceMedia: ReadonlyMap<string, SourceMedia>;
  resolveSourceTarget: (origin: PieceOrigin, target: string) => string;
  /** The new id of a source note; allocates one the first time a note is met. */
  noteId: (kind: NoteKind, sourceId: string) => number;
  /** Register a media part of the output and return its absolute target. */
  addMedia: (sourcePath: string, media: SourceMedia) => string;
  nextDrawingId: () => number;
  /** A fresh id for a tracked change the output carries. */
  nextRevisionId: () => number;
};

export type EmitContext = EmitShared & {
  prefixes: PartPrefixes;
  relationships: RelationshipsEditor;
};

export type LinkAddress = { url: string; anchor: string | null };

/**
 * The external address a link leads to, or null for a link the output cannot
 * carry: one to a bookmark of the source, or one whose relationship is
 * missing or does not name an external address.
 */
export const linkAddress = (
  link: AtomLink | null,
  relationships: RelationshipMap,
): LinkAddress | null => {
  if (link === null) {
    return null;
  }
  switch (link.kind) {
    case "url":
      return { url: link.url, anchor: link.anchor };
    case "relationship": {
      const relationship = relationships.get(link.relationshipId);
      return relationship === undefined ||
        relationship.type !== RELATIONSHIP_TYPES.hyperlink ||
        relationship.targetMode !== "External"
        ? null
        : { url: relationship.target, anchor: link.anchor };
    }
    case "anchor":
      // The bookmark it points at stays in the source.
      return null;
    default:
      return link satisfies never;
  }
};

/** One string per address, for comparing what a link leads to. */
export const linkAddressKey = (address: LinkAddress | null): string =>
  address === null
    ? ""
    : `${address.url}${address.anchor === null ? "" : `#${address.anchor}`}`;

type LinkTarget = { relationshipId: string; anchor: string | null } | null;

const linkTargetOf = (
  { link, origin }: ExpectedPiece,
  context: EmitContext,
): LinkTarget => {
  const address = linkAddress(link, context.sourceRelationships[origin]);
  return address === null
    ? null
    : {
        relationshipId: context.relationships.add({
          type: RELATIONSHIP_TYPES.hyperlink,
          target: address.url,
          external: true,
        }),
        anchor: address.anchor,
      };
};

const sameLink = (left: AtomLink | null, right: AtomLink | null): boolean => {
  if (left === null || right === null) {
    return left === right;
  }
  if (left.kind === "relationship" && right.kind === "relationship") {
    return (
      left.relationshipId === right.relationshipId &&
      left.anchor === right.anchor
    );
  }
  if (left.kind === "url" && right.kind === "url") {
    return left.url === right.url && left.anchor === right.anchor;
  }
  return (
    left.kind === "anchor" &&
    right.kind === "anchor" &&
    left.anchor === right.anchor
  );
};

type ObjectTransfer = Result<string, TemplatePlanError>;

/** An object of the source with its relationships and drawing ids re-homed. */
const transferObject = (
  { atom, origin }: ExpectedPiece,
  context: EmitContext,
): ObjectTransfer => {
  const element = atom?.element;
  if (element === undefined || element === null) {
    return Result.ok("");
  }
  let refusal: TemplatePlanError | null = null;
  const mapRelationship = (id: string): string => {
    const relationship = context.sourceRelationships[origin].get(id);
    if (relationship === undefined) {
      refusal ??= new TemplatePlanError({
        code: "unsupported-content",
        message: `An object refers to the missing relationship ${id}.`,
      });
      return id;
    }
    const external = relationship.targetMode === "External";
    if (relationship.type === RELATIONSHIP_TYPES.hyperlink && !external) {
      // A link to a part of the source would lead nowhere in the output.
      refusal ??= new TemplatePlanError({
        code: "unsupported-content",
        message: "An object links to a part of the source package.",
      });
      return id;
    }
    if (external) {
      return context.relationships.add({
        type: relationship.type,
        target: relationship.target,
        external,
      });
    }
    const path = context.resolveSourceTarget(origin, relationship.target);
    const media = context.sourceMedia.get(path);
    if (relationship.type !== RELATIONSHIP_TYPES.image || media === undefined) {
      refusal ??= new TemplatePlanError({
        code: "unsupported-content",
        message:
          "An embedded object other than a picture cannot be moved onto a template.",
      });
      return id;
    }
    return context.relationships.add({
      type: RELATIONSHIP_TYPES.image,
      target: context.addMedia(path, media),
      external: false,
    });
  };
  const rehome = (node: XmlElement): XmlElement => {
    if (node.type !== "element") {
      return node;
    }
    const attributes: Record<string, string | number | undefined> = {
      ...node.attributes,
    };
    for (const [name, value] of Object.entries(node.attributes ?? {})) {
      const colon = name.indexOf(":");
      if (value === undefined || colon === -1 || name.startsWith("xmlns:")) {
        continue;
      }
      const uri = resolveNamespaceUri(
        node.namespaceScope,
        name.slice(0, colon),
      );
      if (OFFICE_RELATIONSHIP_NAMESPACE_URIS.has(uri ?? "")) {
        attributes[name] = mapRelationship(String(value));
      }
    }
    if (getLocalName(node.name) === "docPr" && attributes["id"] !== undefined) {
      attributes["id"] = String(context.nextDrawingId());
    }
    return cloneElement(node, {
      attributes,
      ...(node.elements === undefined
        ? {}
        : { elements: node.elements.map(rehome) }),
    });
  };
  const xml = captureVerbatimXml(rehome(element));
  return refusal === null ? Result.ok(xml) : Result.err(refusal);
};

const wAttributesXml = (element: XmlElement, prefix: string): string =>
  Object.entries(element.attributes ?? {})
    .map(([name, value]) => {
      const colon = name.indexOf(":");
      if (value === undefined || colon === -1 || name.startsWith("xmlns:")) {
        return "";
      }
      const uri = resolveNamespaceUri(
        element.namespaceScope,
        name.slice(0, colon),
      );
      return WORDPROCESSINGML_NAMESPACE_URIS.has(uri ?? "")
        ? ` ${prefix}:${name.slice(colon + 1)}="${escapeXmlAttribute(String(value))}"`
        : "";
    })
    .join("");

/**
 * A WordprocessingML subtree re-spelled under `prefix`. Elements and
 * attributes of other namespaces are left out: this copies structure such as
 * a table grid, never content.
 */
export const wSubtreeXml = (element: XmlElement, prefix: string): string => {
  const name = wName(element);
  if (name === null) {
    return "";
  }
  const children = elementChildren(element)
    .map((child) => wSubtreeXml(child, prefix))
    .join("");
  const open = `<${prefix}:${name}${wAttributesXml(element, prefix)}`;
  return children === ""
    ? `${open}/>`
    : `${open}>${children}</${prefix}:${name}>`;
};

type RunsResult = Result<string, TemplatePlanError>;

/**
 * What stands in a paragraph's runs for a deleted paragraph mark until the
 * paragraph is written: the attributes of the mark between two of these. No
 * run holds the character (XML text cannot).
 */
const MARK_SPLIT = "\u0000";

/** The inserted mark of a paragraph's last piece, when it has one. */
const insertedMarkOf = (
  segments: readonly ExpectedSegment[],
): AtomRevision | undefined => {
  const last = segments.at(-1);
  const pieces =
    last === undefined || last.kind === "piece" ? [last?.piece] : last.pieces;
  const revision = pieces.at(-1)?.atom?.revision;
  return revision?.mark === true ? revision : undefined;
};

/** The attributes a kept tracked change is written with: its own author and
 *  date, and an id of the output's. */
const revisionAttributesXml = (
  { author, date }: AtomRevision,
  { prefixes: { w }, nextRevisionId }: EmitContext,
): string =>
  ` ${w}:id="${String(nextRevisionId())}" ${w}:author="${escapeXmlAttribute(author)}"${date === null ? "" : ` ${w}:date="${escapeXmlAttribute(date)}"`}`;

/** Deleted text as the content of a deleted run: tabs and line breaks are
 *  elements, the rest is `w:delText`. */
const deletedContentXml = (deleted: string, w: string): string =>
  deleted
    .split(/(\t|\n)/u)
    .map((part) => {
      if (part === "\t") {
        return `<${w}:tab/>`;
      }
      if (part === "\n") {
        return `<${w}:br/>`;
      }
      return part === ""
        ? ""
        : `<${w}:delText xml:space="preserve">${escapeXmlText(part)}</${w}:delText>`;
    })
    .join("");

/**
 * The runs of one piece. A piece of a kept tracked change is written inside
 * its insertion or deletion; a deletion holds the text that was deleted and
 * takes no place in the text around it.
 */
const pieceRunXml = (
  piece: ExpectedPiece,
  context: EmitContext,
  linked: boolean,
): RunsResult => {
  const revision = piece.atom?.revision;
  const checkbox = piece.atom?.checkbox;
  if (revision === undefined && checkbox !== undefined) {
    // The mark of a checkbox goes back into its content control, so the
    // box can still be ticked.
    const mark = plainPieceRunXml(piece, context, linked);
    if (Result.isError(mark)) {
      return mark;
    }
    const { w: prefix } = context.prefixes;
    return Result.ok(
      `<${prefix}:sdt>${captureVerbatimXml(checkbox)}<${prefix}:sdtContent>${mark.value}</${prefix}:sdtContent></${prefix}:sdt>`,
    );
  }
  if (revision === undefined) {
    return plainPieceRunXml(piece, context, linked);
  }
  const { w } = context.prefixes;
  // An inserted paragraph mark is written with its paragraph's properties.
  if (revision.mark === true) {
    return Result.ok("");
  }
  const attributes = revisionAttributesXml(revision, context);
  // A deleted paragraph mark ends a paragraph of its own where it stands.
  if (revision.deleted === DELETED_PARAGRAPH_MARK) {
    return Result.ok(`${MARK_SPLIT}${attributes}${MARK_SPLIT}`);
  }
  if (revision.kind === "del") {
    const properties = runPropertiesXml({
      prefix: w,
      signature: piece.signature,
      rStyle: linked ? context.styles.hyperlink : null,
      escapeAttribute: escapeXmlAttribute,
    });
    const deletion = `<${w}:del${attributes}><${w}:r>${properties}${deletedContentXml(revision.deleted ?? "", w)}</${w}:r></${w}:del>`;
    // Deleted out of an insertion: the insertion is on record around it.
    return Result.ok(
      revision.within === undefined
        ? deletion
        : `<${w}:ins${revisionAttributesXml({ kind: "ins", ...revision.within }, context)}>${deletion}</${w}:ins>`,
    );
  }
  return plainPieceRunXml(piece, context, linked).map((runs) =>
    runs === "" ? "" : `<${w}:ins${attributes}>${runs}</${w}:ins>`,
  );
};

const plainPieceRunXml = (
  piece: ExpectedPiece,
  context: EmitContext,
  linked: boolean,
): RunsResult => {
  const { w } = context.prefixes;
  const properties = (rStyle: string | null): string =>
    runPropertiesXml({
      prefix: w,
      signature: piece.signature,
      rStyle,
      escapeAttribute: escapeXmlAttribute,
    });
  const run = (content: string, rStyle: string | null = null): string =>
    `<${w}:r>${properties(rStyle ?? (linked ? context.styles.hyperlink : null))}${content}</${w}:r>`;
  switch (piece.kind) {
    case "text":
      return Result.ok(
        piece.text === ""
          ? ""
          : run(
              `<${w}:t xml:space="preserve">${escapeXmlText(piece.text)}</${w}:t>`,
            ),
      );
    case "tab":
      return Result.ok(run(`<${w}:tab/>`));
    case "lineBreak":
      return Result.ok(run(`<${w}:br/>`));
    case "pageBreak": {
      const type = wAttr(piece.atom?.element, "type") ?? "page";
      return Result.ok(
        run(`<${w}:br ${w}:type="${escapeXmlAttribute(type)}"/>`),
      );
    }
    case "noBreakHyphen":
      return Result.ok(run(`<${w}:noBreakHyphen/>`));
    case "softHyphen":
      return Result.ok(run(`<${w}:softHyphen/>`));
    case "symbol": {
      const element = piece.atom?.element;
      return Result.ok(
        element === undefined || element === null
          ? ""
          : run(wSubtreeXml(element, w)),
      );
    }
    case "footnoteReference":
    case "endnoteReference": {
      const kind: NoteKind =
        piece.kind === "footnoteReference" ? "footnote" : "endnote";
      const sourceId = wAttr(piece.atom?.element, "id");
      if (sourceId === null) {
        return Result.ok("");
      }
      const style =
        kind === "footnote"
          ? context.styles.footnoteReference
          : context.styles.endnoteReference;
      const mark = `<${w}:${kind}Reference ${w}:id="${context.noteId(kind, sourceId)}"/>`;
      return Result.ok(
        style === null
          ? `<${w}:r><${w}:rPr><${w}:vertAlign ${w}:val="superscript"/></${w}:rPr>${mark}</${w}:r>`
          : `<${w}:r><${w}:rPr><${w}:rStyle ${w}:val="${escapeXmlAttribute(style)}"/></${w}:rPr>${mark}</${w}:r>`,
      );
    }
    case "selfNumber": {
      const instruction = piece.atom?.instruction;
      // The field is written whole: it holds no result to write in its place.
      return Result.ok(
        instruction === undefined
          ? ""
          : [
              `<${w}:fldChar ${w}:fldCharType="begin"/>`,
              `<${w}:instrText xml:space="preserve"> ${escapeXmlText(instruction)} </${w}:instrText>`,
              `<${w}:fldChar ${w}:fldCharType="end"/>`,
            ]
              .map((content) => run(content))
              .join(""),
      );
    }
    case "object": {
      const transferred = transferObject(piece, context);
      if (Result.isError(transferred)) {
        return transferred;
      }
      // An equation is a sibling of runs; everything else is run content.
      return Result.ok(
        piece.atom?.element?.namespaceUri === MATH_NAMESPACE
          ? transferred.value
          : run(transferred.value),
      );
    }
    default:
      return piece.kind satisfies never;
  }
};

/** Adjacent text pieces that state the same emphasis and link become one run. */
const mergedPieces = (pieces: readonly ExpectedPiece[]): ExpectedPiece[] => {
  const merged: ExpectedPiece[] = [];
  for (const piece of pieces) {
    const last = merged.at(-1);
    if (
      last !== undefined &&
      last.kind === "text" &&
      piece.kind === "text" &&
      last.signature === piece.signature &&
      last.origin === piece.origin &&
      // A tracked change is a run of its own, and so is the text beside it.
      last.atom?.revision === undefined &&
      piece.atom?.revision === undefined &&
      // So is the mark of a checkbox: its control holds it alone.
      last.atom?.checkbox === undefined &&
      piece.atom?.checkbox === undefined &&
      sameLink(last.link, piece.link)
    ) {
      merged[merged.length - 1] = { ...last, text: last.text + piece.text };
    } else {
      merged.push(piece);
    }
  }
  return merged;
};

const piecesRunsXml = (
  pieces: readonly ExpectedPiece[],
  context: EmitContext,
): RunsResult => {
  const { w, r } = context.prefixes;
  const out: string[] = [];
  let group: {
    relationshipId: string;
    anchor: string | null;
    link: AtomLink;
    runs: string[];
  } | null = null;
  const flush = (): void => {
    if (group === null) {
      return;
    }
    out.push(
      `<${w}:hyperlink${context.prefixes.rDeclaration} ${r}:id="${escapeXmlAttribute(group.relationshipId)}"${group.anchor === null ? "" : ` ${w}:anchor="${escapeXmlAttribute(group.anchor)}"`} ${w}:history="1">${group.runs.join("")}</${w}:hyperlink>`,
    );
    group = null;
  };
  for (const piece of mergedPieces(pieces)) {
    const target = linkTargetOf(piece, context);
    if (
      group !== null &&
      (target === null || !sameLink(group.link, piece.link))
    ) {
      flush();
    }
    const run = pieceRunXml(piece, context, target !== null);
    if (Result.isError(run)) {
      return run;
    }
    if (target === null || piece.link === null) {
      out.push(run.value);
      continue;
    }
    group ??= { ...target, link: piece.link, runs: [] };
    group.runs.push(run.value);
  }
  flush();
  return Result.ok(out.join(""));
};

/** The runs of a segment list, references written as complex REF fields. */
export const segmentsRunsXml = (
  segments: readonly ExpectedSegment[],
  context: EmitContext,
  instructionOf: (segment: ExpectedSegment) => string | null,
): RunsResult => {
  const out: string[] = [];
  let plain: ExpectedPiece[] = [];
  const flush = (): Result<undefined, TemplatePlanError> => {
    if (plain.length === 0) {
      return Result.ok(undefined);
    }
    const runs = piecesRunsXml(plain, context);
    plain = [];
    if (Result.isError(runs)) {
      return runs;
    }
    out.push(runs.value);
    return Result.ok(undefined);
  };
  for (const segment of segments) {
    if (segment.kind === "piece") {
      plain.push(segment.piece);
      continue;
    }
    const flushed = flush();
    if (Result.isError(flushed)) {
      return flushed;
    }
    const result = piecesRunsXml(segment.pieces, context);
    if (Result.isError(result)) {
      return result;
    }
    const instruction =
      instructionOf(segment) ??
      panic("A resolved reference must carry its field instruction.");
    out.push(
      complexFieldXml({
        prefix: context.prefixes.w,
        instruction,
        resultRunsXml: result.value,
        runPropertiesXml: runPropertiesXml({
          prefix: context.prefixes.w,
          signature: segment.pieces.at(0)?.signature ?? "",
          escapeAttribute: escapeXmlAttribute,
        }),
      }),
    );
  }
  const flushed = flush();
  return Result.isError(flushed) ? flushed : Result.ok(out.join(""));
};

export type ParagraphEmit = {
  paragraph: ResolvedParagraph;
  bookmark: BookmarkWrap | undefined;
  instructionOf: (segment: ExpectedSegment) => string | null;
  /** Whether the paragraph's style binds a list that `numbering: null` must cancel. */
  styleBindsList: boolean;
  /** The `w:sectPr` of the section the paragraph is the last of, or null. */
  sectionXml: string | null;
};

/** Runs of a placed paragraph, inside its bookmark when it is a reference target. */
export const paragraphContentXml = (
  { paragraph, bookmark, instructionOf }: ParagraphEmit,
  context: EmitContext,
): RunsResult => {
  const runs = segmentsRunsXml(paragraph.segments, context, instructionOf);
  if (Result.isError(runs) || bookmark === undefined) {
    return runs;
  }
  const { w } = context.prefixes;
  return Result.ok(
    `${bookmarkStartXml({ prefix: w, ...bookmark })}${runs.value}${bookmarkEndXml({ prefix: w, id: bookmark.id })}`,
  );
};

const NO_LIST_NUM_ID = 0;

/** A whole `w:p` for a placed source paragraph. */
export const paragraphXml = (
  emit: ParagraphEmit,
  context: EmitContext,
): RunsResult => {
  const { w } = context.prefixes;
  const { styleId, numbering } = emit.paragraph;
  const properties: string[] = [];
  if (styleId !== null) {
    properties.push(`<${w}:pStyle ${w}:val="${escapeXmlAttribute(styleId)}"/>`);
  }
  // `w:pPr` keeps its children in schema order: keeping with the next
  // paragraph and the page start stand between the style and the list.
  if (emit.paragraph.keepNext) {
    properties.push(`<${w}:keepNext/>`);
  }
  if (emit.paragraph.pageBreakBefore) {
    properties.push(`<${w}:pageBreakBefore/>`);
  }
  if (numbering.kind === "direct") {
    properties.push(
      `<${w}:numPr><${w}:ilvl ${w}:val="${numbering.ilvl}"/><${w}:numId ${w}:val="${numbering.numId}"/></${w}:numPr>`,
    );
  } else if (numbering.kind === "none" && emit.styleBindsList) {
    properties.push(
      `<${w}:numPr><${w}:numId ${w}:val="${NO_LIST_NUM_ID}"/></${w}:numPr>`,
    );
  }
  // Tab stops follow the list and stand in front of the indent.
  const { indent, tabStops } = emit.paragraph;
  if (tabStops !== null) {
    const stops = tabStops
      .map((stop) => `<${w}:tab ${w}:val="left" ${w}:pos="${String(stop)}"/>`)
      .join("");
    properties.push(`<${w}:tabs>${stops}</${w}:tabs>`);
  }
  // The space in front and behind stands between the stops and the indent.
  const { spacing } = emit.paragraph;
  if (spacing !== null) {
    const sides = (["before", "after"] as const)
      .filter((side) => spacing[side] !== undefined)
      .map((side) => ` ${w}:${side}="${String(spacing[side])}"`)
      .join("");
    properties.push(`<${w}:spacing${sides}/>`);
  }
  if (indent !== null) {
    const sides = (["left", "hanging", "firstLine"] as const)
      .filter((side) => indent[side] !== undefined)
      .map((side) => ` ${w}:${side}="${String(indent[side])}"`)
      .join("");
    properties.push(`<${w}:ind${sides}/>`);
  }
  // The alignment follows the indent.
  const { alignment } = emit.paragraph;
  if (alignment !== null) {
    properties.push(`<${w}:jc ${w}:val="${alignment}"/>`);
  }
  // A section's properties close the paragraph's.
  if (emit.sectionXml !== null) {
    properties.push(emit.sectionXml);
  }
  const runs = segmentsRunsXml(
    emit.paragraph.segments,
    context,
    emit.instructionOf,
  );
  if (Result.isError(runs)) {
    return runs;
  }
  // The paragraph's mark, when the source records it as inserted: its run
  // properties stand in front of a section's.
  const inserted = insertedMarkOf(emit.paragraph.segments);
  if (inserted !== undefined) {
    const mark = `<${w}:rPr><${w}:ins${revisionAttributesXml(inserted, context)}/></${w}:rPr>`;
    properties.splice(
      emit.sectionXml === null ? properties.length : properties.length - 1,
      0,
      mark,
    );
  }
  const pPr =
    properties.length === 0
      ? ""
      : `<${w}:pPr>${properties.join("")}</${w}:pPr>`;
  // A deleted paragraph mark inside the block: what stands in front of it
  // is a paragraph of its own, in the block's style and in no list, whose
  // mark is the deleted one. The block's own paragraph comes last.
  const [last = "", ...earlier] = runs.value.split(MARK_SPLIT).toReversed();
  const halves: string[] = [];
  const parted = earlier.toReversed();
  for (let at = 0; at + 1 < parted.length; at += 2) {
    const style =
      styleId === null
        ? ""
        : `<${w}:pStyle ${w}:val="${escapeXmlAttribute(styleId)}"/>`;
    const list =
      numbering.kind === "direct" || emit.styleBindsList
        ? `<${w}:numPr><${w}:numId ${w}:val="${NO_LIST_NUM_ID}"/></${w}:numPr>`
        : "";
    halves.push(
      `<${w}:p><${w}:pPr>${style}${list}<${w}:rPr><${w}:del${parted[at + 1] ?? ""}/></${w}:rPr></${w}:pPr>${parted[at] ?? ""}</${w}:p>`,
    );
  }
  const xml = `<${w}:p>${pPr}${last}</${w}:p>`;
  if (emit.bookmark === undefined) {
    return Result.ok(`${halves.join("")}${xml}`);
  }
  const wrapped = wrapParagraphContentInBookmark({
    paragraphXml: xml,
    prefixes: context.prefixes,
    ...emit.bookmark,
  });
  // The bookmark marks the block's own paragraph, which bears its number.
  return wrapped === null
    ? Result.err(
        new TemplatePlanError({
          code: "xml-unreadable",
          message: "A reference target could not be bookmarked.",
          blockId: emit.paragraph.parts.at(0)?.block.block.id ?? "",
        }),
      )
    : Result.ok(`${halves.join("")}${wrapped}`);
};

const TABLE_LAYOUT_PROPERTIES = ["tblLayout", "tblLook"] as const;
const ROW_PROPERTIES = [
  "gridBefore",
  "gridAfter",
  "cantSplit",
  "tblHeader",
] as const;
const CELL_SPAN_PROPERTIES = ["tcW", "gridSpan", "hMerge", "vMerge"] as const;
/** What a cell's fill states by itself; a theme colour is the template's. */
const SHADING_KEPT = ["val", "color", "fill"] as const;

/**
 * The fill of a cell as its own colours state it. A theme colour would be
 * read from the template's theme, which is another one; a fill stated by a
 * theme colour alone is left out.
 */
const cellShadingXml = (
  properties: XmlElement | null,
  prefix: string,
): string => {
  const shading = wChild(properties, "shd");
  const fill = wAttr(shading, "fill");
  if (shading === null || fill === null || fill === "auto") {
    return "";
  }
  const kept = SHADING_KEPT.map((name) => [name, wAttr(shading, name)] as const)
    .filter(([, value]) => value !== null)
    .map(
      ([name, value]) =>
        ` ${prefix}:${name}="${escapeXmlAttribute(value ?? "")}"`,
    )
    .join("");
  return `<${prefix}:shd${kept}/>`;
};

const keptChildrenXml = (
  properties: XmlElement | null,
  names: readonly string[],
  prefix: string,
): string =>
  names
    .map((name) => {
      const child = wChild(properties, name);
      return child === null ? "" : wSubtreeXml(child, prefix);
    })
    .join("");

/** The attributes a kept row or mark revision is written with, read off the
 *  source's own record. */
const markAttributesXml = (mark: XmlElement, context: EmitContext): string =>
  revisionAttributesXml(
    {
      kind: "del",
      author: wAttr(mark, "author") ?? "",
      date: wAttr(mark, "date"),
    },
    context,
  );

/**
 * A row the source records as deleted, written back as one: its cells with
 * their spans, and in each cell its paragraphs' text as deleted text with a
 * deleted paragraph mark. The accepted view has no such row, so no plan
 * places its paragraphs; they take no style of their own.
 */
const deletedRowXml = (
  row: XmlElement,
  keepsBorders: boolean,
  context: EmitContext,
): string => {
  const { w } = context.prefixes;
  const mark = deletedRowMarkOf(row);
  const attributes = (): string =>
    mark === null ? "" : markAttributesXml(mark, context);
  const deletedParagraphXml = (text: string): string =>
    `<${w}:p><${w}:pPr><${w}:rPr><${w}:del${attributes()}/></${w}:rPr></${w}:pPr>${text === "" ? "" : `<${w}:del${attributes()}><${w}:r>${deletedContentXml(text, w)}</${w}:r></${w}:del>`}</${w}:p>`;
  const cells: string[] = [];
  const stack = [...elementChildren(row)].toReversed();
  for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
    if (wName(next) === "tc") {
      const properties = wChild(next, "tcPr");
      const kept = `${keptChildrenXml(properties, CELL_SPAN_PROPERTIES, w)}${keepsBorders ? keptChildrenXml(properties, ["tcBorders"], w) : ""}`;
      // The cell's text as one row text gives it: a paragraph per line.
      const lines = rowTextOf({ ...row, elements: [next] }).split("\n");
      cells.push(
        `<${w}:tc>${kept === "" ? "" : `<${w}:tcPr>${kept}</${w}:tcPr>`}${lines.map(deletedParagraphXml).join("")}</${w}:tc>`,
      );
    } else if (wName(next) !== "trPr") {
      stack.push(...elementChildren(next).toReversed());
    }
  }
  return `<${w}:tr><${w}:trPr><${w}:del${attributes()}/></${w}:trPr>${cells.join("")}</${w}:tr>`;
};

export type TableEmit = {
  table: XmlElement;
  /** A table style of the template, or null to keep the source's own borders. */
  styleId: string | null;
  paragraphOf: (element: XmlElement) => RunsResult;
};

/**
 * A source table rebuilt: its grid, spans and merges, with every cell
 * paragraph placed like a body paragraph. Without a template table style the
 * source's borders are kept so the grid stays visible, and the fill of its
 * cells with them: a header row, the colours of a legend.
 */
export const tableXml = (
  { table, styleId, paragraphOf }: TableEmit,
  context: EmitContext,
): RunsResult => {
  const { w } = context.prefixes;
  const blocksXml = (container: XmlElement): RunsResult => {
    const out: string[] = [];
    for (const child of elementChildren(container)) {
      const name = wName(child);
      let result: RunsResult = Result.ok("");
      if (name === "p") {
        result = paragraphOf(child);
      } else if (name === "tbl") {
        result = tableXml({ table: child, styleId, paragraphOf }, context);
      } else if (name === "sdt") {
        const content = wChild(child, "sdtContent");
        result = content === null ? Result.ok("") : blocksXml(content);
      } else if (name === "customXml") {
        result = blocksXml(child);
      }
      if (Result.isError(result)) {
        return result;
      }
      out.push(result.value);
    }
    return Result.ok(out.join(""));
  };
  const cellsXml = (container: XmlElement): RunsResult => {
    const out: string[] = [];
    for (const child of elementChildren(container)) {
      const name = wName(child);
      if (name === "tc") {
        const cellProperties = wChild(child, "tcPr");
        const kept = `${keptChildrenXml(cellProperties, CELL_SPAN_PROPERTIES, w)}${styleId === null ? `${keptChildrenXml(cellProperties, ["tcBorders"], w)}${cellShadingXml(cellProperties, w)}` : ""}${keptChildrenXml(cellProperties, ["vAlign"], w)}`;
        const blocks = blocksXml(child);
        if (Result.isError(blocks)) {
          return blocks;
        }
        // A cell ends in a paragraph, whatever its last block was.
        const content = blocks.value.endsWith(`</${w}:p>`)
          ? blocks.value
          : `${blocks.value}<${w}:p/>`;
        out.push(
          `<${w}:tc>${kept === "" ? "" : `<${w}:tcPr>${kept}</${w}:tcPr>`}${content}</${w}:tc>`,
        );
      } else if (name === "sdt") {
        const content = wChild(child, "sdtContent");
        const inner = content === null ? Result.ok("") : cellsXml(content);
        if (Result.isError(inner)) {
          return inner;
        }
        out.push(inner.value);
      } else if (name === "customXml") {
        const inner = cellsXml(child);
        if (Result.isError(inner)) {
          return inner;
        }
        out.push(inner.value);
      }
    }
    return Result.ok(out.join(""));
  };
  const rowsXml = (container: XmlElement): RunsResult => {
    const out: string[] = [];
    for (const child of elementChildren(container)) {
      const name = wName(child);
      const deleted = keptDeletedRowOf(child);
      if (deleted !== null) {
        out.push(deletedRowXml(deleted, styleId === null, context));
      } else if (name === "tr") {
        const inserted = insertedRowMarkOf(child);
        // A row the source records as inserted keeps that record.
        const kept = `${keptChildrenXml(wChild(child, "trPr"), ROW_PROPERTIES, w)}${inserted === null ? "" : `<${w}:ins${markAttributesXml(inserted, context)}/>`}`;
        const cells = cellsXml(child);
        if (Result.isError(cells)) {
          return cells;
        }
        out.push(
          `<${w}:tr>${kept === "" ? "" : `<${w}:trPr>${kept}</${w}:trPr>`}${cells.value}</${w}:tr>`,
        );
      } else if (name === "sdt") {
        const content = wChild(child, "sdtContent");
        const inner = content === null ? Result.ok("") : rowsXml(content);
        if (Result.isError(inner)) {
          return inner;
        }
        out.push(inner.value);
      } else if (name === "customXml") {
        const inner = rowsXml(child);
        if (Result.isError(inner)) {
          return inner;
        }
        out.push(inner.value);
      }
    }
    return Result.ok(out.join(""));
  };

  const tableProperties = wChild(table, "tblPr");
  const style =
    styleId === null
      ? ""
      : `<${w}:tblStyle ${w}:val="${escapeXmlAttribute(styleId)}"/>`;
  const properties = `${style}${keptChildrenXml(tableProperties, ["tblW"], w)}${styleId === null ? keptChildrenXml(tableProperties, ["tblBorders"], w) : ""}${keptChildrenXml(tableProperties, TABLE_LAYOUT_PROPERTIES, w)}`;
  // Columns only: a grid may also carry the record of a tracked change.
  const columns = wChildren(wChild(table, "tblGrid"), "gridCol")
    .map((column) => wSubtreeXml(column, w))
    .join("");
  const gridXml =
    columns === "" ? "" : `<${w}:tblGrid>${columns}</${w}:tblGrid>`;
  const rows = rowsXml(table);
  if (Result.isError(rows)) {
    return rows;
  }
  return Result.ok(
    `<${w}:tbl><${w}:tblPr>${properties}</${w}:tblPr>${gridXml}${rows.value}</${w}:tbl>`,
  );
};

export type NoteEmit = {
  kind: NoteKind;
  id: number;
  note: AnalyzedNote;
  /** Styles of the source, to read what the note's runs state. */
  source: SignatureContext;
  /** The fields of the source notes part. */
  fields: readonly StoryField[];
};

/** A source note under its new id, its paragraphs in the template's note style. */
export const noteXml = (
  { kind, id, note, source, fields }: NoteEmit,
  context: EmitContext,
): RunsResult => {
  const { w } = context.prefixes;
  const textStyle =
    kind === "footnote"
      ? context.styles.footnoteText
      : context.styles.endnoteText;
  const referenceStyle =
    kind === "footnote"
      ? context.styles.footnoteReference
      : context.styles.endnoteReference;
  const origin: PieceOrigin = kind === "footnote" ? "footnotes" : "endnotes";
  const pPr =
    textStyle === null
      ? ""
      : `<${w}:pPr><${w}:pStyle ${w}:val="${escapeXmlAttribute(textStyle)}"/></${w}:pPr>`;
  const mark = `<${w}:r><${w}:rPr>${
    referenceStyle === null
      ? `<${w}:vertAlign ${w}:val="superscript"/>`
      : `<${w}:rStyle ${w}:val="${escapeXmlAttribute(referenceStyle)}"/>`
  }</${w}:rPr><${w}:${kind}Ref/></${w}:r>`;
  const paragraphs: string[] = [];
  for (const [index, read] of note.paragraphs.entries()) {
    const runs = piecesRunsXml(
      paragraphPieces({ read, context: source, origin, fields }),
      context,
    );
    if (Result.isError(runs)) {
      return runs;
    }
    paragraphs.push(
      `<${w}:p>${pPr}${index === 0 ? mark : ""}${runs.value}</${w}:p>`,
    );
  }
  if (paragraphs.length === 0) {
    paragraphs.push(`<${w}:p>${pPr}${mark}</${w}:p>`);
  }
  return Result.ok(
    `<${w}:${kind} ${w}:id="${id}">${paragraphs.join("")}</${w}:${kind}>`,
  );
};
