/**
 * The content of a source's page headers and footers carried onto the
 * template's.
 *
 * The output's headers and footers are the template's. With
 * `sourceHeaders: "keep"` the content of the headers and footers the
 * source's first section shows (text, tables, pictures, fields) is added
 * after the template's own content, in the header or footer that shows on
 * the same pages: the first page, the other pages, even pages. Bookmarks
 * and comment marks are left out of the copy; pictures are copied with
 * their parts. A header or footer that shows nothing (empty paragraphs)
 * adds nothing.
 *
 * Pages map as the source shows them: without a distinct first page the
 * source's first page shows its default header; without distinct even
 * pages its even pages do too. Where the output has no distinct first page
 * (or even pages) and the source shows other content there than on its
 * other pages, that content is left out and counted, as the template
 * decides which pages differ. Where the template has no header or footer
 * for those pages, a new part holds the carried content.
 */
import { Result } from "better-result";

import { resolveNamespaceUri } from "@stll/docx-core/xml/namespaces";

import type { RelationshipMap } from "../../../types/document";
import {
  parseRelationships,
  RELATIONSHIP_TYPES,
  resolveRelativePath,
} from "../../relsParser";
import { captureVerbatimXml } from "../../verbatimCapture";
import {
  cloneElement,
  getLocalName,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  type XmlElement,
} from "../../xmlParser";
import type { AnalyzedDocument } from "./blocks";
import { EMBEDDED_RELATIONSHIP_TYPES, extensionOf } from "./embeddedParts";
import type { RelationshipsEditor, SourceMedia } from "./emit";
import { readPartText, relationshipsPathFor } from "./package";
import type { ResolvedPlan } from "./plan";
import { createStoryState, readParagraph, toggleOf, walkBlocks } from "./story";
import type { TemplatePlanError, TemplatePlanSourceHeaders } from "./types";
import {
  elementChildren,
  rangeOf,
  readRawPart,
  relationshipAttr,
  wAttr,
  wChild,
  wName,
} from "./xml";

export type HeaderKind = "first" | "default" | "even";
export type HeaderFamily = "header" | "footer";

const HEADER_KINDS: readonly HeaderKind[] = ["first", "default", "even"];
const FAMILIES: readonly HeaderFamily[] = ["header", "footer"];

const REFERENCE: Record<HeaderFamily, string> = {
  header: "headerReference",
  footer: "footerReference",
};

type LeftOut = TemplatePlanSourceHeaders["leftOut"];

const NO_LEFT_OUT = (): LeftOut => ({
  firstPageParts: 0,
  evenPageParts: 0,
  sharedParts: 0,
  unsupportedObjects: 0,
  laterSectionParts: 0,
});

/** A source header or footer part, read. */
export type SourcePart = {
  path: string;
  root: XmlElement;
  relationships: RelationshipMap;
  /** Its top-level blocks a copy carries: every one that refers to no
   *  embedded object other than a picture. */
  blocks: XmlElement[];
  /** Top-level blocks a copy cannot carry. */
  unsupported: number;
  /** The texts of the carried blocks' paragraphs, tables' included. */
  texts: string[];
};

/** One output header or footer that takes a source part's content: the
 *  template's part it is appended to, or null where a new part holds it. */
export type HeaderWrite = {
  family: HeaderFamily;
  kind: HeaderKind;
  source: SourcePart;
  templatePath: string | null;
  /** The output section a later section's content goes to: a new part
   *  that section alone names. Absent for the content every section shows. */
  section?: number;
};

const SHOWN: ReadonlySet<string> = new Set([
  "drawing",
  "pict",
  "object",
  "fldSimple",
  "fldChar",
]);

/** Whether a part shows anything: text, a picture, a field. */
const showsAnything = (root: XmlElement, texts: readonly string[]): boolean => {
  if (texts.some((text) => text.trim() !== "")) {
    return true;
  }
  const stack = [root];
  while (stack.length > 0) {
    const element = stack.pop();
    if (element === undefined) {
      continue;
    }
    if (SHOWN.has(getLocalName(element.name ?? ""))) {
      return true;
    }
    stack.push(...elementChildren(element));
  }
  return false;
};

/** Whether every relationship a block refers to can be re-made in an
 *  output part: an external address, or a picture or embedded object the
 *  package holds. */
const carriable = (
  block: XmlElement,
  relationships: RelationshipMap,
  holds: (target: string) => boolean,
): boolean => {
  const stack = [block];
  while (stack.length > 0) {
    const node = stack.pop();
    if (node === undefined) {
      continue;
    }
    for (const [name, value] of Object.entries(node.attributes ?? {})) {
      const colon = name.indexOf(":");
      if (value === undefined || colon === -1 || name.startsWith("xmlns:")) {
        continue;
      }
      const uri = resolveNamespaceUri(
        node.namespaceScope,
        name.slice(0, colon),
      );
      if (!OFFICE_RELATIONSHIP_NAMESPACE_URIS.has(uri ?? "")) {
        continue;
      }
      const relationship = relationships.get(String(value));
      const fine =
        relationship !== undefined &&
        (relationship.targetMode === "External" ||
          (EMBEDDED_RELATIONSHIP_TYPES.has(relationship.type) &&
            holds(relationship.target)));
      if (!fine) {
        return false;
      }
    }
    stack.push(...elementChildren(node));
  }
  return true;
};

/** The texts of a part's paragraphs as the reader reads them: field
 *  results kept, instructions left out. */
export const partParagraphTexts = (root: XmlElement): string[] => {
  const state = createStoryState();
  const texts: string[] = [];
  walkBlocks(root, {
    paragraph: (paragraph) => {
      texts.push(readParagraph(paragraph, state).text);
    },
    tableStart: () => {},
    cellStart: () => {},
    tableEnd: () => {},
    unread: () => {},
    tracked: () => {},
  });
  return texts;
};

/** The first section's properties: the first a paragraph ends, else the body's. */
export const firstSectionOf = (body: XmlElement): XmlElement | null => {
  for (const block of elementChildren(body)) {
    if (wName(block) === "p") {
      const sectPr = wChild(wChild(block, "pPr"), "sectPr");
      if (sectPr !== null) {
        return sectPr;
      }
    }
  }
  return wChild(body, "sectPr");
};

/** The relationship ids a section names for one family, by kind. */
const referencesOf = (
  sectPr: XmlElement | null,
  family: HeaderFamily,
): Map<HeaderKind, string> => {
  const out = new Map<HeaderKind, string>();
  for (const child of elementChildren(sectPr)) {
    if (wName(child) !== REFERENCE[family]) {
      continue;
    }
    const kind = (wAttr(child, "type") ?? "default") as HeaderKind;
    const id = relationshipAttr(child, "id");
    if (HEADER_KINDS.includes(kind) && id !== null) {
      out.set(kind, id);
    }
  }
  return out;
};

/**
 * The relationship ids every section of the output names for one family,
 * by kind: the output's sections are built from the template's last one,
 * which names its own and takes on those the sections before it named.
 */
export const outputReferences = (
  body: XmlElement,
  family: HeaderFamily,
): Map<HeaderKind, string> => {
  const out = new Map<HeaderKind, string>();
  for (const block of elementChildren(body)) {
    const own = wName(block) === "sectPr" ? block : null;
    const sectPr =
      wName(block) === "p" ? wChild(wChild(block, "pPr"), "sectPr") : own;
    for (const [kind, id] of referencesOf(sectPr, family)) {
      out.set(kind, id);
    }
  }
  return out;
};

/** The part a header or footer relationship of the main document points to. */
export const familyPathOf = (
  { pkg }: AnalyzedDocument,
  family: HeaderFamily,
  id: string,
): string | null => {
  const relationship = pkg.documentRelationships.get(id);
  if (
    relationship === undefined ||
    relationship.type !== RELATIONSHIP_TYPES[family] ||
    relationship.targetMode === "External"
  ) {
    return null;
  }
  const path = pkg.resolveDocumentTarget(relationship.target);
  return pkg.paths.has(path) ? path : null;
};

/** Whether a document's settings give even pages headers of their own. */
export const evenPagesDiffer = async ({
  pkg,
}: AnalyzedDocument): Promise<Result<boolean, TemplatePlanError>> => {
  const xml = await readPartText(pkg, pkg.partPath("settings"));
  if (Result.isError(xml)) {
    return xml;
  }
  const part = xml.value === null ? null : readRawPart("settings", xml.value);
  return Result.ok(
    toggleOf(wChild(part?.root ?? null, "evenAndOddHeaders")) === true,
  );
};

/** A part of a document read with its relationships; null when absent. */
export const readFamilyPart = async (
  document: AnalyzedDocument,
  path: string | null,
): Promise<Result<SourcePart | null, TemplatePlanError>> => {
  const xml = await readPartText(document.pkg, path);
  if (Result.isError(xml)) {
    return xml;
  }
  const part =
    path === null || xml.value === null ? null : readRawPart(path, xml.value);
  if (path === null || part === null) {
    return Result.ok(null);
  }
  const relationshipsPath = relationshipsPathFor(path);
  const relationshipsXml = document.pkg.paths.has(relationshipsPath)
    ? await readPartText(document.pkg, relationshipsPath)
    : Result.ok(null);
  if (Result.isError(relationshipsXml)) {
    return relationshipsXml;
  }
  const relationships = parseRelationships(relationshipsXml.value ?? "");
  const holds = (target: string): boolean =>
    document.pkg.paths.has(resolveRelativePath(relationshipsPath, target));
  const all = elementChildren(part.root).filter(
    (child) => wName(child) !== "sectPr",
  );
  const blocks = all.filter((block) => carriable(block, relationships, holds));
  return Result.ok({
    path,
    root: part.root,
    relationships,
    blocks,
    unsupported: all.length - blocks.length,
    texts: partParagraphTexts(cloneElement(part.root, { elements: blocks })),
  });
};

const sameContent = (a: SourcePart | null, b: SourcePart | null): boolean =>
  a === b ||
  (a !== null &&
    b !== null &&
    (a.path === b.path ||
      (a.texts.length === b.texts.length &&
        a.texts.every((text, index) => text === b.texts[index]))));

type FamilyPages = {
  /** Whether the output's first page shows a header of its own. */
  titlePage: boolean;
  /** Whether the output's even pages do. */
  evenPages: boolean;
  /** Whether the source's even pages do. */
  sourceEven: boolean;
};

/** The sections of a body in order: the properties each section-ending
 *  paragraph holds, then the body's own. */
export const sectionsOf = (body: XmlElement): XmlElement[] => {
  const sections: XmlElement[] = [];
  for (const block of elementChildren(body)) {
    const sectPr =
      wName(block) === "p" ? wChild(wChild(block, "pPr"), "sectPr") : null;
    if (sectPr !== null) {
      sections.push(sectPr);
    }
  }
  const last = wChild(body, "sectPr");
  if (last !== null) {
    sections.push(last);
  }
  return sections;
};

/** The relationship ids section `index` shows for one family, by kind: its
 *  own, else those of the sections before it. */
const sectionReferences = (
  sections: readonly XmlElement[],
  index: number,
  family: HeaderFamily,
): Map<HeaderKind, string> => {
  const out = new Map<HeaderKind, string>();
  for (const section of sections.slice(0, index + 1)) {
    for (const [kind, id] of referencesOf(section, family)) {
      out.set(kind, id);
    }
  }
  return out;
};

/** What a source section shows for one family, by the output kind that
 *  shows on the same pages. */
const carryFamily = async (
  source: AnalyzedDocument,
  family: HeaderFamily,
  section: XmlElement | null,
  { titlePage, evenPages, sourceEven }: FamilyPages,
  leftOut: LeftOut,
  references: Map<HeaderKind, string> = referencesOf(section, family),
): Promise<Result<Map<HeaderKind, SourcePart>, TemplatePlanError>> => {
  const shown = new Map<HeaderKind, SourcePart | null>();
  for (const kind of HEADER_KINDS) {
    const id = references.get(kind);
    const path = id === undefined ? null : familyPathOf(source, family, id);
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const part = await readFamilyPart(source, path);
    if (Result.isError(part)) {
      return part;
    }
    shown.set(
      kind,
      part.value !== null && showsAnything(part.value.root, part.value.texts)
        ? part.value
        : null,
    );
  }
  const others = shown.get("default") ?? null;
  const firstPage =
    toggleOf(wChild(section, "titlePg")) === true
      ? (shown.get("first") ?? null)
      : others;
  const even = sourceEven ? (shown.get("even") ?? null) : others;
  const out = new Map<HeaderKind, SourcePart>();
  if (others !== null) {
    out.set("default", others);
  }
  if (titlePage) {
    if (firstPage !== null) {
      out.set("first", firstPage);
    }
  } else if (firstPage !== null && !sameContent(firstPage, others)) {
    leftOut.firstPageParts += 1;
  }
  if (evenPages) {
    if (even !== null) {
      out.set("even", even);
    }
  } else if (even !== null && !sameContent(even, others)) {
    leftOut.evenPageParts += 1;
  }
  return Result.ok(out);
};

/**
 * Where the content of the source's headers and footers goes. A template
 * part that two kinds share takes it only when every kind it serves takes
 * the same content; the others are left out and counted.
 */
export const plannedHeaderWrites = async (
  source: AnalyzedDocument,
  template: AnalyzedDocument,
  /** For each output section after the first, the index of the source
   *  section it starts with. */
  starts: readonly number[] = [],
): Promise<
  Result<{ writes: HeaderWrite[]; leftOut: LeftOut }, TemplatePlanError>
> => {
  const leftOut = NO_LEFT_OUT();
  const templateEven = await evenPagesDiffer(template);
  if (Result.isError(templateEven)) {
    return templateEven;
  }
  const sourceEven = await evenPagesDiffer(source);
  if (Result.isError(sourceEven)) {
    return sourceEven;
  }
  const pages: FamilyPages = {
    titlePage:
      toggleOf(wChild(wChild(template.body, "sectPr"), "titlePg")) === true,
    evenPages: templateEven.value,
    sourceEven: sourceEven.value,
  };
  const section = firstSectionOf(source.body);
  const writes: HeaderWrite[] = [];
  for (const family of FAMILIES) {
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const carried = await carryFamily(source, family, section, pages, leftOut);
    if (Result.isError(carried)) {
      return carried;
    }
    const references = outputReferences(template.body, family);
    const pathOf = (kind: HeaderKind): string | null => {
      const id = references.get(kind);
      return id === undefined ? null : familyPathOf(template, family, id);
    };
    const written = new Set<string>();
    for (const kind of HEADER_KINDS) {
      const part = carried.value.get(kind);
      if (part === undefined) {
        continue;
      }
      const path = pathOf(kind);
      if (path === null) {
        writes.push({ family, kind, source: part, templatePath: null });
        continue;
      }
      const agree = HEADER_KINDS.filter(
        (other) => other !== kind && pathOf(other) === path,
      ).every((other) => sameContent(carried.value.get(other) ?? null, part));
      if (!agree) {
        leftOut.sharedParts += 1;
        continue;
      }
      if (!written.has(path)) {
        written.add(path);
        writes.push({ family, kind, source: part, templatePath: path });
      }
    }
  }
  const later = await laterSectionWrites(
    source,
    template,
    starts,
    pages,
    leftOut,
  );
  if (Result.isError(later)) {
    return later;
  }
  writes.push(...later.value);
  leftOut.unsupportedObjects = writes.reduce(
    (total, { source: part }) => total + part.unsupported,
    0,
  );
  return Result.ok({ writes, leftOut });
};

/** Later output sections show no first page of their own. */
const LATER_KINDS: readonly HeaderKind[] = ["default", "even"];

/**
 * Where the content of the source's later sections goes. An output section
 * after the first starts with a source section; where that section shows
 * other content than the first one for some pages, a new part holds the
 * template's content for those pages followed by that content, and only
 * that output section names it. A source section the output merges into
 * another shows its content nowhere: each part with other content than
 * the section it went into is counted.
 */
const laterSectionWrites = async (
  source: AnalyzedDocument,
  template: AnalyzedDocument,
  starts: readonly number[],
  pages: FamilyPages,
  leftOut: LeftOut,
): Promise<Result<HeaderWrite[], TemplatePlanError>> => {
  const sections = sectionsOf(source.body);
  const representatives = [0, ...starts];
  const writes: HeaderWrite[] = [];
  for (const family of FAMILIES) {
    const shown: Map<HeaderKind, SourcePart>[] = [];
    for (const [index, section] of sections.entries()) {
      // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
      const carried = await carryFamily(
        source,
        family,
        section,
        { ...pages, titlePage: index === 0 && pages.titlePage },
        // What a later section has no place for is counted below.
        NO_LEFT_OUT(),
        sectionReferences(sections, index, family),
      );
      if (Result.isError(carried)) {
        return carried;
      }
      shown.push(carried.value);
    }
    const first = shown[0] ?? new Map<HeaderKind, SourcePart>();
    const references = outputReferences(template.body, family);
    const counted = new Set<string>();
    for (const [index, parts] of shown.entries()) {
      // The output section the source section went into.
      let output = 0;
      for (const [position, start] of representatives.entries()) {
        if (start <= index) {
          output = position;
        }
      }
      const representative = shown[representatives[output] ?? 0] ?? first;
      for (const kind of LATER_KINDS) {
        const part = parts.get(kind) ?? null;
        if (part === null) {
          continue;
        }
        if (output > 0 && index === representatives[output]) {
          if (!sameContent(part, first.get(kind) ?? null)) {
            const id = references.get(kind);
            writes.push({
              family,
              kind,
              source: part,
              templatePath:
                id === undefined ? null : familyPathOf(template, family, id),
              section: output,
            });
          }
          continue;
        }
        if (
          index > 0 &&
          !sameContent(part, representative.get(kind) ?? null) &&
          !counted.has(part.path)
        ) {
          counted.add(part.path);
          leftOut.laterSectionParts += 1;
        }
      }
    }
  }
  return Result.ok(writes);
};

/** The source paragraphs whose section the output keeps, in order. */
export const keptSectionEnds = (
  resolved: ResolvedPlan,
): (XmlElement | null)[] =>
  resolved.items.flatMap((item) => {
    if (item.kind !== "paragraph" || item.paragraph.endsSection === null) {
      return [];
    }
    const { parts } = item.paragraph;
    const ending =
      parts.find(({ block }) => block.read?.hasSectionBreak === true) ??
      parts.at(-1);
    return [ending?.block.element ?? null];
  });

/** The source section each output section after the first starts with:
 *  the one after the section its preceding section-ending paragraph ends.
 *  `ends` are the source paragraphs whose section the output keeps. */
export const outputSectionStarts = (
  source: AnalyzedDocument,
  ends: readonly (XmlElement | null)[],
): number[] => {
  const sections = sectionsOf(source.body);
  return ends.map((paragraph) => {
    const sectPr =
      paragraph === null ? null : wChild(wChild(paragraph, "pPr"), "sectPr");
    const at = sectPr === null ? undefined : rangeOf(sectPr)?.start;
    const index =
      at === undefined
        ? -1
        : sections.findIndex((section) => rangeOf(section)?.start === at);
    return index === -1 ? sections.length : index + 1;
  });
};

/** What a copy needs of the output: where its relationships go, how a
 *  picture's part is added, and the next free drawing id. */
export type CopyTarget = {
  relationships: RelationshipsEditor;
  addMedia: (sourcePath: string, media: SourceMedia) => string;
  nextDrawingId: () => number;
  /** Reads a part of the source package; null when it holds none. */
  readSourceBytes: (path: string) => Promise<Uint8Array | null>;
  /** The source path a relationship target of `fromPart` names. */
  resolveSourceTarget: (fromPart: string, target: string) => string;
  /** Whether the output defines a style of this id. */
  hasStyle: (styleId: string) => boolean;
};

/** Elements that name a style by its id. */
const STYLE_REFERENCES: ReadonlySet<string> = new Set([
  "pStyle",
  "rStyle",
  "tblStyle",
]);

/** Marks of the source's own bookmarks and comments, left out of a copy. */
const DROPPED: ReadonlySet<string> = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
]);

/** Paragraph ids a copy leaves out: the output's own parts hold theirs. */
const DROPPED_ATTRIBUTES: ReadonlySet<string> = new Set(["paraId", "textId"]);

/** The pictures and embedded objects a part refers to, read from the
 *  source package. */
const picturesOf = async (
  part: SourcePart,
  target: CopyTarget,
): Promise<Map<string, SourceMedia>> => {
  const media = new Map<string, SourceMedia>();
  for (const relationship of part.relationships.values()) {
    if (
      !EMBEDDED_RELATIONSHIP_TYPES.has(relationship.type) ||
      relationship.targetMode === "External"
    ) {
      continue;
    }
    const path = target.resolveSourceTarget(part.path, relationship.target);
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const bytes = await target.readSourceBytes(path);
    if (bytes !== null) {
      media.set(path, {
        bytes,
        extension: extensionOf(path),
        relationshipType: relationship.type,
      });
    }
  }
  return media;
};

/** One block of a source part re-homed for an output part, or null when
 *  it refers to anything other than a picture, an embedded object or an
 *  external address. */
const rehomedBlock = (
  block: XmlElement,
  part: SourcePart,
  media: ReadonlyMap<string, SourceMedia>,
  target: CopyTarget,
): string | null => {
  // Set by a relationship the output part cannot re-make.
  const state = { refused: false };
  const mapRelationship = (id: string): string => {
    const relationship = part.relationships.get(id);
    if (relationship === undefined) {
      state.refused = true;
      return id;
    }
    if (relationship.targetMode === "External") {
      return target.relationships.add({
        type: relationship.type,
        target: relationship.target,
        external: true,
      });
    }
    const path = target.resolveSourceTarget(part.path, relationship.target);
    const picture = media.get(path);
    if (
      !EMBEDDED_RELATIONSHIP_TYPES.has(relationship.type) ||
      picture === undefined
    ) {
      state.refused = true;
      return id;
    }
    return target.relationships.add({
      type: relationship.type,
      target: target.addMedia(path, picture),
      external: false,
    });
  };
  const attributesOf = (
    node: XmlElement,
  ): Record<string, string | number | undefined> => {
    const attributes: Record<string, string | number | undefined> = {};
    for (const [name, value] of Object.entries(node.attributes ?? {})) {
      const colon = name.indexOf(":");
      if (
        value === undefined ||
        (colon !== -1 && DROPPED_ATTRIBUTES.has(name.slice(colon + 1)))
      ) {
        continue;
      }
      const uri =
        colon === -1 || name.startsWith("xmlns:")
          ? undefined
          : resolveNamespaceUri(node.namespaceScope, name.slice(0, colon));
      attributes[name] = OFFICE_RELATIONSHIP_NAMESPACE_URIS.has(uri ?? "")
        ? mapRelationship(String(value))
        : value;
    }
    if (
      getLocalName(node.name ?? "") === "docPr" &&
      attributes["id"] !== undefined
    ) {
      attributes["id"] = String(target.nextDrawingId());
    }
    return attributes;
  };
  const rehome = (node: XmlElement): XmlElement | null => {
    if (node.type !== "element") {
      return node;
    }
    const local = getLocalName(node.name ?? "");
    // A list or a style the output does not define is the source's: the
    // text keeps its own formatting and no list.
    if (
      DROPPED.has(local) ||
      local === "numPr" ||
      (STYLE_REFERENCES.has(local) &&
        !target.hasStyle(wAttr(node, "val") ?? ""))
    ) {
      return null;
    }
    return cloneElement(node, {
      attributes: attributesOf(node),
      ...(node.elements === undefined
        ? {}
        : {
            elements: node.elements.flatMap((child) => {
              const kept = rehome(child);
              return kept === null ? [] : [kept];
            }),
          }),
    });
  };
  const rehomed = rehome(block);
  if (state.refused || rehomed === null) {
    return null;
  }
  // The block brings the namespaces its part declares at the root.
  const declared: Record<string, string | number | undefined> = {};
  for (const [name, value] of Object.entries(part.root.attributes ?? {})) {
    if (name.startsWith("xmlns") || name.endsWith(":Ignorable")) {
      declared[name] = value;
    }
  }
  return captureVerbatimXml(
    cloneElement(rehomed, {
      attributes: { ...declared, ...rehomed.attributes },
    }),
  );
};

/**
 * The carried blocks of a source header or footer, re-homed for an output
 * part: relationships re-made in that part's relationships (pictures with
 * their parts), drawing ids renumbered, bookmarks and comment marks left
 * out. A block whose picture cannot be read is counted as not carried.
 */
export const copyPartContent = async (
  part: SourcePart,
  target: CopyTarget,
): Promise<{ markup: string; unsupported: number }> => {
  const media = await picturesOf(part, target);
  let unsupported = 0;
  const blocks: string[] = [];
  for (const block of part.blocks) {
    const xml = rehomedBlock(block, part, media, target);
    if (xml === null) {
      unsupported += 1;
    } else {
      blocks.push(xml);
    }
  }
  return { markup: blocks.join(""), unsupported };
};

const partRootOf = async (
  document: AnalyzedDocument,
  path: string | null,
): Promise<Result<XmlElement | null, TemplatePlanError>> => {
  const xml = await readPartText(document.pkg, path);
  if (Result.isError(xml)) {
    return xml;
  }
  return Result.ok(
    path === null || xml.value === null
      ? null
      : (readRawPart(path, xml.value)?.root ?? null),
  );
};

/**
 * Checks the output's headers and footers against what the source's give
 * them: each that takes a source part's content reads as the template's
 * (nothing where a new part holds it) followed by that content.
 */
export const verifyCarriedHeaders = async ({
  source,
  template,
  output,
  starts = [],
}: {
  source: AnalyzedDocument;
  template: AnalyzedDocument;
  output: AnalyzedDocument;
  starts?: readonly number[];
}): Promise<
  Result<
    { report: TemplatePlanSourceHeaders; mismatched: string[] },
    TemplatePlanError
  >
> => {
  const planned = await plannedHeaderWrites(source, template, starts);
  if (Result.isError(planned)) {
    return planned;
  }
  const outputSections = sectionsOf(output.body);
  const kept: TemplatePlanSourceHeaders["kept"] = [];
  const mismatched: string[] = [];
  for (const write of planned.value.writes) {
    // Every output section names its headers itself: the first section's
    // show the content all of them share, a later one's its own.
    const id = (
      outputSections.length === 0
        ? outputReferences(output.body, write.family)
        : sectionReferences(outputSections, write.section ?? 0, write.family)
    ).get(write.kind);
    const path =
      id === undefined ? null : familyPathOf(output, write.family, id);
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const outputRoot = await partRootOf(output, path);
    if (Result.isError(outputRoot)) {
      return outputRoot;
    }
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const templateRoot = await partRootOf(template, write.templatePath);
    if (Result.isError(templateRoot)) {
      return templateRoot;
    }
    const expected = [
      ...(templateRoot.value === null
        ? []
        : partParagraphTexts(templateRoot.value)),
      ...write.source.texts,
    ];
    const actual =
      outputRoot.value === null ? null : partParagraphTexts(outputRoot.value);
    if (
      path === null ||
      actual === null ||
      actual.length !== expected.length ||
      actual.some((text, index) => text !== expected[index])
    ) {
      mismatched.push(`${write.kind} ${write.family}`);
      continue;
    }
    kept.push({
      family: write.family,
      kind: write.kind,
      part: path,
      paragraphs: write.source.texts.length,
      created: write.templatePath === null || write.section !== undefined,
      ...(write.section === undefined ? {} : { section: write.section }),
    });
  }
  return Result.ok({
    report: { kept, leftOut: planned.value.leftOut },
    mismatched,
  });
};
