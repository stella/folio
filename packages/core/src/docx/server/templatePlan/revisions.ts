/**
 * Tracked changes resolved to one view.
 *
 * A document with tracked changes is two texts at once: what it reads with
 * every change accepted, and what it reads with every change rejected. A
 * plan is made for one of them. This module rewrites the story parts of a
 * package into the chosen view: no insertion, deletion, move or recorded
 * formatting change is left, and nothing else of the parts changes.
 *
 * - Accepted: inserted content stays and loses its wrapper, deleted content
 *   goes. Rejected: the reverse, and deleted text is ordinary text again.
 * - A paragraph mark that the view removes (deleted and accepted, inserted
 *   and rejected) joins its paragraph to the next one of the same container,
 *   whose mark, and with it whose properties, the joined paragraph keeps.
 * - A table row that the view removes goes whole; so does a cell.
 * - A recorded formatting change is dropped when accepted. Rejected, the
 *   properties it recorded come back.
 */
import { Result } from "better-result";

import { ZipArchive } from "@stll/docx-core/zip";

import type { XmlElement } from "../../xmlParser";
import { loadPackage } from "./package";
import { TemplatePlanError } from "./types";
import type { DocxTrackedChanges, TemplatePlanTrackedView } from "./types";
import {
  elementChildren,
  innerRangeOf,
  rangeOf,
  readRawPart,
  spliceXml,
  wAttr,
  wChild,
  wName,
} from "./xml";

const INSERTED = new Set(["ins", "moveTo"]);
const DELETED = new Set(["del", "moveFrom"]);
const MOVE_MARKS = new Set([
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);
/** A recorded change of properties, by the properties element it sits in. */
const PROPERTY_CHANGES: ReadonlyMap<string, string> = new Map([
  ["rPrChange", "rPr"],
  ["pPrChange", "pPr"],
  ["sectPrChange", "sectPr"],
  ["tblPrChange", "tblPr"],
  ["tblGridChange", "tblGrid"],
  ["tblPrExChange", "tblPrEx"],
  ["trPrChange", "trPr"],
  ["tcPrChange", "tcPr"],
]);
const CELL_MARKS = new Set(["cellIns", "cellDel", "cellMerge"]);
/** What a paragraph's properties keep of their own when the recorded ones
 *  come back: the mark's run properties and a section. */
const PARAGRAPH_OWN = new Set(["rPr", "sectPr"]);
/** What a section keeps of its own: the headers and footers it names. */
const SECTION_OWN = new Set(["headerReference", "footerReference"]);

const OWN_PROPERTIES: ReadonlyMap<string, ReadonlySet<string>> = new Map([
  ["pPr", PARAGRAPH_OWN],
  ["sectPr", SECTION_OWN],
]);
/** Marks of a range that outlive content a view removes around them: a
 *  bookmark or a comment that begins inside a deletion ends outside it. */
const RANGE_MARKS = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "permStart",
  "permEnd",
]);

/** Every element name this module acts on, to skip a subtree without one. */
const REVISION_NAMES = [
  ...INSERTED,
  ...DELETED,
  ...MOVE_MARKS,
  ...PROPERTY_CHANGES.keys(),
  ...CELL_MARKS,
  "numberingChange",
  "delText",
  "delInstrText",
];

type Tally = {
  insertions: number;
  deletions: number;
  moves: number;
  paragraphMarks: number;
  tableRows: number;
  formattingChanges: number;
  authors: Set<string>;
  first: string | null;
  last: string | null;
};

const newTally = (): Tally => ({
  insertions: 0,
  deletions: 0,
  moves: 0,
  paragraphMarks: 0,
  tableRows: 0,
  formattingChanges: 0,
  authors: new Set(),
  first: null,
  last: null,
});

type Walk = {
  xml: string;
  view: TemplatePlanTrackedView;
  tally: Tally;
  /** False once something could not be read where it stands in the part. */
  sound: boolean;
};

/**
 * What a part tracks, whatever view is taken of it: every insertion,
 * deletion, move and recorded formatting change, with who made it and
 * when. A change inside content the view removes is counted all the same.
 */
const tallyPart = (root: XmlElement, tally: Tally): void => {
  const stack = [root];
  for (
    let element = stack.pop();
    element !== undefined;
    element = stack.pop()
  ) {
    const name = wName(element);
    let counted = true;
    if (name === "ins") {
      tally.insertions += 1;
    } else if (name === "del") {
      tally.deletions += 1;
    } else if (name === "moveFrom" || name === "moveTo") {
      tally.moves += 1;
    } else if (
      name !== null &&
      (PROPERTY_CHANGES.has(name) ||
        CELL_MARKS.has(name) ||
        name === "numberingChange")
    ) {
      tally.formattingChanges += 1;
    } else {
      counted = false;
    }
    if (counted) {
      const author = wAttr(element, "author");
      const date = wAttr(element, "date");
      if (author !== null && author !== "") {
        tally.authors.add(author);
      }
      if (date !== null && date !== "") {
        tally.first =
          tally.first === null || date < tally.first ? date : tally.first;
        tally.last =
          tally.last === null || date > tally.last ? date : tally.last;
      }
    }
    stack.push(...elementChildren(element));
  }
};

const mayHoldRevision = (text: string): boolean =>
  REVISION_NAMES.some((name) => text.includes(`:${name}`));

/** The element's source text; the walk is unsound without its range. */
const sourceOf = (walk: Walk, element: XmlElement): string => {
  const range = rangeOf(element);
  if (range === null) {
    walk.sound = false;
    return "";
  }
  return walk.xml.slice(range.start, range.end);
};

/** The range marks inside content a view removes, as they stand. */
const rangeMarksIn = (walk: Walk, element: XmlElement): string => {
  let marks = "";
  const stack = [...elementChildren(element)].toReversed();
  for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
    if (RANGE_MARKS.has(wName(next) ?? "")) {
      marks += sourceOf(walk, next);
    } else {
      stack.push(...elementChildren(next).toReversed());
    }
  }
  return marks;
};

/** Whether the chosen view keeps content marked `name`. */
const keeps = (walk: Walk, name: string): boolean =>
  INSERTED.has(name) ? walk.view === "accept" : walk.view === "reject";

/** Deleted text read as text again: the names of its elements. */
const asLiveText = (xml: string): string =>
  xml
    .replaceAll(/<(\/?)([A-Za-z0-9]+:)delText\b/gu, "<$1$2t")
    .replaceAll(/<(\/?)([A-Za-z0-9]+:)delInstrText\b/gu, "<$1$2instrText");

/** Whether the view removes the paragraph's mark. */
const losesMark = (walk: Walk, paragraph: XmlElement): boolean => {
  const mark = wChild(wChild(paragraph, "pPr"), "rPr");
  const removed =
    walk.view === "accept" ? ["del", "moveFrom"] : ["ins", "moveTo"];
  return removed.some((name) => wChild(mark, name) !== null);
};

/** Whether the view removes the row: the marker of `w:trPr`. */
const losesRow = (walk: Walk, row: XmlElement): boolean => {
  const properties = wChild(row, "trPr");
  return wChild(properties, walk.view === "accept" ? "del" : "ins") !== null;
};

const losesCell = (walk: Walk, cell: XmlElement): boolean =>
  wChild(
    wChild(cell, "tcPr"),
    walk.view === "accept" ? "cellDel" : "cellIns",
  ) !== null;

type Piece = { xml: string; paragraph: XmlElement | null; carried: boolean };

/**
 * The children of `container` in the chosen view, with the text between
 * them as it stands. A paragraph whose mark goes hands its content to the
 * next paragraph of the container.
 */
const resolveChildren = (walk: Walk, container: XmlElement): string => {
  const inner = innerRangeOf(walk.xml, container);
  if (inner === null) {
    return "";
  }
  const pieces: (string | Piece)[] = [];
  let cursor = inner.start;
  const visit = (element: XmlElement): void => {
    const range = rangeOf(element);
    if (range === null) {
      walk.sound = false;
      return;
    }
    pieces.push(walk.xml.slice(cursor, range.start));
    cursor = range.end;
    const name = wName(element);
    if (name !== null && (INSERTED.has(name) || DELETED.has(name))) {
      if (keeps(walk, name)) {
        const kept = resolveChildren(walk, element);
        pieces.push(DELETED.has(name) ? asLiveText(kept) : kept);
      } else {
        pieces.push(rangeMarksIn(walk, element));
      }
      return;
    }
    if (name !== null && MOVE_MARKS.has(name)) {
      return;
    }
    if (name === "numberingChange" || (name !== null && CELL_MARKS.has(name))) {
      return;
    }
    if (name !== null && PROPERTY_CHANGES.has(name)) {
      // Dropped here; a rejected one is restored by the element it sits in.
      return;
    }
    if (name === "tr" && losesRow(walk, element)) {
      walk.tally.tableRows += 1;
      pieces.push(rangeMarksIn(walk, element));
      return;
    }
    if (name === "tc" && losesCell(walk, element)) {
      pieces.push(rangeMarksIn(walk, element));
      return;
    }
    if (name === "p") {
      pieces.push({
        xml: resolveElement(walk, element),
        paragraph: element,
        carried: losesMark(walk, element),
      });
      return;
    }
    pieces.push(resolveElement(walk, element));
  };
  for (const child of elementChildren(container)) {
    visit(child);
  }
  pieces.push(walk.xml.slice(cursor, inner.end));
  return joinParagraphs(walk, pieces);
};

/** The content of a resolved paragraph without its properties. */
const contentOf = (paragraphXml: string): string => {
  const open = paragraphXml.indexOf(">") + 1;
  const close = paragraphXml.lastIndexOf("</");
  if (paragraphXml.endsWith("/>") || close < open) {
    return "";
  }
  const body = paragraphXml.slice(open, close);
  const properties = /^\s*<([A-Za-z0-9]+:)?pPr\b/u.exec(body);
  if (properties === null) {
    return body;
  }
  if (/^\s*<([A-Za-z0-9]+:)?pPr\b[^>]*\/>/u.test(body)) {
    return body.slice(body.indexOf("/>") + 2);
  }
  const end = /<\/([A-Za-z0-9]+:)?pPr>/u.exec(body);
  return end === null ? body : body.slice(end.index + end[0].length);
};

/** `content` put in front of what a resolved paragraph holds. */
const withLeadingContent = (paragraphXml: string, content: string): string => {
  if (content === "") {
    return paragraphXml;
  }
  if (paragraphXml.endsWith("/>")) {
    const name = /^<([^\s/>]+)/u.exec(paragraphXml)?.[1] ?? "w:p";
    return `${paragraphXml.slice(0, -2)}>${content}</${name}>`;
  }
  const open = paragraphXml.indexOf(">") + 1;
  const body = paragraphXml.slice(open);
  const selfClosed = /^\s*<([A-Za-z0-9]+:)?pPr\b[^>]*\/>/u.exec(body);
  if (selfClosed !== null) {
    const at = open + selfClosed[0].length;
    return `${paragraphXml.slice(0, at)}${content}${paragraphXml.slice(at)}`;
  }
  if (/^\s*<([A-Za-z0-9]+:)?pPr\b/u.test(body)) {
    const end = /<\/([A-Za-z0-9]+:)?pPr>/u.exec(body);
    if (end !== null) {
      const at = open + end.index + end[0].length;
      return `${paragraphXml.slice(0, at)}${content}${paragraphXml.slice(at)}`;
    }
  }
  return `${paragraphXml.slice(0, open)}${content}${body}`;
};

const joinParagraphs = (walk: Walk, pieces: (string | Piece)[]): string => {
  let out = "";
  /** Content of paragraphs whose mark went, waiting for the next one. */
  let carried = "";
  /** The last such paragraph, kept whole when no paragraph follows. */
  let held: string | null = null;
  const flush = (): void => {
    if (held !== null) {
      // No paragraph to join: the mark stays, as it must close a cell, a
      // note or the text in front of a table.
      out += withLeadingContent(held, carried);
      carried = "";
      held = null;
    }
  };
  for (const piece of pieces) {
    if (typeof piece === "string") {
      // Text between blocks is white space; it follows whatever came last.
      out += piece;
      continue;
    }
    if (piece.paragraph === null) {
      flush();
      out += piece.xml;
      continue;
    }
    if (piece.carried) {
      walk.tally.paragraphMarks += 1;
      if (held !== null) {
        carried += contentOf(held);
      }
      held = piece.xml;
      continue;
    }
    if (held !== null) {
      carried += contentOf(held);
      held = null;
    }
    out += withLeadingContent(piece.xml, carried);
    carried = "";
  }
  flush();
  return out;
};

/** The properties a recorded change gives back, inside `properties`. */
const restoredProperties = (
  walk: Walk,
  properties: XmlElement,
  name: string,
): string | null => {
  const changeName = [...PROPERTY_CHANGES].find(
    ([, owner]) => owner === name,
  )?.[0];
  const change =
    changeName === undefined ? null : wChild(properties, changeName);
  if (change === null || walk.view !== "reject") {
    return null;
  }
  const recorded = wChild(change, name);
  const outer = rangeOf(properties);
  const inner = innerRangeOf(walk.xml, properties);
  if (outer === null || inner === null) {
    return null;
  }
  const recordedInner =
    recorded === null ? null : innerRangeOf(walk.xml, recorded);
  const old =
    recordedInner === null
      ? ""
      : walk.xml.slice(recordedInner.start, recordedInner.end);
  const own = OWN_PROPERTIES.get(name) ?? null;
  const kept =
    own === null
      ? ""
      : elementChildren(properties)
          .filter((child) => own.has(wName(child) ?? ""))
          .map((child) => resolveElement(walk, child))
          .join("");
  const head = walk.xml.slice(outer.start, inner.start);
  const tail = walk.xml.slice(inner.end, outer.end);
  // A section names its headers first; a paragraph's own properties close it.
  return name === "sectPr"
    ? `${head}${kept}${old}${tail}`
    : `${head}${old}${kept}${tail}`;
};

/** `element` in the chosen view; its own text when it holds no revision. */
const resolveElement = (walk: Walk, element: XmlElement): string => {
  const source = sourceOf(walk, element);
  if (!mayHoldRevision(source)) {
    return source;
  }
  const name = wName(element);
  if (name !== null && [...PROPERTY_CHANGES.values()].includes(name)) {
    const restored = restoredProperties(walk, element, name);
    if (restored !== null) {
      return restored;
    }
  }
  const outer = rangeOf(element);
  const inner = innerRangeOf(walk.xml, element);
  if (outer === null) {
    walk.sound = false;
    return source;
  }
  if (inner === null) {
    return source;
  }
  const content = resolveChildren(walk, element);
  const head = walk.xml.slice(outer.start, inner.start);
  const tail = walk.xml.slice(inner.end, outer.end);
  return `${head}${content}${tail}`;
};

const summaryOf = (tally: Tally): DocxTrackedChanges => ({
  insertions: tally.insertions,
  deletions: tally.deletions,
  moves: tally.moves,
  paragraphMarks: tally.paragraphMarks,
  tableRows: tally.tableRows,
  formattingChanges: tally.formattingChanges,
  authors: [...tally.authors].toSorted(),
  first: tally.first,
  last: tally.last,
});

/** Changes a plan that keeps tracked changes cannot carry yet, counted. */
export type KeptRevisionObstacles = {
  paragraphMarks: number;
  tableRows: number;
  tableCells: number;
  moves: number;
};

const REVISION_WRAPPERS = new Set(["ins", "del", "moveFrom", "moveTo"]);

const countObstacles = (
  root: XmlElement,
  into: KeptRevisionObstacles,
): void => {
  const visit = (
    element: XmlElement,
    parent: string | null,
    grandparent: string | null,
  ): void => {
    const name = wName(element);
    if (name !== null && REVISION_WRAPPERS.has(name)) {
      if (parent === "rPr" && grandparent === "pPr") {
        into.paragraphMarks += 1;
      } else if (parent === "trPr") {
        into.tableRows += 1;
      } else if (name === "moveFrom" || name === "moveTo") {
        into.moves += 1;
      }
    } else if (name !== null && CELL_MARKS.has(name)) {
      into.tableCells += 1;
    }
    for (const child of elementChildren(element)) {
      visit(child, name, parent);
    }
  };
  visit(root, null, null);
};

/**
 * What a document tracks, and which of it a plan that keeps tracked changes
 * cannot carry: a change that makes the paragraphs of the accepted and the
 * rejected view differ.
 */
export const keptRevisionObstacles = async (
  docx: Uint8Array,
): Promise<
  Result<
    { tracked: DocxTrackedChanges; obstacles: KeptRevisionObstacles },
    TemplatePlanError
  >
> => {
  const loaded = await loadPackage(docx);
  if (loaded.isErr()) {
    return Result.err(loaded.error);
  }
  const pkg = loaded.value;
  const tally = newTally();
  const obstacles: KeptRevisionObstacles = {
    paragraphMarks: 0,
    tableRows: 0,
    tableCells: 0,
    moves: 0,
  };
  const paths = [
    pkg.documentPath,
    pkg.partPath("footnotes"),
    pkg.partPath("endnotes"),
  ].filter((path): path is string => path !== null);
  for (const path of paths) {
    const xml =
      path === pkg.documentPath ? pkg.documentXml : await pkg.readText(path);
    if (xml === null || !mayHoldRevision(xml)) {
      continue;
    }
    const part = readRawPart(path, xml);
    if (part !== null) {
      tallyPart(part.root, tally);
      countObstacles(part.root, obstacles);
    }
  }
  return Result.ok({
    tracked: {
      ...summaryOf(tally),
      paragraphMarks: obstacles.paragraphMarks,
      tableRows: obstacles.tableRows,
    },
    obstacles,
  });
};

/** One part in the chosen view, or null when it cannot be read whole. */
export const resolvePartRevisions = (
  path: string,
  xml: string,
  view: TemplatePlanTrackedView,
): { xml: string; tally: Tally } | null => {
  const part = readRawPart(path, xml);
  if (part === null) {
    return null;
  }
  const walk: Walk = { xml, view, tally: newTally(), sound: true };
  const outer = rangeOf(part.root);
  if (outer === null) {
    return null;
  }
  tallyPart(part.root, walk.tally);
  const resolved = resolveElement(walk, part.root);
  if (!walk.sound) {
    return null;
  }
  const rewritten = spliceXml(xml, [{ ...outer, newXml: resolved }]);
  return rewritten === null ? null : { xml: rewritten, tally: walk.tally };
};

const mergeTallies = (tallies: readonly Tally[]): Tally => {
  const total = newTally();
  for (const tally of tallies) {
    total.insertions += tally.insertions;
    total.deletions += tally.deletions;
    total.moves += tally.moves;
    total.paragraphMarks += tally.paragraphMarks;
    total.tableRows += tally.tableRows;
    total.formattingChanges += tally.formattingChanges;
    for (const author of tally.authors) {
      total.authors.add(author);
    }
    for (const date of [tally.first, tally.last]) {
      if (date !== null) {
        total.first =
          total.first === null || date < total.first ? date : total.first;
        total.last =
          total.last === null || date > total.last ? date : total.last;
      }
    }
  }
  return total;
};

/**
 * A package with the tracked changes of its main document and its notes
 * resolved to `view`, and what was resolved. The bytes are the source's own
 * when it tracks nothing.
 */
export const resolveTrackedChanges = async (
  docx: Uint8Array,
  view: TemplatePlanTrackedView,
): Promise<
  Result<{ docx: Uint8Array; resolved: DocxTrackedChanges }, TemplatePlanError>
> => {
  const loaded = await loadPackage(docx);
  if (loaded.isErr()) {
    return Result.err(loaded.error);
  }
  const pkg = loaded.value;
  const paths = [
    pkg.documentPath,
    pkg.partPath("footnotes"),
    pkg.partPath("endnotes"),
  ].filter((path): path is string => path !== null);
  const rewritten = new Map<string, string>();
  const tallies: Tally[] = [];
  for (const path of paths) {
    const xml =
      path === pkg.documentPath ? pkg.documentXml : await pkg.readText(path);
    if (xml === null || !mayHoldRevision(xml)) {
      continue;
    }
    const resolved = resolvePartRevisions(path, xml, view);
    if (resolved === null) {
      return Result.err(
        new TemplatePlanError({
          code: "xml-unreadable",
          message: `The tracked changes of ${path} could not be resolved.`,
        }),
      );
    }
    tallies.push(resolved.tally);
    if (resolved.xml !== xml) {
      rewritten.set(path, resolved.xml);
    }
  }
  const resolved = summaryOf(mergeTallies(tallies));
  if (rewritten.size === 0) {
    return Result.ok({ docx, resolved });
  }
  return await Result.tryPromise({
    try: async () => {
      const zip = await ZipArchive.read(docx);
      for (const [path, content] of rewritten) {
        const current = zip.entry(path);
        zip.set(
          path,
          content,
          current === undefined ? {} : { modified: current.modified },
        );
      }
      return { docx: await zip.generate(), resolved };
    },
    catch: (cause) =>
      new TemplatePlanError({
        code: "archive-unreadable",
        message:
          "The package could not be rewritten without its tracked changes.",
        cause,
      }),
  });
};
