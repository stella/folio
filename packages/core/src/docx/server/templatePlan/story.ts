/**
 * One reader for the inline content of a paragraph, shared by the analysis,
 * the run transfer and the verification, so all three agree on what a block's
 * text is and where an offset into it lands.
 */
import type { ParagraphNumberingOverride } from "@stll/docx-core/model";

import { parseFieldType } from "../../fieldParser";
import { readParagraphNumbering } from "../../numberingReference";
import { getLocalName, type XmlElement } from "../../xmlParser";
import {
  elementChildren,
  isW,
  MARKUP_COMPATIBILITY_NAMESPACE,
  MATH_NAMESPACE,
  relationshipAttr,
  wAttr,
  wChild,
  wChildVal,
  wName,
} from "./xml";

/** The mark a footnote or endnote reference leaves in a block's text. */
export const NOTE_REFERENCE_MARK = "￼";

export type AtomKind =
  | "text"
  | "tab"
  | "lineBreak"
  | "pageBreak"
  | "noBreakHyphen"
  | "softHyphen"
  | "symbol"
  | "footnoteReference"
  | "endnoteReference"
  | "object";

export type AtomLink =
  | { kind: "relationship"; relationshipId: string }
  | { kind: "url"; url: string }
  | { kind: "anchor"; anchor: string };

/** A piece of inline content with the text it contributes to its block. */
export type InlineAtom = {
  kind: AtomKind;
  text: string;
  /** `w:rPr` of the run the atom came from. */
  runProperties: XmlElement | null;
  link: AtomLink | null;
  /** The run child itself: the note reference, symbol, break or object. */
  element: XmlElement | null;
  /** Fields whose result the atom is part of, outermost first. */
  fields: readonly number[];
};

export type StoryField = {
  /** Position among the fields of the story, in begin order. */
  index: number;
  instruction: string;
  kind: string;
  result: string;
  /** Text offset in the block the result starts at; -1 before a result starts. */
  resultStart: number;
  /** Paragraph ordinal within the story the field begins in. */
  paragraphOrdinal: number;
  /** Paragraph ordinal the field ends in; the begin's until the end is seen. */
  endParagraphOrdinal: number;
  /** Whether both ends were seen. */
  closed: boolean;
};

type OpenField = {
  field: StoryField;
  phase: "code" | "result";
};

/** Field and revision state that outlives one paragraph of a story. */
export type StoryState = {
  fields: StoryField[];
  open: OpenField[];
  paragraphOrdinal: number;
  comments: number;
  trackedChanges: number;
  formattingRevisions: number;
  /** Complex field ends met with no field open. */
  strayFieldEnds: number;
};

export const createStoryState = (): StoryState => ({
  fields: [],
  open: [],
  paragraphOrdinal: 0,
  comments: 0,
  trackedChanges: 0,
  formattingRevisions: 0,
  strayFieldEnds: 0,
});

export type ParagraphBookmark = { name: string; id: string };

export type ParagraphRead = {
  element: XmlElement;
  ordinal: number;
  styleId: string | null;
  numPr: ParagraphNumberingOverride | undefined;
  paraId: string | null;
  outlineLevel: number | null;
  atoms: InlineAtom[];
  text: string;
  /** Indices into `StoryState.fields` of the fields beginning here. */
  fieldIndices: number[];
  bookmarks: ParagraphBookmark[];
  bookmarkEndIds: string[];
  inlineObjects: number;
  hasSectionBreak: boolean;
  /** Whether the paragraph mark is a tracked insertion or deletion. */
  hasTrackedMark: boolean;
};

const REVISION_INCLUDED = new Set(["ins", "moveTo"]);
const REVISION_EXCLUDED = new Set(["del", "moveFrom"]);
const TRANSPARENT_WRAPPERS = new Set(["smartTag", "customXml", "dir", "bdo"]);
const PROPERTY_REVISIONS = new Set([
  "rPrChange",
  "pPrChange",
  "sectPrChange",
  "tblPrChange",
  "trPrChange",
  "tcPrChange",
  "numberingChange",
]);
const OBJECT_RUN_CHILDREN = new Set(["drawing", "pict", "object"]);

const symbolText = (element: XmlElement): string => {
  const code = Number.parseInt(wAttr(element, "char") ?? "", 16);
  return Number.isInteger(code) && code > 0 && code <= 0x10ffff
    ? String.fromCodePoint(code)
    : "�";
};

const textOf = (element: XmlElement): string =>
  (element.elements ?? [])
    .map((child) => (child.type === "text" ? String(child.text ?? "") : ""))
    .join("");

const BODY_TEXT_OUTLINE_LEVEL = 9;

/** `w:outlineLvl` as a heading level; the reserved body-text value is none. */
const headingLevelOf = (value: string | null): number | null => {
  const level = value === null ? Number.NaN : Number.parseInt(value, 10);
  return Number.isInteger(level) &&
    level >= 0 &&
    level < BODY_TEXT_OUTLINE_LEVEL
    ? level
    : null;
};

const paraIdOf = (element: XmlElement): string | null => {
  for (const [name, value] of Object.entries(element.attributes ?? {})) {
    if (value !== undefined && getLocalName(name) === "paraId") {
      return String(value);
    }
  }
  return null;
};

const fieldKindOf = (instruction: string): string => {
  const known = parseFieldType(instruction);
  if (known !== "UNKNOWN") {
    return known;
  }
  return (
    /^\s*\\?(?<name>[A-Za-z][A-Za-z0-9]*)/u
      .exec(instruction)
      ?.groups?.["name"]?.toUpperCase() ?? "UNKNOWN"
  );
};

type ParagraphWalk = {
  state: StoryState;
  read: ParagraphRead;
  textLength: number;
};

const inCode = (state: StoryState): boolean =>
  state.open.some(({ phase }) => phase === "code");

const pushAtom = (
  walk: ParagraphWalk,
  atom: Omit<InlineAtom, "fields">,
): void => {
  const { state, read } = walk;
  if (inCode(state)) {
    return;
  }
  const fields = state.open.map(({ field }) => field.index);
  for (const { field } of state.open) {
    field.result += atom.text;
  }
  read.atoms.push({ ...atom, fields });
  walk.textLength += atom.text.length;
};

const beginField = (walk: ParagraphWalk, instruction: string): StoryField => {
  const { state, read } = walk;
  const field: StoryField = {
    index: state.fields.length,
    instruction,
    kind: fieldKindOf(instruction),
    result: "",
    resultStart: -1,
    paragraphOrdinal: read.ordinal,
    endParagraphOrdinal: read.ordinal,
    closed: false,
  };
  state.fields.push(field);
  read.fieldIndices.push(field.index);
  return field;
};

type RunContext = { link: AtomLink | null };

const readRun = (
  run: XmlElement,
  walk: ParagraphWalk,
  { link }: RunContext,
): void => {
  const { state, read } = walk;
  const runProperties = wChild(run, "rPr");
  if (wChild(runProperties, "rPrChange") !== null) {
    state.formattingRevisions += 1;
  }
  const atom = (
    kind: AtomKind,
    text: string,
    element: XmlElement | null,
  ): void => {
    pushAtom(walk, { kind, text, runProperties, link, element });
  };
  for (const child of elementChildren(run)) {
    const name = wName(child);
    if (name === null) {
      if (
        child.namespaceUri === MARKUP_COMPATIBILITY_NAMESPACE &&
        getLocalName(child.name) === "AlternateContent" &&
        !inCode(state)
      ) {
        read.inlineObjects += 1;
        atom("object", "", child);
      }
      continue;
    }
    switch (name) {
      case "t":
        atom("text", textOf(child), null);
        break;
      case "tab":
      case "ptab":
        atom("tab", "\t", child);
        break;
      case "cr":
        atom("lineBreak", "\n", child);
        break;
      case "br": {
        const type = wAttr(child, "type");
        if (type === "page" || type === "column") {
          atom("pageBreak", "", child);
        } else {
          atom("lineBreak", "\n", child);
        }
        break;
      }
      case "noBreakHyphen":
        atom("noBreakHyphen", "‑", child);
        break;
      case "softHyphen":
        atom("softHyphen", "­", child);
        break;
      case "sym":
        atom("symbol", symbolText(child), child);
        break;
      case "footnoteReference":
        atom("footnoteReference", NOTE_REFERENCE_MARK, child);
        break;
      case "endnoteReference":
        atom("endnoteReference", NOTE_REFERENCE_MARK, child);
        break;
      case "commentReference":
        state.comments += 1;
        break;
      case "fldChar": {
        const type = wAttr(child, "fldCharType");
        if (type === "begin") {
          state.open.push({ field: beginField(walk, ""), phase: "code" });
        } else if (type === "separate") {
          const top = state.open.at(-1);
          if (top !== undefined) {
            top.phase = "result";
            top.field.kind = fieldKindOf(top.field.instruction);
            top.field.resultStart =
              top.field.paragraphOrdinal === read.ordinal ? walk.textLength : 0;
          }
        } else if (type === "end") {
          const top = state.open.pop();
          if (top === undefined) {
            state.strayFieldEnds += 1;
          } else {
            top.field.kind = fieldKindOf(top.field.instruction);
            top.field.endParagraphOrdinal = read.ordinal;
            top.field.closed = true;
          }
        }
        break;
      }
      case "instrText": {
        const top = state.open.at(-1);
        if (top !== undefined && top.phase === "code") {
          top.field.instruction += textOf(child);
        }
        break;
      }
      default:
        if (OBJECT_RUN_CHILDREN.has(name) && !inCode(state)) {
          read.inlineObjects += 1;
          atom("object", "", child);
        }
        break;
    }
  }
};

const readInline = (
  container: XmlElement,
  walk: ParagraphWalk,
  context: RunContext,
): void => {
  const { state, read } = walk;
  for (const child of elementChildren(container)) {
    const name = wName(child);
    if (name === null) {
      if (
        child.namespaceUri === MATH_NAMESPACE &&
        !inCode(state) &&
        (getLocalName(child.name) === "oMath" ||
          getLocalName(child.name) === "oMathPara")
      ) {
        read.inlineObjects += 1;
        pushAtom(walk, {
          kind: "object",
          text: "",
          runProperties: null,
          link: context.link,
          element: child,
        });
      }
      continue;
    }
    if (name === "r") {
      readRun(child, walk, context);
    } else if (name === "hyperlink") {
      const relationshipId = relationshipAttr(child, "id");
      const anchor = wAttr(child, "anchor");
      let link: AtomLink | null = context.link;
      if (relationshipId !== null) {
        link = { kind: "relationship", relationshipId };
      } else if (anchor !== null) {
        link = { kind: "anchor", anchor };
      }
      readInline(child, walk, { link });
    } else if (name === "fldSimple") {
      const field = beginField(walk, wAttr(child, "instr") ?? "");
      field.resultStart = walk.textLength;
      state.open.push({ field, phase: "result" });
      readInline(child, walk, context);
      state.open.pop();
      field.closed = true;
    } else if (REVISION_INCLUDED.has(name)) {
      state.trackedChanges += 1;
      readInline(child, walk, context);
    } else if (REVISION_EXCLUDED.has(name)) {
      state.trackedChanges += 1;
    } else if (name === "sdt") {
      const content = wChild(child, "sdtContent");
      if (content !== null) {
        readInline(content, walk, context);
      }
    } else if (TRANSPARENT_WRAPPERS.has(name)) {
      readInline(child, walk, context);
    } else if (name === "bookmarkStart") {
      read.bookmarks.push({
        name: wAttr(child, "name") ?? "",
        id: wAttr(child, "id") ?? "",
      });
    } else if (name === "bookmarkEnd") {
      read.bookmarkEndIds.push(wAttr(child, "id") ?? "");
    }
  }
};

/** Read one `w:p`, advancing the story's field and revision state. */
export const readParagraph = (
  element: XmlElement,
  state: StoryState,
): ParagraphRead => {
  const properties = wChild(element, "pPr");
  const mark = wChild(properties, "rPr");
  const hasTrackedMark =
    wChild(mark, "ins") !== null ||
    wChild(mark, "del") !== null ||
    wChild(mark, "moveFrom") !== null ||
    wChild(mark, "moveTo") !== null;
  if (hasTrackedMark) {
    state.trackedChanges += 1;
  }
  for (const child of elementChildren(properties)) {
    if (PROPERTY_REVISIONS.has(wName(child) ?? "")) {
      state.formattingRevisions += 1;
    }
  }
  const outline = wChildVal(properties, "outlineLvl");
  const read: ParagraphRead = {
    element,
    ordinal: state.paragraphOrdinal,
    styleId: wChildVal(properties, "pStyle"),
    numPr: readParagraphNumbering(wChild(properties, "numPr")),
    paraId: paraIdOf(element),
    outlineLevel: headingLevelOf(outline),
    atoms: [],
    text: "",
    fieldIndices: [],
    bookmarks: [],
    bookmarkEndIds: [],
    inlineObjects: 0,
    hasSectionBreak: wChild(properties, "sectPr") !== null,
    hasTrackedMark,
  };
  state.paragraphOrdinal += 1;
  const walk: ParagraphWalk = { state, read, textLength: 0 };
  // A field continuing from an earlier paragraph resumes its result here.
  for (const open of state.open) {
    if (open.phase === "result" && open.field.result !== "") {
      open.field.result += "\n";
    }
  }
  readInline(element, walk, { link: null });
  read.text = read.atoms.map(({ text }) => text).join("");
  return read;
};

export type StoryBlockVisitor = {
  paragraph: (element: XmlElement, depth: number) => void;
  /** Return the walker for the table's cells; called before its rows are read. */
  tableStart: (element: XmlElement, depth: number) => void;
  cellStart: (row: number, col: number) => void;
  tableEnd: (element: XmlElement) => void;
  /** A block-level element the walk does not read. */
  unread: (element: XmlElement) => void;
};

const BLOCK_WRAPPERS = new Set(["customXml", "ins", "moveTo"]);
const IGNORED_BLOCK_CHILDREN = new Set([
  "sectPr",
  "bookmarkStart",
  "bookmarkEnd",
  "proofErr",
  "permStart",
  "permEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "tblPr",
  "tblGrid",
  "tblPrEx",
  "trPr",
  "tcPr",
  "sdtPr",
  "sdtEndPr",
  "customXmlPr",
]);

const walkRows = (
  container: XmlElement,
  visitor: StoryBlockVisitor,
  depth: number,
  rowCounter: { row: number },
): void => {
  for (const child of elementChildren(container)) {
    const name = wName(child);
    if (name === "tr") {
      const row = rowCounter.row;
      rowCounter.row += 1;
      let col = 0;
      const walkCells = (rowContainer: XmlElement): void => {
        for (const cell of elementChildren(rowContainer)) {
          const cellName = wName(cell);
          if (cellName === "tc") {
            visitor.cellStart(row, col);
            col += 1;
            walkBlocks(cell, visitor, depth + 1);
          } else if (cellName === "sdt") {
            const content = wChild(cell, "sdtContent");
            if (content !== null) {
              walkCells(content);
            }
          } else if (cellName !== null && BLOCK_WRAPPERS.has(cellName)) {
            walkCells(cell);
          }
        }
      };
      walkCells(child);
    } else if (name === "sdt") {
      const content = wChild(child, "sdtContent");
      if (content !== null) {
        walkRows(content, visitor, depth, rowCounter);
      }
    } else if (name !== null && BLOCK_WRAPPERS.has(name)) {
      walkRows(child, visitor, depth, rowCounter);
    }
  }
};

/**
 * Visit the blocks under a story container (`w:body`, `w:tc`, a note, a
 * header) in document order. Block content controls are read through.
 */
export const walkBlocks = (
  container: XmlElement,
  visitor: StoryBlockVisitor,
  depth = 0,
): void => {
  for (const child of elementChildren(container)) {
    const name = wName(child);
    if (name === "p") {
      visitor.paragraph(child, depth);
    } else if (name === "tbl") {
      visitor.tableStart(child, depth);
      walkRows(child, visitor, depth, { row: 0 });
      visitor.tableEnd(child);
    } else if (name === "sdt") {
      const content = wChild(child, "sdtContent");
      if (content !== null) {
        walkBlocks(content, visitor, depth);
      }
    } else if (name !== null && BLOCK_WRAPPERS.has(name)) {
      walkBlocks(child, visitor, depth);
    } else if (name === null || !IGNORED_BLOCK_CHILDREN.has(name)) {
      visitor.unread(child);
    }
  }
};

/** Whether `element` is the `w:p` or `w:tbl` a block id names. */
export const isBlockElement = (element: XmlElement): boolean =>
  isW(element, "p") || isW(element, "tbl");
