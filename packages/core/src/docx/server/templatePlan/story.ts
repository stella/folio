/**
 * One reader for the inline content of a paragraph, shared by the analysis,
 * the run transfer and the verification, so all three agree on what a block's
 * text is and where an offset into it lands.
 */
import type { ParagraphNumberingOverride } from "@stll/docx-core/model";

import { parseFieldType } from "../../fieldParser";
import { readParagraphNumbering } from "../../numberingReference";
import { paraIdAttribute } from "../../paraIdAttribute";
import { getLocalName, type XmlElement } from "../../xmlParser";
import {
  elementChildren,
  MARKUP_COMPATIBILITY_NAMESPACE,
  MATH_NAMESPACE,
  relationshipAttr,
  twipsOrNull,
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
  | "selfNumber"
  | "object";

export type AtomLink =
  /** A link through a relationship; `anchor` is the fragment its element adds. */
  | { kind: "relationship"; relationshipId: string; anchor: string | null }
  | { kind: "url"; url: string; anchor: string | null }
  /** A link to a bookmark of the same document. */
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
  /** The code of the field a `selfNumber` atom stands for. */
  instruction?: string;
  /** The tracked change the atom is part of, in a reading that keeps them. */
  revision?: AtomRevision;
};

/**
 * A tracked change an atom is part of, as a reading that keeps revisions
 * reports it. An inserted atom is content like any other with this tag. A
 * deleted one is a text atom without text: it takes no place in the text
 * the document reads with its changes accepted, and holds what was deleted
 * in `deleted` (a tab as `\t`, a line break as `\n`).
 */
export type AtomRevision = {
  kind: "ins" | "del";
  author: string;
  date: string | null;
  deleted?: string;
};

/**
 * Fields that print a number counted when the document is laid out and
 * that hold no result: the number is text of the document that no cached
 * text stands for, so the field itself is the content.
 */
const SELF_NUMBERING_FIELDS: ReadonlySet<string> = new Set([
  "AUTONUM",
  "AUTONUMLGL",
  "AUTONUMOUT",
]);

/** Whether a field kind prints a number of its own and caches no text. */
export const isSelfNumberingField = (kind: string): boolean =>
  SELF_NUMBERING_FIELDS.has(kind);

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
  /** Whether the field sits in the code of another field. */
  nestedInCode: boolean;
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
  /** Field separators and ends met with no field open to take them. */
  strayFieldEnds: number;
  /**
   * Whether inline insertions and deletions are read as tagged atoms. Left
   * out, an insertion reads as plain content and a deletion is skipped.
   */
  keepsRevisions: boolean;
  /**
   * Revisions a keeping reading could not tag: one inside another, deleted
   * content that is no text, a change inside a drawing or an equation.
   */
  untaggedRevisions: number;
};

export const createStoryState = (
  options: { keepsRevisions?: boolean } = {},
): StoryState => ({
  fields: [],
  open: [],
  paragraphOrdinal: 0,
  comments: 0,
  trackedChanges: 0,
  formattingRevisions: 0,
  strayFieldEnds: 0,
  keepsRevisions: options.keepsRevisions === true,
  untaggedRevisions: 0,
});

/** A bookmark start, with how many atoms of its paragraph come before it. */
/**
 * Close a story at its end: a field still open there never closes, and must
 * not swallow the text of the story read next with the same state.
 */
export const endStory = (state: StoryState): void => {
  state.strayFieldEnds += state.open.length;
  state.open = [];
};

export type ParagraphBookmark = { name: string; id: string; atomIndex: number };

/** A bookmark end, with how many atoms of its paragraph come before it. */
export type ParagraphBookmarkEnd = { id: string; atomIndex: number };

/** The sides of `w:ind` that count in characters of the document grid. */
const CHARACTER_INDENTS = [
  "leftChars",
  "startChars",
  "hangingChars",
  "firstLineChars",
] as const;

/** The sides a `w:ind` states; a hanging indent is a negative first line. */
const ownIndent = (
  indent: XmlElement | null,
): { left: number | null; firstLine: number | null } => {
  const hanging = twipsOrNull(wAttr(indent, "hanging"));
  return {
    left: twipsOrNull(wAttr(indent, "left") ?? wAttr(indent, "start")),
    firstLine:
      hanging === null ? twipsOrNull(wAttr(indent, "firstLine")) : -hanging,
  };
};

export type ParagraphRead = {
  element: XmlElement;
  ordinal: number;
  styleId: string | null;
  /** The language the paragraph mark states (`w:pPr/w:rPr/w:lang`), or null. */
  language: string | null;
  numPr: ParagraphNumberingOverride | undefined;
  paraId: string | null;
  outlineLevel: number | null;
  atoms: InlineAtom[];
  text: string;
  /** Indices into `StoryState.fields` of the fields beginning here. */
  fieldIndices: number[];
  /** Fields of earlier paragraphs still open where this one starts. */
  openFieldsAtStart: number;
  bookmarks: ParagraphBookmark[];
  bookmarkEnds: ParagraphBookmarkEnd[];
  inlineObjects: number;
  /** Names of content the reader has no text for and the transfer cannot carry. */
  unsupported: string[];
  hasSectionBreak: boolean;
  /** `w:pageBreakBefore` as the paragraph itself states it; null when it does not. */
  pageBreakBefore: boolean | null;
  /** `w:jc` as the paragraph itself states it, or null. */
  alignment: string | null;
  /** The sides of `w:ind` the paragraph itself states, in twentieths of a
   *  point: where its lines start, and what its first line adds to that
   *  (negative for a hanging indent). Null for a side it does not state. */
  indent: { left: number | null; firstLine: number | null };
  /** Whether the paragraph states an indent in characters, which stands
   *  over the measured one and cannot be given in twentieths of a point. */
  indentInCharacters: boolean;
  /** Whether the paragraph mark is a tracked insertion or deletion. */
  hasTrackedMark: boolean;
};

const TOGGLE_OFF = new Set(["0", "false", "off"]);

/** An on/off property as its element states it; null without the element. */
export const toggleOf = (element: XmlElement | null): boolean | null =>
  element === null ? null : !TOGGLE_OFF.has(wAttr(element, "val") ?? "1");

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
/** Run children that carry no content of their own, or are read elsewhere. */
const IGNORED_RUN_CHILDREN = new Set([
  "rPr",
  "lastRenderedPageBreak",
  "footnoteRef",
  "endnoteRef",
  "annotationRef",
  "separator",
  "continuationSeparator",
  "delText",
  "delInstrText",
]);
/** Paragraph-level children that mark ranges or carry properties only. */
const IGNORED_INLINE_CHILDREN = new Set([
  "pPr",
  "rPr",
  "sdtPr",
  "sdtEndPr",
  "smartTagPr",
  "customXmlPr",
  "commentRangeStart",
  "commentRangeEnd",
  "proofErr",
  "permStart",
  "permEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "customXmlInsRangeStart",
  "customXmlInsRangeEnd",
  "customXmlDelRangeStart",
  "customXmlDelRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveToRangeStart",
  "customXmlMoveToRangeEnd",
]);
const CONTENT_REVISIONS = new Set(["ins", "del", "moveFrom", "moveTo"]);

const REPLACEMENT_CHARACTER = "\uFFFD";
const FIRST_SURROGATE = 0xd800;
const LAST_SURROGATE = 0xdfff;
const FIRST_PRINTABLE = 0x20;

/** The character a `w:sym` shows; one a string cannot hold whole reads as U+FFFD. */
const symbolText = (element: XmlElement): string => {
  const code = Number.parseInt(wAttr(element, "char") ?? "", 16);
  return Number.isInteger(code) &&
    code >= FIRST_PRINTABLE &&
    code <= 0x10ffff &&
    (code < FIRST_SURROGATE || code > LAST_SURROGATE)
    ? String.fromCodePoint(code)
    : REPLACEMENT_CHARACTER;
};

/** Tracked insertions, deletions and moves anywhere under `element`. */
const countContentRevisions = (element: XmlElement): number => {
  let count = 0;
  const stack: XmlElement[] = [element];
  while (stack.length > 0) {
    const node = stack.pop();
    if (node === undefined) {
      continue;
    }
    if (CONTENT_REVISIONS.has(wName(node) ?? "")) {
      count += 1;
    }
    stack.push(...elementChildren(node));
  }
  return count;
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
    nestedInCode: inCode(state),
  };
  state.fields.push(field);
  read.fieldIndices.push(field.index);
  return field;
};

type RunContext = {
  link: AtomLink | null;
  /** The insertion or deletion the content sits in, when revisions are kept. */
  revision?: AtomRevision | undefined;
};

/** What a deleted run holds as text; null for content that is none. */
const deletedTextOf = (child: XmlElement): string | null => {
  switch (wName(child)) {
    case "delText":
      return textOf(child);
    case "tab":
      return "\t";
    case "cr":
      return "\n";
    case "br":
      return wAttr(child, "type") === null ? "\n" : null;
    default:
      return null;
  }
};

/** A deleted run, read as one atom that holds what was deleted. */
const readDeletedRun = (
  run: XmlElement,
  walk: ParagraphWalk,
  { link, revision }: RunContext & { revision: AtomRevision },
): void => {
  let deleted = "";
  for (const child of elementChildren(run)) {
    const name = wName(child);
    const text = deletedTextOf(child);
    if (text !== null) {
      deleted += text;
    } else if (name !== "rPr" && name !== "lastRenderedPageBreak") {
      walk.state.untaggedRevisions += 1;
    }
  }
  if (deleted !== "") {
    pushAtom(walk, {
      kind: "text",
      text: "",
      runProperties: wChild(run, "rPr"),
      link,
      element: null,
      revision: { ...revision, deleted },
    });
  }
};

const readRun = (
  run: XmlElement,
  walk: ParagraphWalk,
  context: RunContext,
): void => {
  const { link, revision } = context;
  const { state, read } = walk;
  const runProperties = wChild(run, "rPr");
  if (wChild(runProperties, "rPrChange") !== null) {
    state.formattingRevisions += 1;
  }
  if (revision?.kind === "del") {
    readDeletedRun(run, walk, { link, revision });
    return;
  }
  const atom = (
    kind: AtomKind,
    text: string,
    element: XmlElement | null,
  ): void => {
    pushAtom(walk, {
      kind,
      text,
      runProperties,
      link,
      element,
      ...(revision === undefined ? {} : { revision }),
    });
  };
  // An object travels whole, so what it holds inside is checked here.
  const object = (element: XmlElement): void => {
    if (inCode(state)) {
      return;
    }
    read.inlineObjects += 1;
    const inside = countContentRevisions(element);
    state.trackedChanges += inside;
    // A change inside an object is none this reading can carry.
    state.untaggedRevisions += state.keepsRevisions ? inside : 0;
    atom("object", "", element);
  };
  for (const child of elementChildren(run)) {
    const name = wName(child);
    if (name === null) {
      if (
        child.namespaceUri === MARKUP_COMPATIBILITY_NAMESPACE &&
        getLocalName(child.name) === "AlternateContent"
      ) {
        object(child);
      } else if (!inCode(state)) {
        read.unsupported.push(child.name ?? "");
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
          if (top === undefined || top.phase === "result") {
            state.strayFieldEnds += 1;
          } else {
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
            if (top.phase === "code" && isSelfNumberingField(top.field.kind)) {
              pushAtom(walk, {
                kind: "selfNumber",
                text: "",
                runProperties,
                link,
                element: null,
                instruction: top.field.instruction.trim(),
              });
            }
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
        if (OBJECT_RUN_CHILDREN.has(name)) {
          object(child);
        } else if (!IGNORED_RUN_CHILDREN.has(name) && !inCode(state)) {
          // Ruby text, legacy date and page marks: content with no reader.
          read.unsupported.push(name);
        }
        break;
    }
  }
};

/** The tag of an inline insertion or deletion; null for a move. */
const revisionOf = (element: XmlElement, name: string): AtomRevision | null => {
  if (name !== "ins" && name !== "del") {
    return null;
  }
  return {
    kind: name,
    author: wAttr(element, "author") ?? "",
    date: wAttr(element, "date"),
  };
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
        // An equation may hold tracked runs of its own.
        const inside = countContentRevisions(child);
        state.trackedChanges += inside;
        state.untaggedRevisions += state.keepsRevisions ? inside : 0;
        pushAtom(walk, {
          kind: "object",
          text: "",
          runProperties: null,
          link: context.link,
          element: child,
          ...(context.revision === undefined
            ? {}
            : { revision: context.revision }),
        });
      } else if (!inCode(state)) {
        read.unsupported.push(child.name ?? "");
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
        link = { kind: "relationship", relationshipId, anchor };
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
      if (field.result === "" && isSelfNumberingField(field.kind)) {
        pushAtom(walk, {
          kind: "selfNumber",
          text: "",
          runProperties: null,
          link: context.link,
          element: null,
          instruction: field.instruction.trim(),
        });
      }
    } else if (REVISION_INCLUDED.has(name) || REVISION_EXCLUDED.has(name)) {
      state.trackedChanges += 1;
      const kept = state.keepsRevisions ? revisionOf(child, name) : null;
      if (kept !== null && context.revision !== undefined) {
        // A change inside a change: neither tag says both.
        state.untaggedRevisions += 1;
      }
      if (kept !== null) {
        readInline(child, walk, { ...context, revision: kept });
        continue;
      }
      // A move is no insertion or deletion a keeping reading tags.
      state.untaggedRevisions += state.keepsRevisions ? 1 : 0;
      if (REVISION_INCLUDED.has(name)) {
        readInline(child, walk, context);
      }
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
        atomIndex: read.atoms.length,
      });
    } else if (name === "bookmarkEnd") {
      read.bookmarkEnds.push({
        id: wAttr(child, "id") ?? "",
        atomIndex: read.atoms.length,
      });
    } else if (!IGNORED_INLINE_CHILDREN.has(name) && !inCode(state)) {
      read.unsupported.push(name);
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
    language: wChildVal(wChild(properties, "rPr"), "lang"),
    numPr: readParagraphNumbering(wChild(properties, "numPr")),
    paraId: paraIdAttribute(element) ?? null,
    outlineLevel: headingLevelOf(outline),
    atoms: [],
    text: "",
    fieldIndices: [],
    openFieldsAtStart: state.open.length,
    bookmarks: [],
    bookmarkEnds: [],
    inlineObjects: 0,
    unsupported: [],
    hasSectionBreak: wChild(properties, "sectPr") !== null,
    pageBreakBefore: toggleOf(wChild(properties, "pageBreakBefore")),
    alignment: wChildVal(properties, "jc"),
    indent: ownIndent(wChild(properties, "ind")),
    indentInCharacters: CHARACTER_INDENTS.some((name) => {
      const value = wAttr(wChild(properties, "ind"), name);
      return value !== null && !/^[+-]?0*$/u.test(value.trim());
    }),
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
  /** A tracked insertion, deletion or move of a block, row or cell. */
  tracked: (element: XmlElement) => void;
};

const BLOCK_WRAPPERS = new Set(["customXml"]);
const ROW_REVISIONS = ["ins", "del"] as const;
const CELL_REVISIONS = ["cellIns", "cellDel", "cellMerge"] as const;

const hasRevisionChild = (
  properties: XmlElement | null,
  names: readonly string[],
): boolean => names.some((name) => wChild(properties, name) !== null);
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
      if (hasRevisionChild(wChild(child, "trPr"), ROW_REVISIONS)) {
        visitor.tracked(child);
      }
      const row = rowCounter.row;
      rowCounter.row += 1;
      let col = 0;
      const walkCells = (rowContainer: XmlElement): void => {
        for (const cell of elementChildren(rowContainer)) {
          const cellName = wName(cell);
          if (cellName === "tc") {
            if (hasRevisionChild(wChild(cell, "tcPr"), CELL_REVISIONS)) {
              visitor.tracked(cell);
            }
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
          } else if (cellName !== null && CONTENT_REVISIONS.has(cellName)) {
            visitor.tracked(cell);
            if (REVISION_INCLUDED.has(cellName)) {
              walkCells(cell);
            }
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
    } else if (name !== null && CONTENT_REVISIONS.has(name)) {
      visitor.tracked(child);
      if (REVISION_INCLUDED.has(name)) {
        walkRows(child, visitor, depth, rowCounter);
      }
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
    } else if (name !== null && CONTENT_REVISIONS.has(name)) {
      // An insertion's blocks are part of the story; a deletion's are not.
      visitor.tracked(child);
      if (REVISION_INCLUDED.has(name)) {
        walkBlocks(child, visitor, depth);
      }
    } else if (name === null || !IGNORED_BLOCK_CHILDREN.has(name)) {
      visitor.unread(child);
    }
  }
};
