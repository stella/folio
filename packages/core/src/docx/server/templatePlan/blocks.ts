/** The block manifest of a DOCX: its main story as addressable, hash-pinned blocks. */
import { Result } from "better-result";

import type { Style } from "../../../types/document";
import { type NumberingMap, parseNumbering } from "../../numberingParser";
import { parseSectionProperties } from "../../sectionParser";
import { parseStylesPackage, type StyleMap } from "../../styleParser";
import type { XmlElement } from "../../xmlParser";
import { keepsShiftReturnSpacing } from "./builders";
import { defaultsBaseline, type Emphasis } from "./emphasis";
import {
  createLabelCounter,
  type LabelCounter,
  type ParagraphLabel,
} from "./labels";
import {
  guarded,
  type LoadedPackage,
  loadPackage,
  readPartText,
  sha256Hex,
} from "./package";
import {
  keptRevisionObstacles,
  resolveTrackedChanges,
  withDeletedMarksJoined,
} from "./revisions";
import {
  createStoryState,
  endStory,
  type ParagraphRead,
  readParagraph,
  type StoryState,
  toggleOf,
  walkBlocks,
} from "./story";
import type {
  DocxBlock,
  DocxBlockManifest,
  DocxNote,
  DocxSection,
  DocxTableCell,
  TemplatePlanTrackedChanges,
} from "./types";
import { TemplatePlanError } from "./types";
import {
  elementChildren,
  halfPointsOrNull,
  integerOrNull,
  isW,
  type RawPart,
  readRawPart,
  twipsOrNull,
  wAttr,
  wChild,
  wChildren,
  wChildVal,
} from "./xml";

const BLOCK_HASH_LENGTH = 12;

export const blockHash = (text: string, label: string | null): string =>
  sha256Hex(`${text}\u0000${label ?? ""}`).slice(0, BLOCK_HASH_LENGTH);

export type AnalyzedBlock = {
  block: DocxBlock;
  element: XmlElement;
  /** Null for a table. */
  read: ParagraphRead | null;
  label: ParagraphLabel | null;
  /** The table directly holding the block, when it is nested. */
  parentId: string | null;
  /** The top-level block the block is, or sits in. */
  topLevelId: string;
};

export type AnalyzedStory = {
  topLevel: AnalyzedBlock[];
  /** Every block in document order, nested ones included. */
  order: AnalyzedBlock[];
  byId: Map<string, AnalyzedBlock>;
  /** Paragraph blocks by the ordinal their read carries. */
  paragraphs: AnalyzedBlock[];
  state: StoryState;
  /** `w:sectPr` elements in document order, the body's own last. */
  sections: XmlElement[];
  /** Whether the body states section properties of its own. */
  hasBodySection: boolean;
  /** Names of block-level elements the walk has no reading for. */
  unreadBlocks: string[];
  warnings: string[];
};

type StoryResources = {
  styles: StyleMap;
  defaultStyleId: string | null;
  /** The language the document defaults give every run, or null. */
  defaultLanguage: string | null;
  numbering: NumberingMap | null;
  /** The alignment each paragraph style states for itself (`w:jc`). */
  alignments?: ReadonlyMap<string, string>;
  /** The indent each list level states, by definition and level. */
  levelIndents?: ReadonlyMap<string, IndentSides>;
  /** The indent the document's defaults give every paragraph. */
  defaultIndent?: IndentSides;
};

type OpenTable = {
  analyzed: AnalyzedBlock;
  cells: DocxTableCell[];
  nested: DocxBlock[];
  cellTexts: string[][];
};

/** The heading level a style outlines its paragraphs at, 0 being the highest. */
export const styleOutlineLevel = (style: Style | undefined): number | null => {
  const outline = style?.pPr?.outlineLevel;
  return outline?.kind === "heading" ? outline.level : null;
};

const styleNameOf = (style: Style | undefined): string | null =>
  style?.name ?? null;

/** A section's page and how it begins, as its `w:sectPr` states them. */
const sectionOf = (sectPr: XmlElement, index: number): DocxSection => {
  const properties = parseSectionProperties(sectPr);
  return {
    index,
    start: properties.sectionStart ?? "nextPage",
    orientation: properties.orientation ?? "portrait",
    size: {
      width: properties.pageWidth ?? null,
      height: properties.pageHeight ?? null,
    },
    margins: {
      top: properties.marginTop ?? null,
      right: properties.marginRight ?? null,
      bottom: properties.marginBottom ?? null,
      left: properties.marginLeft ?? null,
    },
    columns: properties.columnCount ?? 1,
  };
};

type IndentSides = { left: number | null; firstLine: number | null };

const NO_INDENT: IndentSides = { left: null, firstLine: null };

const levelKey = (abstractNumId: number, ilvl: number): string =>
  `${String(abstractNumId)}:${String(ilvl)}`;

/** The sides a `w:ind` states; a hanging indent is a negative first line. */
const statedIndent = (indent: XmlElement | null): IndentSides => {
  const hanging = twipsOrNull(wAttr(indent, "hanging"));
  return {
    left: twipsOrNull(wAttr(indent, "left") ?? wAttr(indent, "start")),
    firstLine:
      hanging === null ? twipsOrNull(wAttr(indent, "firstLine")) : -hanging,
  };
};

/**
 * The indent each level of a numbering part states, by list definition and
 * level, read off the part itself. An override a list instance makes of a
 * level is not followed.
 */
const statedLevelIndents = (
  numberingPath: string | null,
  numberingXml: string | null,
): Map<string, IndentSides> => {
  const indents = new Map<string, IndentSides>();
  const part =
    numberingPath === null || numberingXml === null
      ? null
      : readRawPart(numberingPath, numberingXml);
  for (const definition of wChildren(part?.root, "abstractNum")) {
    const abstractNumId = integerOrNull(wAttr(definition, "abstractNumId"));
    for (const level of wChildren(definition, "lvl")) {
      const ilvl = integerOrNull(wAttr(level, "ilvl"));
      if (abstractNumId !== null && ilvl !== null) {
        indents.set(
          levelKey(abstractNumId, ilvl),
          statedIndent(wChild(wChild(level, "pPr"), "ind")),
        );
      }
    }
  }
  return indents;
};

/** The indent the document's defaults give every paragraph. */
const statedDefaultIndent = (
  stylesPath: string | null,
  stylesXml: string | null,
): IndentSides => {
  const part =
    stylesPath === null || stylesXml === null
      ? null
      : readRawPart(stylesPath, stylesXml);
  const defaults = wChild(wChild(part?.root, "docDefaults"), "pPrDefault");
  return statedIndent(wChild(wChild(defaults, "pPr"), "ind"));
};

/** The sides a style states, as the document model holds them: a hanging
 *  indent is a negative first line. */
const modelIndent = (
  formatting:
    | {
        indentLeft?: number | undefined;
        indentFirstLine?: number | undefined;
        hangingIndent?: boolean | undefined;
      }
    | undefined,
): IndentSides => {
  const first = formatting?.indentFirstLine;
  let firstLine: number | null = null;
  if (first !== undefined) {
    firstLine = formatting?.hangingIndent === true ? -Math.abs(first) : first;
  }
  return { left: formatting?.indentLeft ?? null, firstLine };
};

/**
 * The indent in effect: each side from the first of `sources` that states
 * it. Null when none of them states a side.
 */
const indentInEffect = (
  ...sources: readonly IndentSides[]
): { left: number; firstLine: number } | null => {
  const left = sources.find((source) => source.left !== null)?.left ?? null;
  const firstLine =
    sources.find((source) => source.firstLine !== null)?.firstLine ?? null;
  return left === null && firstLine === null
    ? null
    : { left: left ?? 0, firstLine: firstLine ?? 0 };
};

/** The alignment each style of a styles part states for itself (`w:jc`). */
export const statedAlignments = (
  stylesPath: string | null,
  stylesXml: string | null,
): Map<string, string> => {
  const alignments = new Map<string, string>();
  const part =
    stylesPath === null || stylesXml === null
      ? null
      : readRawPart(stylesPath, stylesXml);
  for (const style of wChildren(part?.root, "style")) {
    const styleId = wAttr(style, "styleId");
    const alignment = wChildVal(wChild(style, "pPr"), "jc");
    if (styleId !== null && alignment !== null) {
      alignments.set(styleId, alignment);
    }
  }
  return alignments;
};

/** The alignment of the nearest style in the `basedOn` chain that states one. */
export const styleAlignment = (
  styleId: string | null,
  styles: StyleMap,
  stated: ReadonlyMap<string, string>,
): string | null => {
  const seen = new Set<string>();
  let cursor: string | undefined = styleId ?? undefined;
  while (cursor !== undefined && !seen.has(cursor)) {
    seen.add(cursor);
    const alignment = stated.get(cursor);
    if (alignment !== undefined) {
      return alignment;
    }
    cursor = styles.get(cursor)?.basedOn;
  }
  return null;
};

/** What a style states of the look of its text; null where it states
 *  nothing. */
type StatedLook = {
  bold: boolean | null;
  /** In half-points. */
  size: number | null;
  keepNext: boolean | null;
  /** In twentieths of a point. */
  indentLeft: number | null;
};

const statedLook = (
  paragraph: XmlElement | null,
  run: XmlElement | null,
): StatedLook => ({
  bold: toggleOf(wChild(run, "b")),
  size: halfPointsOrNull(wChildVal(run, "sz")),
  keepNext: toggleOf(wChild(paragraph, "keepNext")),
  indentLeft: twipsOrNull(
    wAttr(wChild(paragraph, "ind"), "left") ??
      wAttr(wChild(paragraph, "ind"), "start"),
  ),
});

/** The look each style of the styles part states itself, by style id, and
 *  the look the document's defaults state. */
export const statedLooks = (
  stylesPath: string | null,
  stylesXml: string | null,
): { styles: Map<string, StatedLook>; defaults: StatedLook } => {
  const styles = new Map<string, StatedLook>();
  const part =
    stylesPath === null || stylesXml === null
      ? null
      : readRawPart(stylesPath, stylesXml);
  for (const style of wChildren(part?.root, "style")) {
    const styleId = wAttr(style, "styleId");
    if (styleId !== null) {
      styles.set(
        styleId,
        statedLook(wChild(style, "pPr"), wChild(style, "rPr")),
      );
    }
  }
  const defaults = wChild(part?.root, "docDefaults");
  return {
    styles,
    defaults: statedLook(
      wChild(wChild(defaults, "pPrDefault"), "pPr"),
      wChild(wChild(defaults, "rPrDefault"), "rPr"),
    ),
  };
};

/**
 * The look a style gives its text: each property from the nearest style in
 * the `basedOn` chain that states it, then from the document's defaults.
 * Without a statement text is not bold and a paragraph is not kept with the
 * next; the size and the left indent stay null, which leaves the size to
 * the application and the indent to a list level, if the style has one.
 */
export const styleLook = (
  styleId: string | null,
  styles: StyleMap,
  { styles: stated, defaults }: ReturnType<typeof statedLooks>,
): {
  bold: boolean;
  size: number | null;
  keepNext: boolean;
  indentLeft: number | null;
} => {
  const look: StatedLook = {
    bold: null,
    size: null,
    keepNext: null,
    indentLeft: null,
  };
  const seen = new Set<string>();
  let cursor: string | undefined = styleId ?? undefined;
  while (cursor !== undefined && !seen.has(cursor)) {
    seen.add(cursor);
    const own = stated.get(cursor);
    look.bold ??= own?.bold ?? null;
    look.size ??= own?.size ?? null;
    look.keepNext ??= own?.keepNext ?? null;
    look.indentLeft ??= own?.indentLeft ?? null;
    cursor = styles.get(cursor)?.basedOn;
  }
  return {
    bold: look.bold ?? defaults.bold ?? false,
    size: look.size ?? defaults.size,
    keepNext: look.keepNext ?? defaults.keepNext ?? false,
    indentLeft: look.indentLeft ?? defaults.indentLeft,
  };
};

/**
 * The sections of a story with what each states. A body without section
 * properties of its own still ends in a section, which states nothing.
 */
export const storySections = ({
  sections,
  hasBodySection,
}: AnalyzedStory): DocxSection[] => {
  const stated = sections.map(sectionOf);
  return hasBodySection
    ? stated
    : [
        ...stated,
        {
          index: stated.length,
          start: "nextPage",
          orientation: "portrait",
          size: { width: null, height: null },
          margins: { top: null, right: null, bottom: null, left: null },
          columns: 1,
        },
      ];
};

/** Page breaks among a paragraph's runs; a column break is not one. */
const pageBreaksOf = ({ atoms }: ParagraphRead): number =>
  atoms.filter(
    ({ kind, element }) =>
      kind === "pageBreak" && wAttr(element, "type") === "page",
  ).length;

/** Where in a paragraph's text its self-numbering fields stand. */
const selfNumbersOf = ({ atoms }: ParagraphRead): number[] => {
  const offsets: number[] = [];
  let at = 0;
  for (const { kind, text } of atoms) {
    if (kind === "selfNumber") {
      offsets.push(at);
    }
    at += text.length;
  }
  return offsets;
};

type AnalyzeStoryOptions = {
  container: XmlElement;
  resources: StoryResources;
  counter?: LabelCounter;
  /** Read inline insertions and deletions as tagged atoms. */
  keepsRevisions?: boolean;
};

/** Read the blocks under `container` into ids, text, labels and hashes. */
export const analyzeStory = ({
  container,
  resources,
  counter = createLabelCounter(resources),
  keepsRevisions = false,
}: AnalyzeStoryOptions): AnalyzedStory => {
  const state = createStoryState({ keepsRevisions });
  const story: AnalyzedStory = {
    topLevel: [],
    order: [],
    byId: new Map(),
    paragraphs: [],
    state,
    sections: [],
    hasBodySection: false,
    unreadBlocks: [],
    warnings: [],
  };
  const tables: OpenTable[] = [];
  let index = 0;
  let sectionIndex = 0;
  const unread = new Map<string, number>();

  const register = (analyzed: AnalyzedBlock): void => {
    story.order.push(analyzed);
    story.byId.set(analyzed.block.id, analyzed);
    const parent = tables.at(-1);
    if (parent === undefined) {
      story.topLevel.push(analyzed);
      return;
    }
    parent.cells.at(-1)?.blockIds.push(analyzed.block.id);
    for (const open of tables) {
      open.nested.push(analyzed.block);
    }
  };

  const noteCellText = (text: string): void => {
    tables.at(-1)?.cellTexts.at(-1)?.push(text);
  };

  walkBlocks(container, {
    paragraph: (element) => {
      const read = readParagraph(element, state);
      const styleId = read.styleId;
      const style = resources.styles.get(
        styleId ?? resources.defaultStyleId ?? "",
      );
      const label = counter.next({
        styleId,
        numPr: read.numPr,
        language: read.language,
      });
      const id = `p${index}`;
      const block: DocxBlock = {
        id,
        index,
        kind: "paragraph",
        text: read.text,
        label: label.label,
        numbering: label.numbering,
        styleId,
        styleName: styleNameOf(style),
        outlineLevel: read.outlineLevel ?? styleOutlineLevel(style),
        sectionIndex,
        hash: blockHash(read.text, label.label),
        ...(read.paraId === null ? {} : { paraId: read.paraId }),
        inlineObjects: read.inlineObjects,
      };
      const pageBreaks = pageBreaksOf(read);
      if (pageBreaks > 0) {
        block.pageBreaks = pageBreaks;
      }
      const selfNumbers = selfNumbersOf(read);
      if (selfNumbers.length > 0) {
        block.selfNumbers = selfNumbers;
      }
      if (read.pageBreakBefore ?? style?.pPr?.pageBreakBefore === true) {
        block.pageBreakBefore = true;
      }
      if (read.hasSectionBreak) {
        block.endsSection = true;
      }
      const alignment =
        read.alignment ??
        (resources.alignments === undefined
          ? null
          : styleAlignment(
              styleId ?? resources.defaultStyleId,
              resources.styles,
              resources.alignments,
            ));
      if (alignment !== null) {
        block.alignment = alignment;
      }
      // A list the paragraph names itself gives its level's indent over
      // the style's; a list the style brings is part of the style, whose
      // own indent then stands over the level's.
      const ofLevel =
        (label.numbering === null
          ? undefined
          : resources.levelIndents?.get(
              levelKey(label.numbering.abstractNumId, label.numbering.ilvl),
            )) ?? NO_INDENT;
      const ofStyle = modelIndent(style?.pPr);
      const ofDefaults = resources.defaultIndent ?? NO_INDENT;
      const indent =
        read.numPr === undefined
          ? indentInEffect(read.indent, ofStyle, ofLevel, ofDefaults)
          : indentInEffect(read.indent, ofLevel, ofStyle, ofDefaults);
      // An indent counted in characters cannot be said in this unit.
      if (indent !== null && !read.indentInCharacters) {
        block.indent = indent;
      }
      index += 1;
      const parent = tables.at(-1);
      const analyzed: AnalyzedBlock = {
        block,
        element,
        read,
        label,
        parentId: parent?.analyzed.block.id ?? null,
        topLevelId: tables.at(0)?.analyzed.block.id ?? id,
      };
      register(analyzed);
      story.paragraphs[read.ordinal] = analyzed;
      noteCellText(read.text);
      if (read.hasSectionBreak) {
        const sectPr = wChild(wChild(element, "pPr"), "sectPr");
        if (sectPr !== null) {
          story.sections.push(sectPr);
        }
        sectionIndex += 1;
      }
    },
    tableStart: (element) => {
      const id = `t${index}`;
      const properties = wChild(element, "tblPr");
      const styleId = wAttr(wChild(properties, "tblStyle"), "val");
      const block: DocxBlock = {
        id,
        index,
        kind: "table",
        text: "",
        label: null,
        numbering: null,
        styleId,
        styleName: styleNameOf(resources.styles.get(styleId ?? "")),
        outlineLevel: null,
        sectionIndex,
        hash: "",
        inlineObjects: 0,
        cells: [],
        nested: [],
      };
      index += 1;
      const parent = tables.at(-1);
      const analyzed: AnalyzedBlock = {
        block,
        element,
        read: null,
        label: null,
        parentId: parent?.analyzed.block.id ?? null,
        topLevelId: tables.at(0)?.analyzed.block.id ?? id,
      };
      register(analyzed);
      tables.push({ analyzed, cells: [], nested: [], cellTexts: [] });
    },
    cellStart: (row, col) => {
      const open = tables.at(-1);
      if (open === undefined) {
        return;
      }
      open.cells.push({ row, col, blockIds: [] });
      open.cellTexts.push([]);
    },
    tableEnd: () => {
      const open = tables.pop();
      if (open === undefined) {
        return;
      }
      const rows = new Map<number, string[]>();
      open.cells.forEach(({ row }, cellIndex) => {
        const texts = rows.get(row) ?? [];
        texts.push((open.cellTexts[cellIndex] ?? []).join("\n"));
        rows.set(row, texts);
      });
      const { block } = open.analyzed;
      block.text = [...rows.values()]
        .map((cells) => cells.join("\t"))
        .join("\n");
      block.cells = open.cells;
      block.nested = open.nested;
      let inlineObjects = 0;
      for (const nested of open.nested) {
        if (nested.kind === "paragraph") {
          inlineObjects += nested.inlineObjects;
        }
      }
      block.inlineObjects = inlineObjects;
      block.hash = blockHash(block.text, null);
      noteCellText(block.text);
    },
    unread: (element) => {
      const name = element.name ?? "";
      unread.set(name, (unread.get(name) ?? 0) + 1);
    },
    tracked: () => {
      state.trackedChanges += 1;
    },
  });

  const bodySection = wChild(container, "sectPr");
  if (bodySection !== null) {
    story.sections.push(bodySection);
    story.hasBodySection = true;
  }
  for (const [name, count] of unread) {
    story.unreadBlocks.push(name);
    story.warnings.push(
      `${count} block-level ${name} element(s) are not read as blocks.`,
    );
  }
  if (state.open.length > 0 || state.strayFieldEnds > 0) {
    story.warnings.push(
      "The story holds complex field marks that do not pair up; text after an unclosed field reads as field code.",
    );
  }
  return story;
};

export type AnalyzedNote = {
  id: string;
  element: XmlElement;
  text: string;
  paragraphs: ParagraphRead[];
  /** Whether the note holds a table or another block that is not a paragraph. */
  hasOtherBlocks: boolean;
};

export type AnalyzedNotes = {
  kind: "footnotes" | "endnotes";
  path: string | null;
  part: RawPart | null;
  /** Notes with content; separators are left out. */
  notes: AnalyzedNote[];
  byId: Map<string, AnalyzedNote>;
  /** Every `w:id` the part uses, separators included. */
  usedIds: number[];
  state: StoryState;
};

const NOTE_ELEMENT = { footnotes: "footnote", endnotes: "endnote" } as const;

const analyzeNotes = (
  kind: "footnotes" | "endnotes",
  path: string | null,
  xml: string | null,
  keepsRevisions = false,
): AnalyzedNotes => {
  const notes: AnalyzedNotes = {
    kind,
    path,
    part: null,
    notes: [],
    byId: new Map(),
    usedIds: [],
    state: createStoryState({ keepsRevisions }),
  };
  if (path === null || xml === null) {
    return notes;
  }
  const part = readRawPart(path, xml);
  if (part === null) {
    // A notes part that cannot be read would lose every note it holds.
    throw new TemplatePlanError({
      message: `The notes part ${path} could not be read.`,
      code: "xml-unreadable",
      part: path,
    });
  }
  notes.part = part;
  for (const element of elementChildren(part.root)) {
    const id = wAttr(element, "id");
    if (id === null) {
      continue;
    }
    const numericId = Number.parseInt(id, 10);
    if (Number.isInteger(numericId)) {
      notes.usedIds.push(numericId);
    }
    const type = wAttr(element, "type");
    if (
      !isW(element, NOTE_ELEMENT[kind]) ||
      (type !== null && type !== "normal")
    ) {
      continue;
    }
    const paragraphs: ParagraphRead[] = [];
    let hasOtherBlocks = false;
    walkBlocks(element, {
      paragraph: (paragraph) => {
        paragraphs.push(readParagraph(paragraph, notes.state));
      },
      tableStart: () => {
        hasOtherBlocks = true;
      },
      cellStart: () => {},
      tableEnd: () => {},
      unread: () => {
        hasOtherBlocks = true;
      },
      tracked: () => {
        notes.state.trackedChanges += 1;
      },
    });
    // Each note is a story of its own.
    endStory(notes.state);
    const note: AnalyzedNote = {
      id,
      element,
      text: paragraphs.map(({ text }) => text).join("\n"),
      paragraphs,
      hasOtherBlocks,
    };
    notes.notes.push(note);
    notes.byId.set(id, note);
  }
  return notes;
};

export type AnalyzedDocument = {
  pkg: LoadedPackage;
  document: RawPart;
  body: XmlElement;
  story: AnalyzedStory;
  styles: StyleMap;
  /** The emphasis the document defaults give every run. */
  defaultEmphasis: Emphasis;
  stylesPath: string | null;
  stylesXml: string | null;
  defaultStyleId: string | null;
  numbering: NumberingMap | null;
  numberingPath: string | null;
  numberingXml: string | null;
  footnotes: AnalyzedNotes;
  endnotes: AnalyzedNotes;
  manifest: DocxBlockManifest;
};

const noteList = ({ notes }: AnalyzedNotes): DocxNote[] =>
  notes.map(({ id, text }) => ({ id, text }));

const countComments = (xml: string | null): number => {
  if (xml === null) {
    return 0;
  }
  const part = readRawPart("comments", xml);
  return part === null
    ? 0
    : elementChildren(part.root).filter(({ name }) =>
        (name ?? "").endsWith("comment"),
      ).length;
};

/** Open a package and analyse its main story, notes, styles and numbering. */
export const analyzeDocument = async (
  bytes: Uint8Array,
  { keepsRevisions = false }: { keepsRevisions?: boolean } = {},
): Promise<Result<AnalyzedDocument, TemplatePlanError>> => {
  const loaded = await loadPackage(bytes);
  if (Result.isError(loaded)) {
    return loaded;
  }
  const pkg = loaded.value;
  const document = readRawPart(pkg.documentPath, pkg.documentXml);
  const body = document === null ? null : wChild(document.root, "body");
  if (document === null || body === null) {
    return Result.err(
      new TemplatePlanError({
        message: "The main document part has no readable body.",
        code: "xml-unreadable",
        part: pkg.documentPath,
      }),
    );
  }

  // A part the document names and the package does not hold is not the same
  // as a part the document never had: reading on would lose what it carried.
  for (const kind of [
    "styles",
    "numbering",
    "footnotes",
    "endnotes",
  ] as const) {
    const missing = pkg.missingPartPath(kind);
    if (missing !== null) {
      return Result.err(
        new TemplatePlanError({
          message: `The main document names the part ${missing}, which the package does not hold.`,
          code: "part-missing",
          part: missing,
        }),
      );
    }
  }
  const stylesPath = pkg.partPath("styles");
  const numberingPath = pkg.partPath("numbering");
  const footnotesPath = pkg.partPath("footnotes");
  const endnotesPath = pkg.partPath("endnotes");
  const texts: (string | null)[] = [];
  for (const path of [
    stylesPath,
    numberingPath,
    footnotesPath,
    endnotesPath,
    pkg.partPath("comments"),
    pkg.partPath("settings"),
  ]) {
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const text = await readPartText(pkg, path);
    if (Result.isError(text)) {
      return text;
    }
    texts.push(text.value);
  }
  const [
    stylesXml = null,
    numberingXml = null,
    footnotesXml = null,
    endnotesXml = null,
    commentsXml = null,
    settingsXml = null,
  ] = texts;

  const read = Result.try({
    try: () => {
      const stylesPackage =
        stylesXml === null ? null : parseStylesPackage(stylesXml, null);
      const styles: StyleMap = stylesPackage?.styles ?? new Map();
      const defaultEmphasis = defaultsBaseline(
        stylesPackage?.styleDefinitions.docDefaults?.rPr,
      );
      const defaultStyleId =
        [...styles.values()].find(
          (style) => style.type === "paragraph" && style.default === true,
        )?.styleId ?? null;
      const numbering =
        numberingXml === null ? null : parseNumbering(numberingXml);
      const story = analyzeStory({
        container: body,
        resources: {
          styles,
          defaultStyleId,
          defaultLanguage:
            stylesPackage?.styleDefinitions.docDefaults?.rPr?.language?.val ??
            null,
          numbering,
          alignments: statedAlignments(stylesPath, stylesXml),
          levelIndents: statedLevelIndents(numberingPath, numberingXml),
          defaultIndent: statedDefaultIndent(stylesPath, stylesXml),
        },
        keepsRevisions,
      });
      const footnotes = analyzeNotes(
        "footnotes",
        footnotesPath,
        footnotesXml,
        keepsRevisions,
      );
      const endnotes = analyzeNotes(
        "endnotes",
        endnotesPath,
        endnotesXml,
        keepsRevisions,
      );
      return {
        styles,
        defaultEmphasis,
        defaultStyleId,
        numbering,
        story,
        footnotes,
        endnotes,
      };
    },
    catch: (cause) =>
      cause instanceof TemplatePlanError
        ? cause
        : new TemplatePlanError({
            message: "The DOCX parts could not be analysed.",
            code: "xml-unreadable",
            cause,
          }),
  });
  if (Result.isError(read)) {
    return read;
  }
  const {
    styles,
    defaultEmphasis,
    defaultStyleId,
    numbering,
    story,
    footnotes,
    endnotes,
  } = read.value;

  const warnings = [...story.warnings];
  const formattingRevisions =
    story.state.formattingRevisions +
    footnotes.state.formattingRevisions +
    endnotes.state.formattingRevisions;
  if (formattingRevisions > 0) {
    warnings.push(
      `${formattingRevisions} tracked formatting change(s) are not counted as tracked changes.`,
    );
  }
  const manifest: DocxBlockManifest = {
    sha256: pkg.sha256,
    blocks: story.topLevel.map(({ block }) => block),
    sections: storySections(story),
    compat: {
      doNotExpandShiftReturn:
        settingsXml !== null && keepsShiftReturnSpacing(settingsXml),
    },
    footnotes: noteList(footnotes),
    endnotes: noteList(endnotes),
    comments: countComments(commentsXml),
    trackedChanges:
      story.state.trackedChanges +
      footnotes.state.trackedChanges +
      endnotes.state.trackedChanges,
    warnings,
  };
  return Result.ok({
    pkg,
    document,
    body,
    story,
    styles,
    defaultEmphasis,
    stylesPath,
    stylesXml,
    defaultStyleId,
    numbering,
    numberingPath,
    numberingXml,
    footnotes,
    endnotes,
    manifest,
  });
};

/** The view a plan states, when what was handed in is shaped like a plan
 *  at all; its shape is checked where the plan is resolved. */
export const viewOf = (
  plan: unknown,
): TemplatePlanTrackedChanges | undefined => {
  const view =
    typeof plan === "object" && plan !== null && "trackedChanges" in plan
      ? plan.trackedChanges
      : undefined;
  return view === "accept" || view === "reject" || view === "keep"
    ? view
    : undefined;
};

/**
 * A source read with its tracked changes kept: the blocks are those of the
 * accepted view, and the atoms of an insertion or a deletion carry its tag,
 * a paragraph mark's among them. A change that is not carried yet (a table
 * row or cell, a move) or cannot be tagged is refused with what stands in
 * the way, for the caller to offer one view instead.
 */
/**
 * A document read with its tracked changes kept, source or output alike:
 * deleted paragraph marks carried as deletions inside the joined paragraph,
 * inline insertions and deletions as tagged atoms. The manifest keeps the
 * hash of the bytes that were given.
 */
export const analyzeKeptForm = async (
  bytes: Uint8Array,
): Promise<Result<AnalyzedDocument, TemplatePlanError>> => {
  const joined = await withDeletedMarksJoined(bytes);
  if (Result.isError(joined)) {
    return joined;
  }
  const analyzed = await analyzeDocument(joined.value, {
    keepsRevisions: true,
  });
  if (Result.isError(analyzed)) {
    return analyzed;
  }
  return Result.ok({
    ...analyzed.value,
    manifest: { ...analyzed.value.manifest, sha256: sha256Hex(bytes) },
  });
};

const analyzeKeeping = async (
  bytes: Uint8Array,
): Promise<Result<AnalyzedDocument, TemplatePlanError>> => {
  const obstacles = await keptRevisionObstacles(bytes);
  if (Result.isError(obstacles)) {
    return obstacles;
  }
  const analyzed = await analyzeKeptForm(bytes);
  if (Result.isError(analyzed)) {
    return analyzed;
  }
  const { story, footnotes, endnotes, manifest } = analyzed.value;
  const untagged =
    story.state.untaggedRevisions +
    footnotes.state.untaggedRevisions +
    endnotes.state.untaggedRevisions;
  const { tableRows, tableCells, moves } = obstacles.value.obstacles;
  const found = [
    [tableRows, "inserted or deleted table row(s)"],
    [tableCells, "table cell change(s)"],
    [moves, "move(s)"],
    [
      untagged,
      "change(s) inside another change, a drawing or an equation, or deleted paragraph mark(s) with no paragraph to join",
    ],
  ].flatMap(([count, what]) =>
    typeof count === "number" && count > 0 ? [`${String(count)} ${what}`] : [],
  );
  if (found.length > 0) {
    return Result.err(
      new TemplatePlanError({
        code: "tracked-changes",
        message: `The tracked changes of the source cannot be kept: ${found.join(", ")}. Format one view of it instead.`,
      }),
    );
  }
  return Result.ok({
    ...analyzed.value,
    manifest: { ...manifest, trackedChangesKept: obstacles.value.tracked },
  });
};

/**
 * Analyse a source in the view a caller chose of its tracked changes, or
 * as it stands. In a view, the package that is analysed is the resolved
 * one, while the manifest keeps the hash of the bytes that were given: a
 * plan pins the file, not a rewriting of it. Without a view, a document
 * that tracks changes says what it tracks.
 */
export const analyzeSource = async (
  bytes: Uint8Array,
  view: TemplatePlanTrackedChanges | undefined,
): Promise<Result<AnalyzedDocument, TemplatePlanError>> => {
  if (view === "keep") {
    return await analyzeKeeping(bytes);
  }
  if (view === undefined) {
    const analyzed = await analyzeDocument(bytes);
    if (
      Result.isError(analyzed) ||
      analyzed.value.manifest.trackedChanges === 0
    ) {
      return analyzed;
    }
    const tracked = await resolveTrackedChanges(bytes, "accept");
    return Result.isError(tracked)
      ? analyzed
      : Result.ok({
          ...analyzed.value,
          manifest: {
            ...analyzed.value.manifest,
            revisions: tracked.value.resolved,
          },
        });
  }
  const resolved = await resolveTrackedChanges(bytes, view);
  if (Result.isError(resolved)) {
    return resolved;
  }
  const analyzed = await analyzeDocument(resolved.value.docx);
  if (Result.isError(analyzed)) {
    return analyzed;
  }
  return Result.ok({
    ...analyzed.value,
    manifest: {
      ...analyzed.value.manifest,
      sha256: sha256Hex(bytes),
      trackedChangesResolved: { view, ...resolved.value.resolved },
    },
  });
};

/**
 * Read a DOCX into the blocks a template plan addresses: ids, visible text,
 * rendered list labels, styles and a hash per block, plus its notes. With
 * `trackedChanges`, the blocks are those of that view of a document with
 * tracked changes; with `keep`, those of the accepted view, read off the
 * document as it stands, or the refusal a plan that keeps them would get.
 */
export const analyzeDocxBlocks = async (
  docx: Uint8Array,
  options: { trackedChanges?: TemplatePlanTrackedChanges } = {},
): Promise<Result<DocxBlockManifest, TemplatePlanError>> =>
  await guarded(async () =>
    (await analyzeSource(docx, options.trackedChanges)).map(
      ({ manifest }) => manifest,
    ),
  );
