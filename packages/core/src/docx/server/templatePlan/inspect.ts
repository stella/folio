/** A template's inventory: what a plan may name when it places blocks onto it. */
import { Result } from "better-result";

import { RELATIONSHIP_TYPES } from "../../relsParser";
import { parseSectionProperties } from "../../sectionParser";
import {
  type AnalyzedDocument,
  analyzeDocument,
  styleOutlineLevel,
} from "./blocks";
import { styleNumbering } from "./labels";
import { guarded, readPartText } from "./package";
import {
  createStoryState,
  endStory,
  readParagraph,
  type StoryField,
  walkBlocks,
} from "./story";
import {
  type DocxTemplateBookmark,
  type DocxTemplateField,
  type DocxTemplateInspection,
  type DocxTemplateNumbering,
  type DocxTemplateNumberingLevel,
  type DocxTemplateSection,
  type DocxTemplateStyle,
  TemplatePlanError,
} from "./types";
import {
  isW,
  type RawPart,
  readRawPart,
  twipsOrNull,
  walkElements,
  wAttr,
  wChild,
  wChildren,
  wChildVal,
} from "./xml";

const DEFAULT_LEVEL_START = 1;

/** The `w:jc` each style of the styles part states itself, by style id. */
const statedAlignments = ({
  stylesPath,
  stylesXml,
}: AnalyzedDocument): Map<string, string> => {
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

type LevelIndent = DocxTemplateNumberingLevel["indent"];

/** The indent each level of the numbering part states, by definition and level. */
const statedIndents = ({
  numberingPath,
  numberingXml,
}: AnalyzedDocument): Map<string, LevelIndent> => {
  const indents = new Map<string, LevelIndent>();
  const part =
    numberingPath === null || numberingXml === null
      ? null
      : readRawPart(numberingPath, numberingXml);
  for (const definition of wChildren(part?.root, "abstractNum")) {
    for (const level of wChildren(definition, "lvl")) {
      const indent = wChild(wChild(level, "pPr"), "ind");
      const stated = (name: string): number | null =>
        twipsOrNull(wAttr(indent, name));
      indents.set(
        `${wAttr(definition, "abstractNumId") ?? ""}:${wAttr(level, "ilvl") ?? ""}`,
        {
          left: stated("left") ?? stated("start"),
          hanging: stated("hanging"),
          firstLine: stated("firstLine"),
        },
      );
    }
  }
  return indents;
};

const NO_INDENT: LevelIndent = {
  left: null,
  hanging: null,
  firstLine: null,
};

const inspectStyles = (analyzed: AnalyzedDocument): DocxTemplateStyle[] => {
  const { styles, numbering } = analyzed;
  const stated = statedAlignments(analyzed);
  /** The alignment of the nearest style in the `basedOn` chain that states one. */
  const alignmentOf = (styleId: string): string | null => {
    const seen = new Set<string>();
    let cursor: string | undefined = styleId;
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
  return [...styles.values()].map((style) => {
    const bound = styleNumbering(style.styleId, styles, numbering);
    return {
      styleId: style.styleId,
      name: style.name ?? null,
      type: style.type,
      basedOn: style.basedOn ?? null,
      numPr:
        bound?.kind === "reference"
          ? { numId: bound.numId, ilvl: bound.ilvl ?? 0 }
          : null,
      outlineLvl: styleOutlineLevel(style),
      alignment: alignmentOf(style.styleId),
      isDefault: style.default === true,
    };
  });
};

const inspectNumbering = (
  analyzed: AnalyzedDocument,
): DocxTemplateNumbering => {
  const { numbering } = analyzed;
  const indents = statedIndents(analyzed);
  return {
    abstractNums: (numbering?.definitions.abstractNums ?? []).map(
      ({ abstractNumId, levels }) => ({
        abstractNumId,
        levels: levels.map((level) => ({
          ilvl: level.ilvl,
          numFmt: level.numFmt,
          lvlText: level.lvlText,
          start: level.start ?? DEFAULT_LEVEL_START,
          suff: level.suffix ?? "tab",
          pStyle: level.pStyle ?? null,
          indent:
            indents.get(`${String(abstractNumId)}:${String(level.ilvl)}`) ??
            NO_INDENT,
        })),
      }),
    ),
    nums: (numbering?.definitions.nums ?? []).map(
      ({ numId, abstractNumId, levelOverrides }) => ({
        numId,
        abstractNumId,
        overrides: (levelOverrides ?? []).map(({ ilvl, startOverride }) => ({
          ilvl,
          startOverride: startOverride ?? null,
        })),
      }),
    ),
  };
};

const inspectSections = ({
  story,
  pkg,
}: AnalyzedDocument): DocxTemplateSection[] =>
  story.sections.map((sectPr, index) => {
    const properties = parseSectionProperties(sectPr);
    const partOf = (rId: string): string | null => {
      const relationship = pkg.documentRelationships.get(rId);
      return relationship === undefined
        ? null
        : pkg.resolveDocumentTarget(relationship.target);
    };
    return {
      index,
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
      headers: (properties.headerReferences ?? []).map(({ type, rId }) => ({
        type,
        part: partOf(rId),
      })),
      footers: (properties.footerReferences ?? []).map(({ type, rId }) => ({
        type,
        part: partOf(rId),
      })),
    };
  });

const fieldOf = (
  part: string,
  { instruction, kind, result }: StoryField,
): DocxTemplateField => ({
  part,
  instruction: instruction.trim(),
  kind,
  result,
});

/** The fields of a part other than the main document, read as one story. */
export const partFields = (part: RawPart): StoryField[] => {
  const state = createStoryState();
  const readUnder = (container: typeof part.root): void => {
    walkBlocks(container, {
      paragraph: (paragraph) => {
        readParagraph(paragraph, state);
      },
      tableStart: () => {},
      cellStart: () => {},
      tableEnd: () => {},
      unread: () => {},
      tracked: () => {},
    });
  };
  if (isW(part.root, "footnotes") || isW(part.root, "endnotes")) {
    for (const note of part.root.elements ?? []) {
      if (note.type === "element") {
        readUnder(note);
        // A field left open in one note does not reach into the next.
        endStory(state);
      }
    }
  } else {
    readUnder(part.root);
  }
  return state.fields;
};

const bookmarksOf = (part: RawPart): DocxTemplateBookmark[] => {
  const bookmarks: DocxTemplateBookmark[] = [];
  walkElements(part.root, (element) => {
    if (isW(element, "bookmarkStart")) {
      bookmarks.push({
        name: wAttr(element, "name") ?? "",
        id: wAttr(element, "id") ?? "",
        part: part.path,
      });
    }
  });
  return bookmarks;
};

/** The header and footer parts the main document names, in relationship order. */
export const headerFooterPaths = ({ pkg }: AnalyzedDocument): string[] =>
  [...pkg.documentRelationships.values()]
    .filter(
      ({ type, targetMode }) =>
        (type === RELATIONSHIP_TYPES.header ||
          type === RELATIONSHIP_TYPES.footer) &&
        targetMode !== "External",
    )
    .map(({ target }) => pkg.resolveDocumentTarget(target))
    .filter((path) => pkg.paths.has(path));

/**
 * The header and footer parts of a package, read. A part the document names
 * that has no readable root is a refusal: skipping it would hide its fields.
 */
export const readHeaderFooterParts = async (
  analyzed: AnalyzedDocument,
): Promise<Result<RawPart[], TemplatePlanError>> => {
  const parts: RawPart[] = [];
  for (const path of headerFooterPaths(analyzed)) {
    // oxlint-disable-next-line no-await-in-loop -- serialized reads keep the archive budget exact
    const xml = await readPartText(analyzed.pkg, path);
    if (Result.isError(xml)) {
      return xml;
    }
    const part = xml.value === null ? null : readRawPart(path, xml.value);
    if (part === null) {
      return Result.err(
        new TemplatePlanError({
          message: `The part ${path} could not be read.`,
          code: "xml-unreadable",
          part: path,
        }),
      );
    }
    parts.push(part);
  }
  return Result.ok(parts);
};

/** Build the inspection of an analysed template. */
export const inspectAnalyzedTemplate = async (
  analyzed: AnalyzedDocument,
): Promise<Result<DocxTemplateInspection, TemplatePlanError>> => {
  const { pkg, story, document, footnotes, endnotes, manifest } = analyzed;
  const fields: DocxTemplateField[] = story.state.fields.map((field) => {
    const inspected = fieldOf(pkg.documentPath, field);
    const block = story.paragraphs[field.paragraphOrdinal];
    if (block !== undefined) {
      inspected.blockId = block.topLevelId;
    }
    const endBlock = story.paragraphs[field.endParagraphOrdinal];
    if (endBlock !== undefined && field.closed) {
      inspected.endBlockId = endBlock.topLevelId;
    }
    return inspected;
  });
  const bookmarks = bookmarksOf(document);
  const headersAndFooters = await readHeaderFooterParts(analyzed);
  if (Result.isError(headersAndFooters)) {
    return headersAndFooters;
  }
  const others: RawPart[] = [...headersAndFooters.value];
  for (const notes of [footnotes, endnotes]) {
    if (notes.part !== null) {
      others.push(notes.part);
    }
  }
  for (const part of others) {
    fields.push(...partFields(part).map((field) => fieldOf(part.path, field)));
    bookmarks.push(...bookmarksOf(part));
  }
  return Result.ok({
    ...manifest,
    styles: inspectStyles(analyzed),
    numbering: inspectNumbering(analyzed),
    fields,
    sections: inspectSections(analyzed),
    bookmarks,
  });
};

/**
 * Read a template into what a plan may name: the block manifest plus its
 * styles, numbering definitions, fields, sections and bookmarks.
 */
export const inspectDocxTemplate = async (
  docx: Uint8Array,
): Promise<Result<DocxTemplateInspection, TemplatePlanError>> =>
  await guarded(async () => {
    const analyzed = await analyzeDocument(docx);
    if (Result.isError(analyzed)) {
      return Result.err(analyzed.error);
    }
    return await inspectAnalyzedTemplate(analyzed.value);
  });
