import { panic } from "better-result";

import type { Document, FontInfo, Style } from "../types/document";
import { parseDocx } from "../docx/parser";
import type { DocxInput } from "../utils/docxInput";
import { DOCUMENT_STYLE_SET_VERSION, type DocumentStyleSet } from "./types";

export type ExtractDocumentStyleSetOptions = {
  name: string;
  /** Style IDs selected by the user. Omit to extract every style. */
  styleIds?: readonly string[];
  /** Defaults to the source document's default paragraph style. */
  initialParagraphStyleId?: string;
};

export type DocumentStyleCatalogEntry = {
  styleId: string;
  name: string;
  type: Style["type"];
  role: "default" | "quick" | "available" | "supporting";
  dependencies: string[];
  numberingId?: number;
};

export type DocumentStyleCatalog = {
  defaultParagraphStyleId?: string;
  styles: DocumentStyleCatalogEntry[];
};

export const inspectDocumentStyles = (document: Document): DocumentStyleCatalog => {
  const styles = document.package.styles?.styles ?? [];
  const defaultParagraphStyleId = styles.find(
    (style) => style.type === "paragraph" && style.default,
  )?.styleId;
  return {
    ...(defaultParagraphStyleId ? { defaultParagraphStyleId } : {}),
    styles: styles.map(toCatalogEntry),
  };
};

export const inspectDocumentStylesFromDocx = async (
  input: DocxInput,
): Promise<DocumentStyleCatalog> =>
  inspectDocumentStyles(
    await parseDocx(input, {
      preloadFonts: false,
      parseHeadersFooters: false,
      parseNotes: false,
      detectVariables: false,
    }),
  );

/**
 * Extract a sanitized style set from a parsed document.
 *
 * Selected styles expand to a dependency closure over basedOn, next, and link
 * references. Only numbering definitions referenced by the resulting styles
 * are retained. No source content or package relationship can enter the result.
 */
export const extractDocumentStyleSet = (
  document: Document,
  options: ExtractDocumentStyleSetOptions,
): DocumentStyleSet => {
  const definitions = document.package.styles;
  if (!definitions || definitions.styles.length === 0) {
    return panic("Cannot extract a style set from a document without styles");
  }

  const stylesById = new Map(definitions.styles.map((style) => [style.styleId, style]));
  const selectedStyleIds =
    options.styleIds === undefined
      ? new Set(stylesById.keys())
      : collectStyleDependencyClosure(stylesById, options.styleIds);
  const styles = definitions.styles.filter((style) => selectedStyleIds.has(style.styleId));
  if (styles.length === 0) {
    return panic("Cannot extract an empty style set");
  }

  const defaultParagraphStyle = styles.find((style) => style.type === "paragraph" && style.default);
  const initialParagraphStyleId =
    options.initialParagraphStyleId ?? defaultParagraphStyle?.styleId ?? "Normal";
  const initialParagraphStyle = styles.find(
    (style) => style.styleId === initialParagraphStyleId && style.type === "paragraph",
  );
  if (!initialParagraphStyle) {
    return panic(
      `Initial paragraph style "${initialParagraphStyleId}" is not present in the extracted set`,
    );
  }

  const numbering = extractReferencedNumbering(styles, document);
  const fontTable = sanitizeFontTable(document.package.fontTable);

  return structuredClone({
    version: DOCUMENT_STYLE_SET_VERSION,
    name: options.name,
    initialParagraphStyleId,
    styles: {
      ...(definitions.docDefaults ? { docDefaults: definitions.docDefaults } : {}),
      ...(definitions.latentStyles ? { latentStyles: definitions.latentStyles } : {}),
      styles,
    },
    ...(numbering ? { numbering } : {}),
    ...(document.package.theme ? { theme: document.package.theme } : {}),
    ...(fontTable ? { fontTable } : {}),
    ...(document.package.settings ? { settings: document.package.settings } : {}),
  });
};

export const extractDocumentStyleSetFromDocx = async (
  input: DocxInput,
  options: ExtractDocumentStyleSetOptions,
): Promise<DocumentStyleSet> => {
  const document = await parseDocx(input, {
    preloadFonts: false,
    parseHeadersFooters: false,
    parseNotes: false,
    detectVariables: false,
  });
  return extractDocumentStyleSet(document, options);
};

const collectStyleDependencyClosure = (
  stylesById: ReadonlyMap<string, Style>,
  requestedStyleIds: readonly string[],
): Set<string> => {
  const selected = new Set<string>();
  const pending = [...requestedStyleIds];

  while (pending.length > 0) {
    const styleId = pending.pop();
    if (styleId === undefined || selected.has(styleId)) {
      continue;
    }
    const style = stylesById.get(styleId);
    if (!style) {
      return panic(`Cannot extract unknown style "${styleId}"`);
    }
    selected.add(styleId);
    for (const dependency of [style.basedOn, style.next, style.link]) {
      if (dependency !== undefined && !selected.has(dependency)) {
        pending.push(dependency);
      }
    }
  }

  return selected;
};

const extractReferencedNumbering = (
  styles: readonly Style[],
  document: Document,
): Document["package"]["numbering"] => {
  const source = document.package.numbering;
  if (!source) {
    return undefined;
  }

  const referencedNumIds = new Set<number>();
  for (const style of styles) {
    const numId = style.pPr?.numPr?.numId;
    if (numId !== undefined) {
      referencedNumIds.add(numId);
    }
  }
  if (referencedNumIds.size === 0) {
    return undefined;
  }

  const nums = source.nums.filter((numbering) => referencedNumIds.has(numbering.numId));
  const referencedAbstractNumIds = new Set(nums.map((numbering) => numbering.abstractNumId));
  const abstractNums = source.abstractNums.filter((numbering) =>
    referencedAbstractNumIds.has(numbering.abstractNumId),
  );
  return { abstractNums, nums };
};

const sanitizeFontTable = (
  fontTable: Document["package"]["fontTable"],
): Document["package"]["fontTable"] => {
  if (!fontTable) {
    return undefined;
  }
  return {
    fonts: fontTable.fonts.map(stripEmbeddedFontRelationships),
  };
};

const stripEmbeddedFontRelationships = ({
  embedRegular: _embedRegular,
  embedBold: _embedBold,
  embedItalic: _embedItalic,
  embedBoldItalic: _embedBoldItalic,
  ...font
}: FontInfo): FontInfo => font;

const toCatalogEntry = (style: Style): DocumentStyleCatalogEntry => {
  const entry: DocumentStyleCatalogEntry = {
    styleId: style.styleId,
    name: style.name ?? style.styleId,
    type: style.type,
    role: styleRole(style),
    dependencies: [
      ...new Set(
        [style.basedOn, style.next, style.link].filter((dependency) => dependency !== undefined),
      ),
    ],
  };
  const numberingId = style.pPr?.numPr?.numId;
  if (numberingId !== undefined) {
    entry.numberingId = numberingId;
  }
  return entry;
};

const styleRole = (style: Style): DocumentStyleCatalogEntry["role"] => {
  if (style.default) {
    return "default";
  }
  if (style.qFormat) {
    return "quick";
  }
  if (!style.hidden && !style.semiHidden) {
    return "available";
  }
  return "supporting";
};
