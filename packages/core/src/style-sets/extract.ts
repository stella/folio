import { panic } from "better-result";

import type { Document, FontInfo, Style } from "../types/document";
import {
  BUILT_IN_DEFAULT_PARAGRAPH_FORMATTING,
  BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID,
  BUILT_IN_DEFAULT_PARAGRAPH_STYLE_NAME,
  resolveDefaultParagraphStyle,
} from "../docx/defaultParagraphStyle";
import { getCachedNumberingMap } from "../docx/numberingParser";
import { isNumberingReference } from "../docx/numberingReference";
import { normalizeStyleNumberingReferences } from "../docx/numberingReferenceNormalization";
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
  const defaultParagraphStyleId = resolveDefaultParagraphStyle(styles)?.styleId;
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
  if (options.styleIds?.length === 0) {
    return panic("Cannot extract an empty style set");
  }

  const definitions = document.package.styles;
  const sourceStyles = definitions?.styles ?? [];
  const stylesById = new Map(sourceStyles.map((style) => [style.styleId, style]));
  const selectedStyleIds =
    options.styleIds === undefined
      ? new Set(stylesById.keys())
      : collectStyleDependencyClosure(stylesById, options.styleIds);
  const styles = structuredClone(
    sourceStyles.filter((style) => selectedStyleIds.has(style.styleId)),
  );
  // A style set can be minted from any Document, including one this package did
  // not parse, so the tolerance the parser applies is applied here too: a style
  // whose numbering the source never defined carries the "no numbering"
  // sentinel into the set instead of a reference nothing can resolve.
  normalizeStyleNumberingReferences({
    styles,
    numbering: document.package.numbering
      ? getCachedNumberingMap(document.package.numbering)
      : undefined,
  });

  const initialParagraphStyleId = ensureInitialParagraphStyle({
    styles,
    requested: options.initialParagraphStyleId,
    hasDocDefaults: definitions?.docDefaults !== undefined,
  });

  const numbering = extractReferencedNumbering(styles, document);
  const fontTable = sanitizeFontTable(document.package.fontTable);

  return structuredClone({
    version: DOCUMENT_STYLE_SET_VERSION,
    name: options.name,
    initialParagraphStyleId,
    styles: {
      ...(definitions?.docDefaults ? { docDefaults: definitions.docDefaults } : {}),
      ...(definitions?.latentStyles ? { latentStyles: definitions.latentStyles } : {}),
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

/**
 * The id of the set's initial paragraph style, minting the style if need be.
 *
 * A style the caller named must be in the set: naming one the selection
 * excluded is programmer misuse. Otherwise the source's own default is
 * resolved the way Word resolves it, and a source that declares none (a
 * generated package, a localized one, a package with no styles part at all)
 * gets a minted default appended to the set, because a set has to name a style
 * it contains.
 */
type EnsureInitialParagraphStyleOptions = {
  styles: Style[];
  requested: string | undefined;
  /** Whether the source declared `w:docDefaults`, which the set carries over. */
  hasDocDefaults: boolean;
};

const ensureInitialParagraphStyle = ({
  styles,
  requested,
  hasDocDefaults,
}: EnsureInitialParagraphStyleOptions): string => {
  if (requested !== undefined) {
    const named = styles.find((style) => style.styleId === requested && style.type === "paragraph");
    if (!named) {
      return panic(`Initial paragraph style "${requested}" is not present in the extracted set`);
    }
    return requested;
  }

  const resolved = resolveDefaultParagraphStyle(styles);
  if (resolved) {
    return resolved.styleId;
  }
  const minted = mintDefaultParagraphStyle({
    takenStyleIds: new Set(styles.map((style) => style.styleId)),
    hasDocDefaults,
  });
  styles.push(minted);
  return minted.styleId;
};

type MintDefaultParagraphStyleOptions = {
  takenStyleIds: ReadonlySet<string>;
  hasDocDefaults: boolean;
};

/**
 * The default paragraph style a set needs when its source declared none.
 *
 * The id only has to be free, because the set is what defines it. The
 * formatting has to be the built-in template's whenever the source had no
 * `w:docDefaults`, because that is what the source itself rendered as: a
 * consumer applies its built-in Normal only where no default paragraph style
 * exists, and this minted style is one. Where the source did declare
 * `w:docDefaults`, the set carries them and they remain authoritative, so the
 * minted style states nothing.
 */
const mintDefaultParagraphStyle = ({
  takenStyleIds,
  hasDocDefaults,
}: MintDefaultParagraphStyleOptions): Style => {
  let styleId = BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID;
  for (let suffix = 1; takenStyleIds.has(styleId); suffix += 1) {
    styleId = `${BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID}${suffix}`;
  }
  return {
    styleId,
    type: "paragraph",
    name: BUILT_IN_DEFAULT_PARAGRAPH_STYLE_NAME,
    default: true,
    ...(hasDocDefaults ? {} : { pPr: { ...BUILT_IN_DEFAULT_PARAGRAPH_FORMATTING } }),
  };
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
    if (isNumberingReference(numId)) {
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
