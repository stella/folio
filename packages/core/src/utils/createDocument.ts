/**
 * Create Document Utility
 *
 * Provides functions to create new documents programmatically.
 */

import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";
import { panic } from "better-result";

import { BUILT_IN_STYLE_NAME, builtInHeadingStyleName } from "../docx/builtInStyles";

import type {
  Document,
  DocxPackage,
  DocumentBody,
  Paragraph,
  Run,
  TextContent,
  SectionProperties,
  Style,
} from "../types/document";
import { BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID } from "../docx/defaultParagraphStyle";
import { createParseWarningCollector } from "../docx/parseContext";
import { formatParseWarnings } from "../docx/parseWarningMessage";
import { normalizeDocumentStyleSet } from "../style-sets/styleSetNormalization";
import {
  DOCUMENT_PRESET_VERSION,
  type DocumentPreset,
  type DocumentStyleSet,
} from "../style-sets/types";

// ============================================================================
// DEFAULT STYLES
// ============================================================================

/**
 * Get default paragraph styles (matching Google Docs defaults)
 *
 * Font sizes are in half-points (e.g., 22 = 11pt, 40 = 20pt)
 * Colors are RGB hex without # prefix
 *
 * Each style carries the built-in `w:name` it stands for, which is how every
 * consumer recognises it (`docx/builtInStyles.ts`), and the headings carry
 * `w:outlineLvl` so a Word TOC field and the navigation pane list them.
 */
function getDefaultStyles(): Style[] {
  return [
    // Normal - base style for body text (11pt Arial)
    {
      styleId: "Normal",
      type: "paragraph",
      name: BUILT_IN_STYLE_NAME.normal,
      default: true,
      qFormat: true,
      uiPriority: 0,
      rPr: {
        fontSize: 22, // 11pt
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        lineSpacing: 276, // 1.15 spacing
      },
    },
    // Title - document title (26pt, bold)
    {
      styleId: "Title",
      type: "paragraph",
      name: BUILT_IN_STYLE_NAME.title,
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 10,
      rPr: {
        fontSize: 52, // 26pt
        bold: true,
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        lineSpacing: 240, // Single spacing
      },
    },
    // Subtitle (15pt, gray)
    {
      styleId: "Subtitle",
      type: "paragraph",
      name: BUILT_IN_STYLE_NAME.subtitle,
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 11,
      rPr: {
        fontSize: 30, // 15pt
        color: { rgb: "666666" }, // Gray
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        lineSpacing: 240,
      },
    },
    // Heading 1 (20pt, bold)
    {
      styleId: "Heading1",
      type: "paragraph",
      name: builtInHeadingStyleName(0),
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      rPr: {
        fontSize: 40, // 20pt
        bold: true,
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        spaceBefore: 400, // 20pt before
        spaceAfter: 120, // 6pt after
        lineSpacing: 240,
        outlineLevel: 0,
      },
    },
    // Heading 2 (16pt, bold)
    {
      styleId: "Heading2",
      type: "paragraph",
      name: builtInHeadingStyleName(1),
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      rPr: {
        fontSize: 32, // 16pt
        bold: true,
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        spaceBefore: 360, // 18pt before
        spaceAfter: 80, // 4pt after
        lineSpacing: 240,
        outlineLevel: 1,
      },
    },
    // Heading 3 (14pt, bold)
    {
      styleId: "Heading3",
      type: "paragraph",
      name: builtInHeadingStyleName(2),
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      rPr: {
        fontSize: 28, // 14pt
        bold: true,
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        spaceBefore: 320, // 16pt before
        spaceAfter: 80, // 4pt after
        lineSpacing: 240,
        outlineLevel: 2,
      },
    },
    // Heading 4 (12pt, bold)
    {
      styleId: "Heading4",
      type: "paragraph",
      name: builtInHeadingStyleName(3),
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      rPr: {
        fontSize: 24, // 12pt
        bold: true,
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        spaceBefore: 280, // 14pt before
        spaceAfter: 80, // 4pt after
        lineSpacing: 240,
        outlineLevel: 3,
      },
    },
    // Heading 5 and 6 complete the range `docx/server/build.ts` accepts
    // (`HEADING_LEVELS`); without them a level-5 heading referenced a style
    // this set never defined.
    {
      styleId: "Heading5",
      type: "paragraph",
      name: builtInHeadingStyleName(4),
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      rPr: {
        fontSize: 22, // 11pt
        bold: true,
        fontFamily: { ascii: "Arial", hAnsi: "Arial" },
      },
      pPr: {
        spaceBefore: 240, // 12pt before
        spaceAfter: 80, // 4pt after
        lineSpacing: 240,
        outlineLevel: 4,
      },
    },
    {
      styleId: "Heading6",
      type: "paragraph",
      name: builtInHeadingStyleName(5),
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      // Distinguished from Heading 5 by colour rather than italics: an italic
      // face in the default style set would make the layout engine preload a
      // font variant no plain document uses.
      rPr: {
        fontSize: 22, // 11pt
        bold: true,
        color: { rgb: "595959" },
        fontFamily: { ascii: "Arial", hAnsi: "Arial" },
      },
      pPr: {
        spaceBefore: 240, // 12pt before
        spaceAfter: 80, // 4pt after
        lineSpacing: 240,
        outlineLevel: 5,
      },
    },
    // The table style `docx/server/build.ts` applies to every table it builds.
    {
      styleId: "TableNormal",
      type: "table",
      name: BUILT_IN_STYLE_NAME.normalTable,
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
    },
    {
      styleId: "TableGrid",
      type: "table",
      name: BUILT_IN_STYLE_NAME.tableGrid,
      basedOn: "TableNormal",
      uiPriority: 39,
      tblPr: {
        borders: {
          top: { style: "single", size: 4, space: 0 },
          bottom: { style: "single", size: 4, space: 0 },
          left: { style: "single", size: 4, space: 0 },
          right: { style: "single", size: 4, space: 0 },
          insideH: { style: "single", size: 4, space: 0 },
          insideV: { style: "single", size: 4, space: 0 },
        },
      },
    },
    // Quote — the built-in a markdown blockquote compiles to
    // (`compileMarkdownToContent`). Without a definition here that paragraph
    // carries a `w:pStyle` pointing at nothing.
    {
      styleId: "Quote",
      type: "paragraph",
      name: BUILT_IN_STYLE_NAME.quote,
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 29,
      rPr: {
        italic: true,
        color: { rgb: "404040" },
      },
      pPr: {
        indentLeft: 720, // 0.5" — Word's built-in quote indent
        indentRight: 720,
        spaceBefore: 160,
        spaceAfter: 160,
        lineSpacing: 240,
      },
    },
  ];
}

// ============================================================================
// DEFAULT SECTION PROPERTIES
// ============================================================================

/**
 * Get default section properties (US Letter, 1 inch margins)
 */
function getDefaultSectionProperties(): SectionProperties {
  return {
    pageWidth: 12_240, // 8.5 inches in twips
    pageHeight: 15_840, // 11 inches in twips
    orientation: "portrait",
    marginTop: 1440, // 1 inch
    marginBottom: 1440,
    marginLeft: 1440,
    marginRight: 1440,
    headerDistance: 720, // 0.5 inch
    footerDistance: 720,
    gutter: 0,
    columnCount: 1,
    columnSpace: 720,
    equalWidth: true,
    sectionStart: "nextPage",
    verticalAlign: "top",
  };
}

// ============================================================================
// EMPTY DOCUMENT
// ============================================================================

/**
 * Options for creating an empty document
 */
type CreateEmptyDocumentBaseOptions = {
  /** Page width in twips (default: 12240 = 8.5 inches) */
  pageWidth?: number;
  /** Page height in twips (default: 15840 = 11 inches) */
  pageHeight?: number;
  /** Page orientation (default: 'portrait') */
  orientation?: "portrait" | "landscape";
  /** Top margin in twips (default: 1440 = 1 inch) */
  marginTop?: number;
  /** Bottom margin in twips (default: 1440 = 1 inch) */
  marginBottom?: number;
  /** Left margin in twips (default: 1440 = 1 inch) */
  marginLeft?: number;
  /** Right margin in twips (default: 1440 = 1 inch) */
  marginRight?: number;
  /** Initial text content (default: empty string) */
  initialText?: string;
};

type CreateEmptyDocumentStyleOptions =
  | {
      /** Reusable formatting resources for the new document. */
      styleSet?: DocumentStyleSet;
      preset?: never;
    }
  | {
      /** Style resources plus page-level defaults for the new document. */
      preset: DocumentPreset;
      styleSet?: never;
    };

export type CreateEmptyDocumentOptions = CreateEmptyDocumentBaseOptions &
  CreateEmptyDocumentStyleOptions;

type CreateDocumentWithTextOptions = Omit<CreateEmptyDocumentBaseOptions, "initialText"> &
  CreateEmptyDocumentStyleOptions;

/**
 * Create an empty document with a single paragraph
 *
 * @param options - Optional configuration for the document
 * @returns A new empty Document object
 *
 * @example
 * ```ts
 * // Create a blank document
 * const doc = createEmptyDocument();
 *
 * // Create with custom margins
 * const doc = createEmptyDocument({
 *   marginTop: 720,  // 0.5 inch
 *   marginBottom: 720,
 * });
 *
 * // Create with initial text
 * const doc = createEmptyDocument({
 *   initialText: 'Hello, World!'
 * });
 * ```
 */
export function createEmptyDocument(options: CreateEmptyDocumentOptions = {}): Document {
  const suppliedStyleSet = options.preset?.styleSet ?? options.styleSet;
  if (options.preset?.version !== undefined && options.preset.version !== DOCUMENT_PRESET_VERSION) {
    return panic(`Unsupported document preset version: ${options.preset.version}`);
  }
  // A style set is portable and may have been persisted before folio learned
  // to repair one, so it enters through the same normalisation a `.docx` does
  // rather than being trusted or asserted about.
  const { context: styleSetContext, warnings: styleSetWarnings } = createParseWarningCollector();
  const styleSet =
    suppliedStyleSet === undefined
      ? undefined
      : normalizeDocumentStyleSet(suppliedStyleSet, styleSetContext);

  const sectionProps = structuredClone(
    options.preset?.sectionProperties ?? getDefaultSectionProperties(),
  );

  // Twips are integer-typed in OOXML, and callers typically compute them as
  // `inches * 1440` which produces drift like `0.7 * 1440 === 1008.0000000000001`.
  // Round at the API boundary so the model never carries a fractional twip.
  if (options.pageWidth !== undefined) {
    sectionProps.pageWidth = Math.round(options.pageWidth);
  }
  if (options.pageHeight !== undefined) {
    sectionProps.pageHeight = Math.round(options.pageHeight);
  }
  if (options.orientation !== undefined) {
    sectionProps.orientation = options.orientation;
  }
  if (options.marginTop !== undefined) {
    sectionProps.marginTop = Math.round(options.marginTop);
  }
  if (options.marginBottom !== undefined) {
    sectionProps.marginBottom = Math.round(options.marginBottom);
  }
  if (options.marginLeft !== undefined) {
    sectionProps.marginLeft = Math.round(options.marginLeft);
  }
  if (options.marginRight !== undefined) {
    sectionProps.marginRight = Math.round(options.marginRight);
  }

  // Create initial paragraph content
  const initialText = options.initialText ?? "";
  const textContent: TextContent = {
    type: "text",
    text: initialText,
  };

  const run: Run = {
    type: "run",
    content: initialText ? [textContent] : [],
    formatting: {},
  };

  const paragraph: Paragraph = {
    type: "paragraph",
    content: [run],
    formatting: {
      styleId: styleSet?.initialParagraphStyleId ?? BUILT_IN_DEFAULT_PARAGRAPH_STYLE_ID,
    },
  };

  // Create document body
  const documentBody: DocumentBody = {
    content: [paragraph],
    finalSectionProperties: sectionProps,
  };

  // Create package with default styles
  const defaultStyleDefinitions = {
    docDefaults: {
      rPr: {
        fontSize: 22,
        fontFamily: {
          ascii: "Arial",
          hAnsi: "Arial",
        },
      },
      pPr: {
        lineSpacing: 276,
      },
    },
    styles: getDefaultStyles(),
  };
  const docxPackage: DocxPackage = {
    conformanceClass: DOCX_CONFORMANCE_CLASSES.TRANSITIONAL,
    document: documentBody,
    relationships: new Map(),
    styles: structuredClone(styleSet?.styles ?? defaultStyleDefinitions),
  };
  if (styleSet?.numbering) {
    docxPackage.numbering = structuredClone(styleSet.numbering);
  }
  if (styleSet?.theme) {
    docxPackage.theme = structuredClone(styleSet.theme);
  }
  if (styleSet?.fontTable) {
    docxPackage.fontTable = structuredClone(styleSet.fontTable);
  }
  if (styleSet?.settings) {
    docxPackage.settings = structuredClone(styleSet.settings);
  }

  // Create document
  const parseWarnings = styleSetWarnings();
  const document: Document = {
    package: docxPackage,
    templateVariables: [],
    warnings: formatParseWarnings(parseWarnings),
  };
  if (parseWarnings.length > 0) {
    document.parseWarnings = parseWarnings;
  }

  return document;
}

/**
 * Create a document with a single paragraph containing the given text
 *
 * @param text - The text content for the document
 * @param options - Optional configuration for the document
 * @returns A new Document object with the specified text
 */
export function createDocumentWithText(
  text: string,
  options: CreateDocumentWithTextOptions = {},
): Document {
  return createEmptyDocument({ ...options, initialText: text });
}
