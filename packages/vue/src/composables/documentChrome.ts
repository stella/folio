import type {
  Document,
  SectionProperties,
  Style,
} from "@stll/folio-core/types/document";
import {
  getDocumentWatermark,
  type Watermark,
} from "@stll/folio-core/watermark";

/**
 * The document-wide values the toolbar, the rulers and the page setup controls
 * show. Derived from the document model by reading it, never from a document
 * built for the purpose, so deriving them costs nothing per keystroke.
 */
export type DocumentChrome = {
  sectionProperties: SectionProperties | null;
  watermark: Watermark | undefined;
  styles: Style[] | undefined;
};

export const documentChrome = (document: Document | null): DocumentChrome => {
  if (!document) {
    return { sectionProperties: null, watermark: undefined, styles: undefined };
  }
  const body = document.package.document;
  return {
    sectionProperties:
      body.finalSectionProperties ?? body.sections?.[0]?.properties ?? null,
    watermark: getDocumentWatermark(document),
    styles: document.package.styles?.styles,
  };
};
