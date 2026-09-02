/**
 * Markdown → DOCX-document import — the inverse of {@link toMarkdown} and the
 * second half of the skills bridge. The parsing lives in `@stll/docx-core`
 * (`compileMarkdownToContent`), the same GFM reader the legal-source compiler
 * uses; this wrapper only places the blocks into an editor-ready `Document`.
 *
 * Round-trip notes:
 * - Lists arrive as real list paragraphs (`listRendering` plus `numPr`) with a
 *   matching `document.package.numbering`, so the editor shows a marker and
 *   {@link toMarkdown} re-derives `- ` / `1. ` rather than leaking a literal
 *   bullet glyph. Merging this content onto another document that has its own
 *   numbering (e.g. a styled preset) needs `mergeDocumentContent` to renumber
 *   the two numbering namespaces apart.
 * - Markdown carries no page geometry, so the section is flattened to a
 *   continuous, header/footer-free band (a skill body is a document, not a
 *   Word page). Headers/footers live outside `document.content` and are never
 *   produced here.
 */
import { compileMarkdownToContent } from "@stll/docx-core";

import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

// Markdown has no running header/footer, so flatten that band (content sits
// near the top, no inter-page header/footer gap). The page width and side
// margins stay at the default (Letter, 1" sides) so the editor fits to width
// exactly like the DOCX inspector — the body text fills the panel.
const applyMarkdownPageGeometry = (document: Document): void => {
  const section = document.package.document.finalSectionProperties;
  if (!section) {
    return;
  }
  // A markdown surface has no use for Word's 1-inch print margins; tighten all
  // four to a thin uniform gutter so body text fills the page, and zero the
  // header/footer distances since markdown has no running head/foot.
  section.marginTop = 480;
  section.marginBottom = 480;
  section.marginLeft = 480;
  section.marginRight = 480;
  section.headerDistance = 0;
  section.footerDistance = 0;
};

/**
 * Convert a markdown string to a parsed `Document`. Synchronous. The result is
 * ready to hand to the editor (`<DocxEditor document={…} />`) and to re-export
 * via {@link toMarkdown}.
 */
export function fromMarkdown(markdown: string): Document {
  const document = createEmptyDocument();
  const { content, numbering } = compileMarkdownToContent(markdown);
  if (content.length > 0) {
    document.package.document.content = content;
  }
  if (numbering) {
    document.package.numbering = numbering;
  }
  applyMarkdownPageGeometry(document);
  return document;
}
