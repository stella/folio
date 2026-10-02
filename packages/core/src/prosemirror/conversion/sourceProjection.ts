import type { Node as PMNode } from "prosemirror-model";

import { createModelVersionTracker } from "../../utils/modelVersion";
import type { Document, Paragraph } from "../../types/document";
import {
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySourceToken,
  visitDocumentStoryParagraphs,
} from "../../docx/paragraphPropertySource";

/** The model inputs consumed by the main-story projection. */
export const documentProjectionInput = (document: Document) => ({
  content: document.package.document.content,
  styles: document.package.styles,
  theme: document.package.theme ?? null,
  finalSectionStart: document.package.document.sections?.at(-1)?.properties.sectionStart ?? null,
  adjustLineHeightInTable: document.package.settings?.adjustLineHeightInTable === true,
  doNotUseIndentAsNumberingTabStop:
    document.package.settings?.doNotUseIndentAsNumberingTabStop === true,
  contract: getDocumentParagraphPropertySourceContract(document) ?? null,
});

const readVersion = createModelVersionTracker();

type SourceProjection = {
  input: ReturnType<typeof documentProjectionInput>;
  version: unknown;
  tokens: Map<Paragraph, string | undefined>;
  projection: PMNode;
};

const projections = new WeakMap<Document, SourceProjection>();

export const rememberSourceProjection = (document: Document, projection: PMNode): void => {
  const tokens = new Map<Paragraph, string | undefined>();
  visitDocumentStoryParagraphs(document.package.document.content, (paragraph) => {
    tokens.set(paragraph, getParagraphPropertySourceToken(paragraph));
  });
  const input = documentProjectionInput(document);
  projections.set(document, { input, version: readVersion(input), tokens, projection });
};

export const currentSourceProjection = (document: Document): PMNode | undefined => {
  const cached = projections.get(document);
  if (!cached) return undefined;
  Object.assign(cached.input, documentProjectionInput(document));
  if (!Object.is(readVersion(cached.input), cached.version)) return undefined;
  for (const [paragraph, token] of cached.tokens) {
    if (getParagraphPropertySourceToken(paragraph) !== token) return undefined;
  }
  return cached.projection;
};
