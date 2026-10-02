import type { Node as PMNode } from "prosemirror-model";

import { getSourceReplayToken, type SourceReplayToken } from "@stll/docx-core/ops";
import type { Document } from "../../types/document";
import { getDocumentParagraphPropertySourceContract } from "../../docx/paragraphPropertySource";

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

type SourceProjection = {
  input: ReturnType<typeof documentProjectionInput>;
  token: SourceReplayToken;
  projection: PMNode;
};

const projections = new WeakMap<Document, SourceProjection>();

export const rememberSourceProjection = (document: Document, projection: PMNode): void => {
  const token = getSourceReplayToken(document);
  if (token === undefined) return;
  projections.set(document, { input: documentProjectionInput(document), token, projection });
};

export const currentSourceProjection = (document: Document): PMNode | undefined => {
  const cached = projections.get(document);
  if (!cached || getSourceReplayToken(document) !== cached.token) return undefined;
  const input = documentProjectionInput(document);
  // Tracked blocks and projection resources are immutable; only their root
  // references and the scalar document settings can change between projections.
  for (const [key, field] of Object.entries(input)) {
    if (!Object.is(field, Reflect.get(cached.input, key))) return undefined;
  }
  return cached.projection;
};
