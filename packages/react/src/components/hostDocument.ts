import { cloneDocumentWithParagraphPropertySources } from "@stll/folio-core/docx/document-clone";
import type { Document } from "@stll/folio-core/types/document";
import {
  readModel,
  type ReadonlyDocument,
} from "@stll/folio-core/types/documentModel";

/**
 * The host's own copy of the document the editor shows: the committed
 * document with the live body, numbering and notes in place of its own.
 *
 * Both inputs share objects with the editor's model, so neither is written:
 * the document is put together over them one level deep, and that whole is
 * copied once by the copy owner. The copy shares no object with either.
 */
export const copyHostDocument = (
  committed: ReadonlyDocument,
  live: ReadonlyDocument,
): Document => {
  const model = readModel(committed);
  const stories = readModel(live).package;
  const documentPackage = {
    ...model.package,
    document: {
      ...model.package.document,
      content: stories.document.content,
    },
  };
  if (stories.numbering) {
    documentPackage.numbering = stories.numbering;
  } else {
    delete documentPackage.numbering;
  }
  if (stories.footnotes) {
    documentPackage.footnotes = stories.footnotes;
  }
  if (stories.endnotes) {
    documentPackage.endnotes = stories.endnotes;
  }
  return cloneDocumentWithParagraphPropertySources({
    ...model,
    package: documentPackage,
  });
};
