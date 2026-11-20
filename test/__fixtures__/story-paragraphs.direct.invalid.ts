import type { Document } from "@stll/docx-core";
import {
  storyParagraphs,
  storyParagraphs as ownParagraphs,
} from "@stll/docx-core/ops";

export const direct = (document: Document) =>
  storyParagraphs(document.package.document);

export const renamed = (document: Document) =>
  ownParagraphs(document.package.document);
