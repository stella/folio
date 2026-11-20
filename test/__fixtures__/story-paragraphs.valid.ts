import type { Document } from "@stll/docx-core";
import * as ops from "@stll/docx-core/ops";
import { OP_STORIES, paragraphsThroughTextBoxes } from "@stll/docx-core/ops";

export const throughBoxes = (document: Document) =>
  paragraphsThroughTextBoxes(document, OP_STORIES.MAIN);

export const located = (document: Document, blockId: string) =>
  ops
    .paragraphsThroughTextBoxes(document, OP_STORIES.MAIN)
    .find(({ paragraph }) => paragraph.paraId === blockId);


export const locatedWithSiblings = (document: Document, blockId: string) =>
  ops.locateParagraphThroughTextBoxes(document, OP_STORIES.MAIN, blockId);
