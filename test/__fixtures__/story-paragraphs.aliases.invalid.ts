import type { Document } from "@stll/docx-core";
import * as ops from "@stll/docx-core/ops";

export const namespaced = (document: Document) =>
  ops.storyParagraphs(document.package.document);

export const computed = async (document: Document) =>
  (await import("@stll/docx-core/ops"))["storyParagraphs"](
    document.package.document,
  );

export const destructured = (document: Document) => {
  const { storyParagraphs: own } = ops;
  return own(document.package.document);
};

export const loaded = async (document: Document) =>
  (await import("@stll/docx-core/ops")).storyParagraphs(
    document.package.document,
  );

export { storyParagraphs } from "@stll/docx-core/ops";
