import type { TextFormatting } from "../types/document";

/**
 * Exact target paragraph-mark formatting kept beside compare operations.
 * `undefined` is a meaningful value: the target has no direct `w:pPr/w:rPr`.
 *
 * @internal
 */
export type FolioParagraphMarkFormattingTemplates = {
  paired: ReadonlyMap<string, TextFormatting | undefined>;
  inserted: ReadonlyMap<string, TextFormatting | undefined>;
};

const paragraphMarkFormattingTemplatesByBatch = new WeakMap<
  object,
  FolioParagraphMarkFormattingTemplates
>();

/** Keep comparison-only formatting beside the JSON operation contract. */
export const recordParagraphMarkFormattingTemplates = (
  batch: object,
  templates: FolioParagraphMarkFormattingTemplates,
): void => {
  paragraphMarkFormattingTemplatesByBatch.set(batch, templates);
};

export const paragraphMarkFormattingTemplatesOf = (
  batch: object,
): FolioParagraphMarkFormattingTemplates | undefined =>
  paragraphMarkFormattingTemplatesByBatch.get(batch);
