import type { ParagraphBlock } from "./types";

/** Whether a paragraph has no visible run content. */
export function isEmptyParagraph(block: ParagraphBlock): boolean {
  if (block.runs.length === 0) {
    return true;
  }
  if (block.runs.length !== 1) {
    return false;
  }
  const run = block.runs.at(0);
  return run?.kind === "text" && run.text === "";
}

/**
 * Resolve leading paragraph spacing after applying the empty-paragraph
 * inherited-spacing collapse rule.
 */
export function getParagraphSpacingBefore(block: ParagraphBlock): number {
  const value = block.attrs?.spacing?.before ?? 0;
  if (value === 0) {
    return 0;
  }
  if (
    isEmptyParagraph(block) &&
    !block.attrs?.styleId &&
    !block.attrs?.hasDirectParagraphFormatting &&
    !block.attrs?.spacingExplicit?.before
  ) {
    return 0;
  }
  return value;
}

/**
 * Resolve trailing paragraph spacing after applying the empty-paragraph
 * inherited-spacing collapse rule.
 */
export function getParagraphSpacingAfter(block: ParagraphBlock): number {
  const value = block.attrs?.spacing?.after ?? 0;
  if (value === 0) {
    return 0;
  }
  if (
    isEmptyParagraph(block) &&
    !block.attrs?.styleId &&
    !block.attrs?.hasDirectParagraphFormatting &&
    !block.attrs?.spacingExplicit?.after
  ) {
    return 0;
  }
  return value;
}
