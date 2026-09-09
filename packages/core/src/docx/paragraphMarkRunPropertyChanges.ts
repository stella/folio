import type { Paragraph, RunPropertyChange } from "../types/document";

/**
 * `w:pPr/w:rPrChange` belongs to the paragraph mark, not to the paragraph's
 * `w:pPrChange` collection and not to any visible run. Keep it private on the
 * Document model; ProseMirror has an explicit sparse attr for persistence.
 */
const changesByParagraph = new WeakMap<Paragraph, readonly RunPropertyChange[]>();

export const assignParagraphMarkRunPropertyChanges = (
  paragraph: Paragraph,
  changes: readonly RunPropertyChange[],
): void => {
  if (changes.length === 0) {
    changesByParagraph.delete(paragraph);
    return;
  }
  changesByParagraph.set(paragraph, [...changes]);
};

export const getParagraphMarkRunPropertyChanges = (
  paragraph: Paragraph,
): readonly RunPropertyChange[] | undefined => changesByParagraph.get(paragraph);

export const copyParagraphMarkRunPropertyChanges = (target: Paragraph, source: Paragraph): void => {
  const changes = changesByParagraph.get(source);
  if (changes !== undefined) {
    changesByParagraph.set(target, structuredClone(changes));
  }
};
