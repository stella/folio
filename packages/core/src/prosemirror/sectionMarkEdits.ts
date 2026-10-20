/**
 * What an edit did to the paragraph marks that end sections.
 *
 * A section ends at the mark of the paragraph holding its record (ECMA-376
 * Part 1 §17.6.18). Deleting that mark deletes the break: Backspace or Delete
 * across it, a selection or a cut spanning it, a paste over it, a join. The
 * section's content then runs on into the next section, whose properties it
 * takes, which is what deleting the paragraph as a block does too.
 *
 * ProseMirror does not see marks: a replacement that joins two paragraphs
 * keeps the FIRST paragraph's node and attrs, so a joined paragraph would go on
 * holding the break whose mark was deleted, and would drop the one whose mark
 * survived. This reads each replacement of an edit for the marks it deleted,
 * and says which record every joined paragraph must hold: the record of the
 * paragraph whose mark now ends it.
 *
 * @packageDocumentation
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { Mapping, ReplaceStep } from "prosemirror-transform";

import type { SectionProperties } from "../types/document";
import { sectionPropertiesOf } from "./sectionCarrier";

export type SectionMarkEdits = {
  /** Section records whose paragraph mark a replacement deleted. */
  deletedMarks: ReadonlySet<SectionProperties>;
  /**
   * Paragraphs of the final document, by position, and the section record
   * each must hold because of the mark it now ends with.
   */
  carriers: readonly { position: number; record: SectionProperties | null }[];
};

const isParagraph = (node: PMNode): boolean => node.type.name === "paragraph";

/** Read `transactions`, applied in order and ending at `finalDoc`. */
export const sectionMarkEditsOf = (
  transactions: readonly Transaction[],
  finalDoc: PMNode,
): SectionMarkEdits => {
  const deletedMarks = new Set<SectionProperties>();
  const carriers: { position: number; record: SectionProperties | null }[] = [];
  const replacements: { step: ReplaceStep; doc: PMNode; mapIndex: number }[] = [];
  const mapping = new Mapping();
  for (const transaction of transactions) {
    for (const [index, step] of transaction.steps.entries()) {
      const doc = transaction.docs[index];
      if (step instanceof ReplaceStep && doc) {
        replacements.push({ step, doc, mapIndex: mapping.maps.length });
      }
      const map = transaction.mapping.maps[index];
      if (map) mapping.appendMap(map);
    }
  }

  const authoredRecordsByFinalParagraph = new Map<number, SectionProperties>();
  const authoredDoc = transactions.at(0)?.before;
  authoredDoc?.descendants((node, position) => {
    if (!isParagraph(node)) return true;
    const record = sectionPropertiesOf(node);
    if (record === null) return false;

    const mark = mapping.mapResult(position + node.nodeSize - 1, 1);
    if (mark.deleted) {
      deletedMarks.add(record);
      return false;
    }
    const $mark = finalDoc.resolve(mark.pos);
    if (isParagraph($mark.parent)) authoredRecordsByFinalParagraph.set($mark.before(), record);
    return false;
  });

  const addCarrier = (position: number): void => {
    carriers.push({ position, record: authoredRecordsByFinalParagraph.get(position) ?? null });
  };

  for (const { step, doc, mapIndex } of replacements) {
    const { from, to } = step;
    const after = mapping.slice(mapIndex);
    if (from === to) {
      const left = after.mapResult(from, -1);
      const right = after.mapResult(from, 1);
      if (left.deleted || right.deleted) continue;
      const $left = finalDoc.resolve(left.pos);
      const $right = finalDoc.resolve(right.pos);
      if (
        isParagraph($left.parent) &&
        isParagraph($right.parent) &&
        $left.start() !== $right.start()
      ) {
        // Splitting inserts a new paragraph mark before the authored mark. The
        // new leading paragraph does not end the section; the trailing half,
        // which still owns the authored mark, does.
        addCarrier($left.before());
        addCarrier($right.before());
      }
      continue;
    }
    const $from = doc.resolve(from);
    const $to = doc.resolve(to);
    if (!isParagraph($from.parent) || !isParagraph($to.parent) || $from.start() === $to.start()) {
      continue;
    }
    // The paragraph that ends with the right paragraph's mark holds its
    // record; a different paragraph left holding the first one's node ends
    // with a mark the replacement brought, and holds no break.
    const left = after.mapResult(from, -1);
    const right = after.mapResult(to, 1);
    if (right.deleted) continue;
    const $right = finalDoc.resolve(right.pos);
    if (!isParagraph($right.parent)) continue;
    addCarrier($right.before());
    if (left.deleted) continue;
    const $left = finalDoc.resolve(left.pos);
    if (isParagraph($left.parent) && $left.start() !== $right.start()) {
      addCarrier($left.before());
    }
  }
  return { deletedMarks, carriers };
};
