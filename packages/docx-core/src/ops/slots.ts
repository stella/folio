/**
 * The ids records carry of their own.
 *
 * A run with a tracked property change (`w:rPrChange`), a tracked insertion,
 * deletion or move, a content control with a `w:id`, and a paragraph with a
 * tracked property or mark change are identified records. Each id is a slot:
 * its space (revision or content control) and its value, listed in a fixed
 * order per record so two records with the same fields list the same slots.
 */

import type { Paragraph, ParagraphPropertyChange, RunPropertyChange } from "../model/document";
import { IDENTITY_SPACES, type IdentitySlot } from "./ids";
import type { InlineNode } from "./leaves";

type IdentifiedRecord = Paragraph | InlineNode;

const revision = (id: number): IdentitySlot => ({ space: IDENTITY_SPACES.REVISION, id });

/** The ids a record carries itself, in a fixed order; its children's are their own. */
export const identitySlots = (record: IdentifiedRecord): IdentitySlot[] => {
  switch (record.type) {
    case "paragraph": {
      const out = (record.propertyChanges ?? []).map((change) => revision(change.info.id));
      if (record.pPrMark !== undefined) out.push(revision(record.pPrMark.info.id));
      return out;
    }
    case "run":
      return (record.propertyChanges ?? []).map((change) => revision(change.info.id));
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
      return [revision(record.info.id)];
    case "inlineSdt":
      return record.properties.id === undefined
        ? []
        : [{ space: IDENTITY_SPACES.CONTROL, id: record.properties.id }];
    default:
      return [];
  }
};

const withRunChangeIds = (
  changes: readonly RunPropertyChange[],
  ids: readonly number[],
): RunPropertyChange[] => {
  const out: RunPropertyChange[] = [];
  for (const [index, change] of changes.entries()) {
    out.push({ ...change, info: { ...change.info, id: ids[index] ?? change.info.id } });
  }
  return out;
};

const withParagraphChangeIds = (
  changes: readonly ParagraphPropertyChange[],
  ids: readonly number[],
): ParagraphPropertyChange[] => {
  const out: ParagraphPropertyChange[] = [];
  for (const [index, change] of changes.entries()) {
    out.push({ ...change, info: { ...change.info, id: ids[index] ?? change.info.id } });
  }
  return out;
};

/** The same inline record carrying other ids, slot for slot. */
export const withInlineIdentity = (node: InlineNode, ids: readonly number[]): InlineNode => {
  switch (node.type) {
    case "run":
      return node.propertyChanges === undefined
        ? node
        : {
            ...node,
            propertyChanges: withRunChangeIds(node.propertyChanges, ids),
          };
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
      return { ...node, info: { ...node.info, id: ids[0] ?? node.info.id } };
    case "inlineSdt":
      return node.properties.id === undefined
        ? node
        : { ...node, properties: { ...node.properties, id: ids[0] ?? node.properties.id } };
    default:
      return node;
  }
};

/** The same paragraph carrying other ids on its own tracked changes, slot for slot. */
export const withParagraphIdentity = (paragraph: Paragraph, ids: readonly number[]): Paragraph => {
  const next: Paragraph = { ...paragraph };
  const changes = paragraph.propertyChanges;
  if (changes !== undefined) {
    next.propertyChanges = withParagraphChangeIds(changes, ids);
  }
  if (paragraph.pPrMark !== undefined) {
    const id = ids[changes?.length ?? 0] ?? paragraph.pPrMark.info.id;
    next.pPrMark = { ...paragraph.pPrMark, info: { ...paragraph.pPrMark.info, id } };
  }
  return next;
};

/** A record with every id it carries set to zero: the form two records are compared in. */
export const maskIdentity = (node: InlineNode): InlineNode => {
  const slots = identitySlots(node);
  return slots.length === 0
    ? node
    : withInlineIdentity(
        node,
        slots.map(() => 0),
      );
};
