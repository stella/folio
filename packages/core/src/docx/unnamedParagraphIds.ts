/**
 * The id a paragraph the package names none for takes in the editor, carried
 * from the parser past the model.
 *
 * `ensureParaIds` writes an id into every story paragraph a package names
 * none for, keyed on the package itself, and a canonical session opens the
 * package it writes. The legacy editor opens the package as it is: its model
 * keeps holding no id there, so a save writes none, while the editor's
 * paragraph takes the id `ensureParaIds` writes. Either engine, opening the
 * same bytes, so shows the same block ids. See `packageUnnamedParagraphIds`
 * for the marking.
 */
import type { Node as PMNode } from "prosemirror-model";

import type { Paragraph } from "../types/document";
import type { XmlElement } from "./xmlParser";

/** The id of a part's unnamed paragraph, by the element that holds its place. */
const slotParagraphIds = new WeakMap<XmlElement, string>();

/**
 * Mark the element holding an unnamed paragraph's place, its `w:p` or the
 * cell whose implied paragraph it is, with the paragraph's id.
 */
export const markSlotParagraphId = (element: XmlElement, id: string): void => {
  slotParagraphIds.set(element, id);
};

const positionalParagraphIds = new WeakMap<Paragraph, string>();

/**
 * Hand a parsed paragraph the id its place in the package gives it: `element`
 * is its `w:p`, or the cell whose implied paragraph it is.
 */
export const readUnnamedParagraphId = (
  paragraph: Paragraph,
  element: XmlElement,
): void => {
  const id = slotParagraphIds.get(element);
  if (id !== undefined) positionalParagraphIds.set(paragraph, id);
};

/** Carry a paragraph's positional id to its copy. */
export const copyPositionalParagraphId = (
  target: Paragraph,
  source: Paragraph,
): void => {
  const id = positionalParagraphIds.get(source);
  if (id !== undefined) positionalParagraphIds.set(target, id);
};

/**
 * The positional id of an editor paragraph's parsed paragraph. The editor
 * mints it where it gives the paragraph an id, so converting the model to the
 * editor and back adds no id to the model.
 */
const prosePositionalParagraphIds = new WeakMap<PMNode, string>();

/** Link an editor paragraph to the positional id of the paragraph it shows. */
export const linkProsePositionalParagraphId = (
  node: PMNode,
  paragraph: Paragraph,
): void => {
  const id = positionalParagraphIds.get(paragraph);
  if (id !== undefined) prosePositionalParagraphIds.set(node, id);
};

/** Carry an editor paragraph's positional id across a node rebuild. */
export const copyProsePositionalParagraphId = (
  target: PMNode,
  source: PMNode,
): void => {
  const id = prosePositionalParagraphIds.get(source);
  if (id !== undefined) prosePositionalParagraphIds.set(target, id);
};

/** The id the package's position gives an editor paragraph, if it names none. */
export const prosePositionalParagraphId = (node: PMNode): string | undefined =>
  prosePositionalParagraphIds.get(node);
