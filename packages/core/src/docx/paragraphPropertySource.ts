import { panic } from "better-result";
import type { Fragment, Mark, Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import type { Document, Paragraph } from "../types/document";
import { visitDocxParagraphs } from "./paragraphTraversal";

type ParagraphPropertySource = {
  xml: string;
  formattingJson: string;
};

const paragraphPropertySources = new WeakMap<Paragraph, ParagraphPropertySource>();
const paragraphPropertySourceOwners = new WeakMap<Paragraph, Paragraph>();
const proseParagraphSourceOwners = new WeakMap<PMNode, Paragraph>();
const paragraphPropertySourceCandidates = new WeakMap<Paragraph, Paragraph>();
const paragraphPropertySourceTransferIds = new WeakMap<Paragraph, string>();

export const assignParagraphPropertySource = (
  paragraph: Paragraph,
  source: ParagraphPropertySource,
): void => {
  paragraphPropertySources.set(paragraph, source);
  paragraphPropertySourceOwners.set(paragraph, paragraph);
};

export const getParagraphPropertySource = (
  paragraph: Paragraph,
): ParagraphPropertySource | undefined => paragraphPropertySources.get(paragraph);

export const copyParagraphPropertySource = (target: Paragraph, source: Paragraph): void => {
  const propertySource = paragraphPropertySources.get(source);
  if (propertySource) {
    paragraphPropertySources.set(target, { ...propertySource });
    paragraphPropertySourceOwners.set(target, paragraphPropertySourceOwners.get(source) ?? source);
  }
  const transferId = paragraphPropertySourceTransferIds.get(source);
  if (transferId) {
    paragraphPropertySourceTransferIds.set(target, transferId);
  }
  const candidate = paragraphPropertySourceCandidates.get(source);
  if (candidate) {
    paragraphPropertySourceCandidates.set(target, candidate);
  }
};

type ParagraphCloneOverrides = Omit<Partial<Paragraph>, "type">;

/** Clone a paragraph while deliberately retaining its parsed `w:pPr` owner. */
export const cloneParagraphWithPropertySource = (
  paragraph: Paragraph,
  overrides: ParagraphCloneOverrides,
): Paragraph => {
  const cloned: Paragraph = { ...paragraph, ...overrides };
  copyParagraphPropertySource(cloned, paragraph);
  return cloned;
};

/** Clone a paragraph whose formatting provenance is deliberately no longer applicable. */
export const cloneParagraphWithoutPropertySource = (
  paragraph: Paragraph,
  overrides: ParagraphCloneOverrides,
): Paragraph => ({ ...paragraph, ...overrides });

const paragraphsIn = (document: Document): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  visitDocxParagraphs(
    {
      documentBody: document.package.document,
      headers: document.package.headers,
      footers: document.package.footers,
      footnotes: document.package.footnotes,
      endnotes: document.package.endnotes,
    },
    (paragraph) => paragraphs.push(paragraph),
  );
  return paragraphs;
};

/**
 * Deep-clone a document and transfer each private paragraph capture across the
 * exact graph clone. `structuredClone` preserves graph topology, so any count
 * mismatch is an internal invariant failure rather than a position heuristic.
 */
export const cloneDocumentWithParagraphPropertySources = (document: Document): Document => {
  const cloned = structuredClone(document);
  const sources = paragraphsIn(document);
  const targets = paragraphsIn(cloned);
  if (sources.length !== targets.length) {
    panic("The cloned document changed paragraph graph ownership.");
  }
  for (const [index, source] of sources.entries()) {
    const target = targets.at(index);
    if (!target) {
      panic("The cloned document lost a paragraph owner.");
    }
    copyParagraphPropertySource(target, source);
  }
  return cloned;
};

export const linkProseParagraphPropertySource = (
  proseParagraph: PMNode,
  sourceParagraph: Paragraph,
): void => {
  if (paragraphPropertySources.has(sourceParagraph)) {
    proseParagraphSourceOwners.set(
      proseParagraph,
      paragraphPropertySourceOwners.get(sourceParagraph) ?? sourceParagraph,
    );
  }
};

/** Carry a parser-linked paragraph owner across an immutable PM node rebuild. */
const copyProseParagraphPropertySource = (target: PMNode, source: PMNode): void => {
  if (target.type.name !== "paragraph" || source.type.name !== "paragraph") {
    return;
  }
  const sourceOwner = proseParagraphSourceOwners.get(source);
  if (sourceOwner) {
    proseParagraphSourceOwners.set(target, sourceOwner);
  }
};

type RecreateProseNodeOptions = {
  attrs?: PMNode["attrs"];
  content?: Fragment | PMNode | readonly PMNode[] | null;
  marks?: readonly Mark[];
};

/** Rebuild a PM node and retain paragraph provenance when the node is one. */
export const recreateProseNodeWithParagraphPropertySource = (
  source: PMNode,
  options: RecreateProseNodeOptions = {},
): PMNode => {
  const target = source.type.create(
    options.attrs ?? source.attrs,
    options.content === undefined ? source.content : options.content,
    options.marks ?? source.marks,
  );
  copyProseParagraphPropertySource(target, source);
  return target;
};

type SetProseParagraphMarkupOptions = {
  attrs: PMNode["attrs"];
  ownership: "preserve" | "transfer-allocated-id";
  pos: number;
  transaction: Transaction;
};

/** Replace paragraph markup without losing its private parser-owner link. */
export const setProseParagraphMarkupWithPropertySource = ({
  attrs,
  ownership,
  pos,
  transaction,
}: SetProseParagraphMarkupOptions): void => {
  const source = transaction.doc.nodeAt(pos);
  transaction.setNodeMarkup(pos, undefined, attrs);
  const target = transaction.doc.nodeAt(pos);
  if (!source || !target) {
    return;
  }
  if (ownership === "preserve") {
    copyProseParagraphPropertySource(target, source);
    return;
  }
  const paraId = target.attrs["paraId"];
  if (typeof paraId === "string") {
    transferProseParagraphPropertySource(target, source, paraId);
  }
};

/**
 * Carry the parser-owned source identity across load-time paraId allocation.
 * The generated id is safe to use later because it was assigned while the
 * ProseMirror node still had an unambiguous source paragraph owner.
 */
export const transferProseParagraphPropertySource = (
  target: PMNode,
  source: PMNode,
  paraId: string,
): void => {
  const sourceOwner = proseParagraphSourceOwners.get(source);
  if (!sourceOwner) {
    return;
  }
  proseParagraphSourceOwners.set(target, sourceOwner);
  paragraphPropertySourceTransferIds.set(sourceOwner, paraId);
};

export const linkParagraphPropertySourceCandidate = (target: Paragraph, source: PMNode): void => {
  const sourceOwner = proseParagraphSourceOwners.get(source);
  if (sourceOwner) {
    paragraphPropertySourceCandidates.set(target, sourceOwner);
  }
};

export const getParagraphPropertySourceCandidate = (paragraph: Paragraph): Paragraph | undefined =>
  paragraphPropertySourceCandidates.get(paragraph);

export const getParagraphPropertySourceTransferId = (paragraph: Paragraph): string | undefined =>
  paragraphPropertySourceTransferIds.get(paragraph);
