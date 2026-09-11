import { panic, TaggedError } from "better-result";
import type { Fragment, Mark, Node as PMNode, NodeType } from "prosemirror-model";
import { PluginKey, type Transaction } from "prosemirror-state";

import type { Document, Paragraph, TableCell } from "../types/document";
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
const paragraphPropertySourceTokens = new WeakMap<Paragraph, string>();
// Enumerable symbols follow ordinary immutable `{ ...document }` derivations,
// while JSON and other string-key serialization cannot expose the contract.
// `structuredClone` deliberately drops symbols, so the one sanctioned deep
// clone path transfers this value explicitly below.
const documentParagraphPropertySourceContract = Symbol("paragraphPropertySourceContract");

export const PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR = "_docxParagraphSourceToken";
export const PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR = "_docxParagraphSourceContract";

const PARAGRAPH_SOURCE_TOKEN_VERSION = "folio-ppr-v1";
const PARAGRAPH_SOURCE_TOKEN_PREFIX = "p1d";
const SHA256_HEX = /^[0-9a-f]{64}$/;
const SOURCE_TOKEN = /^p1d:([0-9a-f]{32}):(0|[1-9a-z][0-9a-z]*)$/;

export const PARAGRAPH_PROPERTY_SOURCE_VALIDATION_CODES = [
  "contract_mismatch",
  "duplicate_token",
  "invalid_token",
  "unknown_token",
] as const;

export type ParagraphPropertySourceValidationCode =
  (typeof PARAGRAPH_PROPERTY_SOURCE_VALIDATION_CODES)[number];

export class ParagraphPropertySourceValidationError extends TaggedError(
  "ParagraphPropertySourceValidationError",
)<{
  code: ParagraphPropertySourceValidationCode;
  message: string;
  token?: unknown;
}> {}

const contractForDigest = (sourceDigest: string): string => {
  if (!SHA256_HEX.test(sourceDigest)) {
    panic("Paragraph-property source digest must be lowercase SHA-256 hex");
  }
  return `${PARAGRAPH_SOURCE_TOKEN_VERSION}:${sourceDigest}`;
};

const fingerprintFromContract = (contract: string): string | null => {
  const prefix = `${PARAGRAPH_SOURCE_TOKEN_VERSION}:`;
  const digest = contract.startsWith(prefix) ? contract.slice(prefix.length) : "";
  return SHA256_HEX.test(digest) ? digest.slice(0, 32) : null;
};

const tokenForOrdinal = (contract: string, ordinal: number): string => {
  const fingerprint = fingerprintFromContract(contract);
  if (!fingerprint) {
    panic("Cannot mint a paragraph-property token for an invalid source contract");
  }
  return `${PARAGRAPH_SOURCE_TOKEN_PREFIX}:${fingerprint}:${ordinal.toString(36)}`;
};

const setDocumentParagraphPropertySourceContract = (document: Document, contract: string): void => {
  if (!fingerprintFromContract(contract)) {
    panic("Cannot attach an invalid paragraph-property source contract");
  }
  const existing = Object.hasOwn(document, documentParagraphPropertySourceContract)
    ? Reflect.get(document, documentParagraphPropertySourceContract)
    : undefined;
  if (existing === contract) {
    return;
  }
  if (
    !Reflect.defineProperty(document, documentParagraphPropertySourceContract, {
      enumerable: true,
      value: contract,
    })
  ) {
    panic("Cannot attach the paragraph-property source contract to the document");
  }
};

export const isParagraphPropertySourceToken = (token: unknown): token is string =>
  typeof token === "string" && SOURCE_TOKEN.test(token);

/** True only for a body-story token derived from the exact source contract. */
export const paragraphPropertySourceTokenMatchesContract = (
  token: unknown,
  contract: string,
): token is string => {
  if (!isParagraphPropertySourceToken(token)) {
    return false;
  }
  const match = SOURCE_TOKEN.exec(token);
  const fingerprint = fingerprintFromContract(contract);
  return match !== null && fingerprint !== null && match[1] === fingerprint;
};

/** Visit the v1 provenance scope in canonical OOXML body-story order. */
export const visitDocumentStoryParagraphs = (
  content: Document["package"]["document"]["content"],
  visit: (paragraph: Paragraph) => void,
): void => {
  visitDocxParagraphs({ documentBody: { content } }, visit);
};

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

/** Copy the captured `w:pPr` without claiming the source paragraph's durable identity. */
export const copyParagraphPropertyCapture = (target: Paragraph, source: Paragraph): void => {
  const propertySource = paragraphPropertySources.get(source);
  if (propertySource) {
    paragraphPropertySources.set(target, { ...propertySource });
    paragraphPropertySourceOwners.set(target, paragraphPropertySourceOwners.get(source) ?? source);
  }
};

export const copyParagraphPropertySource = (target: Paragraph, source: Paragraph): void => {
  copyParagraphPropertyCapture(target, source);
  const transferId = paragraphPropertySourceTransferIds.get(source);
  if (transferId) {
    paragraphPropertySourceTransferIds.set(target, transferId);
  }
  const candidate = paragraphPropertySourceCandidates.get(source);
  if (candidate) {
    paragraphPropertySourceCandidates.set(target, candidate);
  }
  const token = paragraphPropertySourceTokens.get(source);
  if (token) {
    paragraphPropertySourceTokens.set(target, token);
  }
};

/** Bind parsed body paragraphs to one exact source package. */
export const assignDocumentParagraphPropertySourceContract = (
  document: Document,
  sourceDigest: string,
): void => {
  const contract = contractForDigest(sourceDigest);
  setDocumentParagraphPropertySourceContract(document, contract);
  let ordinal = 0;
  // The traversal is part of the v1 durable identity contract. Any ordering
  // change requires a token-version bump and collaboration reseed.
  visitDocumentStoryParagraphs(document.package.document.content, (paragraph) => {
    paragraphPropertySourceTokens.set(paragraph, tokenForOrdinal(contract, ordinal));
    ordinal += 1;
  });
};

export const getDocumentParagraphPropertySourceContract = (
  document: Document,
): string | undefined => {
  if (!Object.hasOwn(document, documentParagraphPropertySourceContract)) {
    return undefined;
  }
  const contract = Reflect.get(document, documentParagraphPropertySourceContract);
  if (typeof contract !== "string" || !fingerprintFromContract(contract)) {
    panic("The document carries an invalid paragraph-property source contract");
  }
  return contract;
};

export const getProseDocumentParagraphPropertySourceContract = (
  document: PMNode,
): string | null => {
  const contract = document.attrs[PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR];
  return typeof contract === "string" && fingerprintFromContract(contract) ? contract : null;
};

export const getProseParagraphPropertySourceToken = (paragraph: PMNode): unknown =>
  paragraph.attrs[PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR];

export const copyDocumentParagraphPropertySourceContract = (
  target: Document,
  source: Document,
): void => {
  const contract = getDocumentParagraphPropertySourceContract(source);
  if (contract) {
    setDocumentParagraphPropertySourceContract(target, contract);
  }
};

export const getParagraphPropertySourceToken = (paragraph: Paragraph): string | undefined =>
  paragraphPropertySourceTokens.get(paragraph);

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
  copyDocumentParagraphPropertySourceContract(cloned, document);
  return cloned;
};

const paragraphsInTableCells = (cells: TableCell[]): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  for (const cell of cells) {
    visitDocxParagraphs({ documentBody: { content: cell.content } }, (paragraph) =>
      paragraphs.push(paragraph),
    );
  }
  return paragraphs;
};

/**
 * Clone package-crossing vertical-merge payloads with their captured `w:pPr`,
 * but without the durable paragraph tokens owned by the source package.
 */
export const cloneTableCellsWithParagraphPropertyCaptures = (cells: TableCell[]): TableCell[] => {
  const cloned = structuredClone(cells);
  const sources = paragraphsInTableCells(cells);
  const targets = paragraphsInTableCells(cloned);
  if (sources.length !== targets.length) {
    panic("The cloned table cells changed paragraph graph ownership.");
  }
  for (const [index, source] of sources.entries()) {
    const target = targets.at(index);
    if (!target) {
      panic("The cloned table cells lost a paragraph owner.");
    }
    copyParagraphPropertyCapture(target, source);
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

type CreateProseParagraphOptions = {
  attrs?: PMNode["attrs"];
  content?: Fragment | PMNode | readonly PMNode[] | null;
  marks?: readonly Mark[];
};

/** Create a parsed paragraph with both its same-process owner and durable body token. */
export const createProseParagraphWithPropertySource = (
  nodeType: NodeType | undefined,
  sourceParagraph: Paragraph,
  options: CreateProseParagraphOptions = {},
): PMNode => {
  if (!nodeType || nodeType.name !== "paragraph") {
    panic("Paragraph-property provenance can only seed a paragraph node");
  }
  const paragraph = nodeType.create(
    {
      ...options.attrs,
      [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]:
        paragraphPropertySourceTokens.get(sourceParagraph) ?? null,
    },
    options.content,
    options.marks,
  );
  linkProseParagraphPropertySource(paragraph, sourceParagraph);
  return paragraph;
};

type ParagraphPropertySourceTransfer = {
  displacedToken: string | null;
  selectedToken: string | null;
};

const paragraphPropertySourceTransfersKey = new PluginKey<
  readonly ParagraphPropertySourceTransfer[]
>("paragraphPropertySourceTransfers");

type JoinProseParagraphsWithRightPropertySourceOptions = {
  attrs: PMNode["attrs"];
  pos: number;
  transaction: Transaction;
};

/** Join adjacent paragraphs when revision semantics deliberately select the right owner. */
export const joinProseParagraphsWithRightPropertySource = ({
  attrs,
  pos,
  transaction,
}: JoinProseParagraphsWithRightPropertySourceOptions): Transaction => {
  const $pos = transaction.doc.resolve(pos);
  const left = $pos.nodeBefore;
  const right = $pos.nodeAfter;
  if (!left || !right || left.type.name !== "paragraph" || right.type.name !== "paragraph") {
    panic("Paragraph-property ownership can only join adjacent paragraphs");
  }
  const leftToken = getProseParagraphPropertySourceToken(left);
  const rightToken = getProseParagraphPropertySourceToken(right);
  const leftPos = pos - left.nodeSize;
  transaction.join(pos);
  const joined = transaction.doc.nodeAt(leftPos);
  if (!joined || joined.type.name !== "paragraph") {
    panic("Joining paragraphs did not produce a paragraph");
  }
  transaction.setNodeMarkup(
    leftPos,
    undefined,
    { ...attrs, [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: rightToken ?? null },
    joined.marks,
  );
  const transfers = transaction.getMeta(paragraphPropertySourceTransfersKey) ?? [];
  transaction.setMeta(paragraphPropertySourceTransfersKey, [
    ...transfers,
    {
      displacedToken: typeof leftToken === "string" ? leftToken : null,
      selectedToken: typeof rightToken === "string" ? rightToken : null,
    },
  ]);
  return transaction;
};

export const getExplicitParagraphPropertySourceTransfers = (
  transaction: Transaction,
): readonly ParagraphPropertySourceTransfer[] =>
  transaction.getMeta(paragraphPropertySourceTransfersKey) ?? [];

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

/**
 * Rebuild a PM node that is crossing from another package: retain its
 * same-process paragraph-property capture, but detach the durable token that
 * can only name a paragraph in the package it came from.
 */
export const recreateProseNodeWithDetachedParagraphPropertySource = (
  source: PMNode,
  options: RecreateProseNodeOptions = {},
): PMNode => {
  const attrs = options.attrs ?? source.attrs;
  return recreateProseNodeWithParagraphPropertySource(source, {
    ...options,
    attrs:
      source.type.name === "paragraph"
        ? { ...attrs, [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: null }
        : attrs,
  });
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
  const token = source.attrs[PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR];
  if (typeof token === "string") {
    paragraphPropertySourceTokens.set(target, token);
  }
  const sourceOwner = proseParagraphSourceOwners.get(source);
  if (sourceOwner) {
    paragraphPropertySourceCandidates.set(target, sourceOwner);
  }
};

export const getParagraphPropertySourceCandidate = (paragraph: Paragraph): Paragraph | undefined =>
  paragraphPropertySourceCandidates.get(paragraph);

export const getParagraphPropertySourceTransferId = (paragraph: Paragraph): string | undefined =>
  paragraphPropertySourceTransferIds.get(paragraph);
