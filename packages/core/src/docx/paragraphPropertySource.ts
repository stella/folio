import { panic, TaggedError } from "better-result";
import type { Fragment, Mark, Node as PMNode, NodeType } from "prosemirror-model";
import { PluginKey, type Transaction } from "prosemirror-state";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  RunContent,
  TableCell,
} from "../types/document";
import {
  PARAGRAPH_ID_STABILITY_ATTR,
  POSITIONAL_PARAGRAPH_ID_STABILITY,
} from "../prosemirror/paragraphProjectionAttrs";
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
const documentParagraphPropertySourceBindingBrand = Symbol(
  "documentParagraphPropertySourceBinding",
);
// Enumerable symbols follow ordinary immutable `{ ...document }` derivations,
// while JSON and other string-key serialization cannot expose the contract.
// `structuredClone` deliberately drops symbols, so the one sanctioned deep
// clone path transfers this value explicitly below.
const documentParagraphPropertySourceBinding = Symbol("paragraphPropertySourceBinding");

type DocumentParagraphPropertySourceBinding = Readonly<{
  [documentParagraphPropertySourceBindingBrand]: true;
  contract: string;
  sourceOwners: ReadonlySet<Paragraph>;
  sources: ReadonlyMap<string, Paragraph>;
}>;

export const PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR = "_docxParagraphSourceToken";
export const PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR = "_docxParagraphSourceContract";
/** Durable identity nested inside opaque collapsed-cell ProseMirror attrs; stripped on restore. */
export const TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR = "_docxParagraphSourceBinding";

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

export const TABLE_CELL_PARAGRAPH_SOURCE_PAYLOAD_ERROR_CLASSIFICATIONS = [
  "complexity_limit",
  "cyclic_or_aliased_graph",
  "depth_limit",
  "expected_array",
  "expected_block",
  "expected_cell",
  "expected_inline_content",
  "expected_paragraph",
  "expected_record",
  "expected_row",
  "invalid_binding",
  "invalid_block_type",
  "invalid_cell_type",
  "invalid_graph_value",
  "invalid_inline_content_type",
  "invalid_paragraph_type",
  "invalid_row_type",
  "malformed_source_token",
] as const;

export type TableCellParagraphSourcePayloadErrorClassification =
  (typeof TABLE_CELL_PARAGRAPH_SOURCE_PAYLOAD_ERROR_CLASSIFICATIONS)[number];

export class ParagraphPropertySourceValidationError extends TaggedError(
  "ParagraphPropertySourceValidationError",
)<{
  code: ParagraphPropertySourceValidationCode;
  classification?: TableCellParagraphSourcePayloadErrorClassification;
  message: string;
  path?: string;
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

const isDocumentParagraphPropertySourceBinding = (
  value: unknown,
): value is DocumentParagraphPropertySourceBinding =>
  typeof value === "object" &&
  value !== null &&
  documentParagraphPropertySourceBindingBrand in value &&
  value[documentParagraphPropertySourceBindingBrand] === true &&
  "contract" in value &&
  typeof value.contract === "string" &&
  "sourceOwners" in value &&
  value.sourceOwners instanceof Set &&
  "sources" in value &&
  value.sources instanceof Map;

const setDocumentParagraphPropertySourceBinding = (
  document: Document,
  binding: DocumentParagraphPropertySourceBinding,
): void => {
  if (!fingerprintFromContract(binding.contract)) {
    panic("Cannot attach an invalid paragraph-property source contract");
  }
  const existing: unknown = Object.hasOwn(document, documentParagraphPropertySourceBinding)
    ? Reflect.get(document, documentParagraphPropertySourceBinding)
    : undefined;
  if (existing === binding) {
    return;
  }
  if (
    !Reflect.defineProperty(document, documentParagraphPropertySourceBinding, {
      enumerable: true,
      value: binding,
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

type SynthesizedParagraphIdentity = Readonly<{
  type: "synthesized";
  paraId: string;
}>;

/**
 * Identity provenance exists only while a parsed Document and its live PM
 * projection share this process. It must not become a Document field: a
 * generated paraId is positional until it has actually been serialized and
 * parsed back from the package.
 */
const synthesizedIdentityByParagraph = new WeakMap<Paragraph, SynthesizedParagraphIdentity>();
const proseParagraphProjectionOwners = new WeakMap<PMNode, Paragraph>();

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
  const synthesizedIdentity = synthesizedIdentityByParagraph.get(source);
  if (synthesizedIdentity) {
    synthesizedIdentityByParagraph.set(target, synthesizedIdentity);
  }
};

/** Bind parsed body paragraphs to one exact source package. */
export const assignDocumentParagraphPropertySourceContract = (
  document: Document,
  sourceDigest: string,
): void => {
  const contract = contractForDigest(sourceDigest);
  const sources = new Map<string, Paragraph>();
  let ordinal = 0;
  // The traversal is part of the v1 durable identity contract. Any ordering
  // change requires a token-version bump and collaboration reseed.
  visitDocumentStoryParagraphs(document.package.document.content, (paragraph) => {
    const token = tokenForOrdinal(contract, ordinal);
    paragraphPropertySourceTokens.set(paragraph, token);
    sources.set(token, paragraph);
    ordinal += 1;
  });
  setDocumentParagraphPropertySourceBinding(
    document,
    Object.freeze({
      [documentParagraphPropertySourceBindingBrand]: true,
      contract,
      sourceOwners: new Set(sources.values()),
      sources,
    } satisfies DocumentParagraphPropertySourceBinding),
  );
};

const getDocumentParagraphPropertySourceBinding = (
  document: Document,
): DocumentParagraphPropertySourceBinding | undefined => {
  if (!Object.hasOwn(document, documentParagraphPropertySourceBinding)) {
    return undefined;
  }
  const binding: unknown = Reflect.get(document, documentParagraphPropertySourceBinding);
  if (
    !isDocumentParagraphPropertySourceBinding(binding) ||
    !fingerprintFromContract(binding.contract)
  ) {
    panic("The document carries an invalid paragraph-property source contract");
  }
  return binding;
};

export const getDocumentParagraphPropertySourceContract = (
  document: Document,
): string | undefined => getDocumentParagraphPropertySourceBinding(document)?.contract;

export const copyDocumentParagraphPropertySources = (
  document: Document,
): Map<string, Paragraph> | undefined => {
  const sources = getDocumentParagraphPropertySourceBinding(document)?.sources;
  return sources ? new Map(sources) : undefined;
};

export const paragraphPropertySourceBelongsToDocument = (
  paragraph: Paragraph,
  document: Document,
): boolean => {
  const owner = paragraphPropertySourceOwners.get(paragraph);
  const binding = getDocumentParagraphPropertySourceBinding(document);
  return owner !== undefined && binding !== undefined && binding.sourceOwners.has(owner);
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
  const binding = getDocumentParagraphPropertySourceBinding(source);
  if (binding) {
    setDocumentParagraphPropertySourceBinding(target, binding);
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

const paragraphsInTableCells = (cells: readonly TableCell[]): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  for (const cell of cells) {
    visitDocxParagraphs({ documentBody: { content: cell.content } }, (paragraph) =>
      paragraphs.push(paragraph),
    );
  }
  return paragraphs;
};

export type TableCellParagraphPropertySourceBinding =
  | Readonly<{ type: "authored" }>
  | Readonly<{ token: string; type: "source" }>;

type TableCellParagraphPropertySourceBindingInspection =
  | { status: "absent" }
  | { binding: Readonly<{ type: "authored" }>; status: "authored" }
  | { binding: Readonly<{ token: string; type: "source" }>; status: "source" }
  | { status: "invalid" };

const isPropertySourceBindingRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const inspectTableCellParagraphPropertySourceBinding = (
  paragraph: Paragraph,
): TableCellParagraphPropertySourceBindingInspection => {
  if (!Object.hasOwn(paragraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR)) {
    return { status: "absent" };
  }
  const value: unknown = Reflect.get(paragraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR);
  if (!isPropertySourceBindingRecord(value)) {
    return { status: "invalid" };
  }
  const keys = Reflect.ownKeys(value);
  if (
    keys.length === 1 &&
    keys[0] === "type" &&
    Object.hasOwn(value, "type") &&
    value["type"] === "authored"
  ) {
    return { binding: { type: "authored" }, status: "authored" };
  }
  if (
    keys.length === 2 &&
    keys.includes("type") &&
    keys.includes("token") &&
    Object.hasOwn(value, "type") &&
    value["type"] === "source" &&
    Object.hasOwn(value, "token") &&
    typeof value["token"] === "string"
  ) {
    return { binding: { token: value["token"], type: "source" }, status: "source" };
  }
  return { status: "invalid" };
};

const TABLE_CELL_PARAGRAPH_SOURCE_MAX_DEPTH = 128;
const TABLE_CELL_PARAGRAPH_SOURCE_MAX_VALUES = 100_000;
const ARRAY_INDEX = /^(?:0|[1-9][0-9]*)$/;

const tableCellParagraphSourcePayloadMessages = {
  complexity_limit: "The hidden table-cell payload exceeds its complexity limit.",
  cyclic_or_aliased_graph: "The hidden table-cell payload contains a cyclic or aliased graph.",
  depth_limit: "The hidden table-cell payload exceeds its nesting limit.",
  expected_array: "The hidden table-cell payload field must be an array.",
  expected_block: "The hidden table-cell payload contains a malformed block.",
  expected_cell: "The hidden table-cell payload contains a malformed cell.",
  expected_inline_content: "The hidden table-cell payload contains malformed inline content.",
  expected_paragraph: "The hidden table-cell payload contains a malformed paragraph.",
  expected_record: "The hidden table-cell payload contains a malformed object.",
  expected_row: "The hidden table-cell payload contains a malformed row.",
  invalid_binding: "A hidden table-cell paragraph has an invalid source binding.",
  invalid_block_type: "The hidden table-cell payload contains an unsupported block type.",
  invalid_cell_type: "The hidden table-cell payload contains an invalid cell type.",
  invalid_graph_value: "The hidden table-cell payload contains an invalid graph value.",
  invalid_inline_content_type:
    "The hidden table-cell payload contains an unsupported inline content type.",
  invalid_paragraph_type: "The hidden table-cell payload contains an invalid paragraph type.",
  invalid_row_type: "The hidden table-cell payload contains an invalid row type.",
  malformed_source_token: "A hidden table-cell paragraph has a malformed source token.",
} as const satisfies Record<TableCellParagraphSourcePayloadErrorClassification, string>;
const tableCellParagraphSourceValidationErrors =
  new WeakSet<ParagraphPropertySourceValidationError>();

const invalidTableCellParagraphSourcePayload = (
  classification: TableCellParagraphSourcePayloadErrorClassification,
  path: string,
): never => {
  const error = new ParagraphPropertySourceValidationError({
    code: "invalid_token",
    classification,
    message: tableCellParagraphSourcePayloadMessages[classification],
    path,
  });
  tableCellParagraphSourceValidationErrors.add(error);
  throw error;
};

type TableCellParagraphSourceDecodeContext = {
  paragraphs: Array<{
    binding: TableCellParagraphPropertySourceBinding;
    paragraph: Paragraph;
    path: string;
  }>;
};

type TableCellParagraphSourceGraphContext = {
  objects: object[];
  seen: WeakSet<object>;
  values: number;
};

type TableCellBlockTraversal = "blockSdt" | "paragraph" | "table";

const tableCellBlockTraversalByType = {
  blockSdt: "blockSdt",
  paragraph: "paragraph",
  table: "table",
} as const satisfies Record<BlockContent["type"], TableCellBlockTraversal>;

type TableCellParagraphContentTraversal = "children" | "complexField" | "content" | "leaf" | "run";

const tableCellParagraphContentTraversalByType = {
  bookmarkEnd: "leaf",
  bookmarkStart: "leaf",
  commentRangeEnd: "leaf",
  commentRangeStart: "leaf",
  commentReference: "leaf",
  complexField: "complexField",
  deletion: "content",
  hyperlink: "children",
  inlineSdt: "content",
  insertion: "content",
  mathEquation: "leaf",
  moveFrom: "content",
  moveFromRangeEnd: "leaf",
  moveFromRangeStart: "leaf",
  moveTo: "content",
  moveToRangeEnd: "leaf",
  moveToRangeStart: "leaf",
  run: "run",
  simpleField: "content",
} as const satisfies Record<ParagraphContent["type"], TableCellParagraphContentTraversal>;

type TableCellRunContentTraversal = "leaf" | "shape";

const tableCellRunContentTraversalByType = {
  break: "leaf",
  drawing: "leaf",
  endnoteRef: "leaf",
  fieldChar: "leaf",
  footnoteRef: "leaf",
  instrText: "leaf",
  noBreakHyphen: "leaf",
  renderedPageBreak: "leaf",
  shape: "shape",
  softHyphen: "leaf",
  symbol: "leaf",
  tab: "leaf",
  text: "leaf",
} as const satisfies Record<RunContent["type"], TableCellRunContentTraversal>;

const isTableCellBlockContentType = (
  value: unknown,
): value is keyof typeof tableCellBlockTraversalByType =>
  typeof value === "string" && Object.hasOwn(tableCellBlockTraversalByType, value);

const isTableCellParagraphContentType = (
  value: unknown,
): value is keyof typeof tableCellParagraphContentTraversalByType =>
  typeof value === "string" && Object.hasOwn(tableCellParagraphContentTraversalByType, value);

const isTableCellRunContentType = (
  value: unknown,
): value is keyof typeof tableCellRunContentTraversalByType =>
  typeof value === "string" && Object.hasOwn(tableCellRunContentTraversalByType, value);

const inspectTableCellParagraphSourceGraph = (
  value: unknown,
  path: string,
  depth: number,
  context: TableCellParagraphSourceGraphContext,
): void => {
  if (depth > TABLE_CELL_PARAGRAPH_SOURCE_MAX_DEPTH) {
    invalidTableCellParagraphSourcePayload("depth_limit", path);
  }
  context.values += 1;
  if (context.values > TABLE_CELL_PARAGRAPH_SOURCE_MAX_VALUES) {
    invalidTableCellParagraphSourcePayload("complexity_limit", path);
  }
  if (
    value === null ||
    value === undefined ||
    typeof value === "boolean" ||
    typeof value === "number" ||
    typeof value === "string"
  ) {
    return;
  }
  if (typeof value !== "object") {
    return invalidTableCellParagraphSourcePayload("invalid_graph_value", path);
  }
  if (context.seen.has(value)) {
    return invalidTableCellParagraphSourcePayload("cyclic_or_aliased_graph", path);
  }
  context.seen.add(value);
  context.objects.push(value);

  const prototype = Object.getPrototypeOf(value);
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) {
    return invalidTableCellParagraphSourcePayload("invalid_graph_value", path);
  }

  for (const key of Reflect.ownKeys(value)) {
    if (Array.isArray(value) && key === "length") {
      if (value.length > TABLE_CELL_PARAGRAPH_SOURCE_MAX_VALUES) {
        invalidTableCellParagraphSourcePayload("complexity_limit", path);
      }
      continue;
    }
    let childPath = `${path}.*`;
    if (Array.isArray(value)) {
      if (typeof key !== "string" || !ARRAY_INDEX.test(key)) {
        return invalidTableCellParagraphSourcePayload("invalid_graph_value", `${path}[*]`);
      }
      childPath = `${path}[${key}]`;
    }
    if (typeof key !== "string") {
      return invalidTableCellParagraphSourcePayload("invalid_graph_value", childPath);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (descriptor === undefined || !descriptor.enumerable || !("value" in descriptor)) {
      return invalidTableCellParagraphSourcePayload("invalid_graph_value", childPath);
    }
    inspectTableCellParagraphSourceGraph(descriptor.value, childPath, depth + 1, context);
  }
};

const tableCellParagraphSourceArray = (value: unknown, path: string): unknown[] => {
  if (!Array.isArray(value)) {
    return invalidTableCellParagraphSourcePayload("expected_array", path);
  }
  return value;
};

const tableCellParagraphSourceRecord = (
  value: unknown,
  path: string,
  classification: Extract<
    TableCellParagraphSourcePayloadErrorClassification,
    | "expected_block"
    | "expected_cell"
    | "expected_inline_content"
    | "expected_paragraph"
    | "expected_record"
    | "expected_row"
  >,
): Record<string, unknown> => {
  if (!isPropertySourceBindingRecord(value)) {
    return invalidTableCellParagraphSourcePayload(classification, path);
  }
  return value;
};

const isDecodedTableCellParagraph = (value: unknown): value is Paragraph =>
  isPropertySourceBindingRecord(value) &&
  value["type"] === "paragraph" &&
  Array.isArray(value["content"]);

const isDecodedTableCell = (value: unknown): value is TableCell =>
  isPropertySourceBindingRecord(value) &&
  value["type"] === "tableCell" &&
  Array.isArray(value["content"]);

const visitDecodedTableCellRunContent = (
  value: unknown,
  path: string,
  context: TableCellParagraphSourceDecodeContext,
): void => {
  const content = tableCellParagraphSourceRecord(value, path, "expected_inline_content");
  const contentType = content["type"];
  if (!isTableCellRunContentType(contentType)) {
    return invalidTableCellParagraphSourcePayload("invalid_inline_content_type", path);
  }
  const traversal = tableCellRunContentTraversalByType[contentType];
  switch (traversal) {
    case "shape": {
      const shapePath = `${path}.shape`;
      const shape = tableCellParagraphSourceRecord(content["shape"], shapePath, "expected_record");
      const rawTextBody = shape["textBody"];
      if (rawTextBody === undefined || rawTextBody === null) {
        return;
      }
      const textBodyPath = `${shapePath}.textBody`;
      const textBody = tableCellParagraphSourceRecord(rawTextBody, textBodyPath, "expected_record");
      const blocks = tableCellParagraphSourceArray(textBody["content"], `${textBodyPath}.content`);
      for (const [index, block] of blocks.entries()) {
        visitDecodedTableCellBlock(block, `${textBodyPath}.content[${index}]`, context);
      }
      return;
    }
    case "leaf":
      return;
    default: {
      const exhaustiveTraversal: never = traversal;
      return exhaustiveTraversal;
    }
  }
};

const visitDecodedTableCellRun = (
  run: Record<string, unknown>,
  path: string,
  context: TableCellParagraphSourceDecodeContext,
): void => {
  const content = tableCellParagraphSourceArray(run["content"], `${path}.content`);
  for (const [index, item] of content.entries()) {
    visitDecodedTableCellRunContent(item, `${path}.content[${index}]`, context);
  }
};

const visitDecodedTableCellInlineContent = (
  value: unknown,
  path: string,
  context: TableCellParagraphSourceDecodeContext,
): void => {
  const content = tableCellParagraphSourceRecord(value, path, "expected_inline_content");
  const contentType = content["type"];
  if (!isTableCellParagraphContentType(contentType)) {
    return invalidTableCellParagraphSourcePayload("invalid_inline_content_type", path);
  }
  const traversal = tableCellParagraphContentTraversalByType[contentType];
  switch (traversal) {
    case "run":
      visitDecodedTableCellRun(content, path, context);
      return;
    case "children": {
      const children = tableCellParagraphSourceArray(content["children"], `${path}.children`);
      for (const [index, child] of children.entries()) {
        visitDecodedTableCellInlineContent(child, `${path}.children[${index}]`, context);
      }
      return;
    }
    case "content": {
      const children = tableCellParagraphSourceArray(content["content"], `${path}.content`);
      for (const [index, child] of children.entries()) {
        visitDecodedTableCellInlineContent(child, `${path}.content[${index}]`, context);
      }
      return;
    }
    case "complexField": {
      for (const field of ["fieldCode", "fieldResult"] as const) {
        const runs = tableCellParagraphSourceArray(content[field], `${path}.${field}`);
        for (const [index, runValue] of runs.entries()) {
          const runPath = `${path}.${field}[${index}]`;
          const run = tableCellParagraphSourceRecord(runValue, runPath, "expected_inline_content");
          if (run["type"] !== "run") {
            return invalidTableCellParagraphSourcePayload("invalid_inline_content_type", runPath);
          }
          visitDecodedTableCellRun(run, runPath, context);
        }
      }
      return;
    }
    case "leaf":
      return;
    default: {
      const exhaustiveTraversal: never = traversal;
      return exhaustiveTraversal;
    }
  }
};

const visitDecodedTableCellParagraph = (
  paragraph: Record<string, unknown>,
  path: string,
  context: TableCellParagraphSourceDecodeContext,
): void => {
  if (paragraph["type"] !== "paragraph") {
    return invalidTableCellParagraphSourcePayload("invalid_paragraph_type", path);
  }
  if (!isDecodedTableCellParagraph(paragraph)) {
    return invalidTableCellParagraphSourcePayload("expected_paragraph", path);
  }
  const inspection = inspectTableCellParagraphPropertySourceBinding(paragraph);
  switch (inspection.status) {
    case "authored":
    case "source":
      if (
        inspection.status === "source" &&
        !isParagraphPropertySourceToken(inspection.binding.token)
      ) {
        invalidTableCellParagraphSourcePayload(
          "malformed_source_token",
          `${path}.${TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR}.token`,
        );
      }
      context.paragraphs.push({
        binding: inspection.binding,
        paragraph,
        path,
      });
      break;
    case "absent":
    case "invalid":
      return invalidTableCellParagraphSourcePayload(
        "invalid_binding",
        `${path}.${TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR}`,
      );
    default: {
      const exhaustiveInspection: never = inspection;
      return exhaustiveInspection;
    }
  }
  const content = tableCellParagraphSourceArray(paragraph.content, `${path}.content`);
  for (const [index, item] of content.entries()) {
    visitDecodedTableCellInlineContent(item, `${path}.content[${index}]`, context);
  }
};

const visitDecodedTableCell = (
  value: unknown,
  path: string,
  context: TableCellParagraphSourceDecodeContext,
): TableCell => {
  const cell = tableCellParagraphSourceRecord(value, path, "expected_cell");
  if (cell["type"] !== "tableCell") {
    return invalidTableCellParagraphSourcePayload("invalid_cell_type", path);
  }
  const content = tableCellParagraphSourceArray(cell["content"], `${path}.content`);
  for (const [index, block] of content.entries()) {
    visitDecodedTableCellBlock(block, `${path}.content[${index}]`, context);
  }
  if (!isDecodedTableCell(cell)) {
    return invalidTableCellParagraphSourcePayload("expected_cell", path);
  }
  return cell;
};

const visitDecodedTableCellRow = (
  value: unknown,
  path: string,
  context: TableCellParagraphSourceDecodeContext,
): void => {
  const row = tableCellParagraphSourceRecord(value, path, "expected_row");
  if (row["type"] !== "tableRow") {
    return invalidTableCellParagraphSourcePayload("invalid_row_type", path);
  }
  const cells = tableCellParagraphSourceArray(row["cells"], `${path}.cells`);
  for (const [index, cell] of cells.entries()) {
    visitDecodedTableCell(cell, `${path}.cells[${index}]`, context);
  }
};

function visitDecodedTableCellBlock(
  value: unknown,
  path: string,
  context: TableCellParagraphSourceDecodeContext,
): void {
  const block = tableCellParagraphSourceRecord(value, path, "expected_block");
  const blockType = block["type"];
  if (!isTableCellBlockContentType(blockType)) {
    return invalidTableCellParagraphSourcePayload("invalid_block_type", path);
  }
  const traversal = tableCellBlockTraversalByType[blockType];
  switch (traversal) {
    case "paragraph":
      visitDecodedTableCellParagraph(block, path, context);
      return;
    case "table": {
      const rows = tableCellParagraphSourceArray(block["rows"], `${path}.rows`);
      for (const [index, row] of rows.entries()) {
        visitDecodedTableCellRow(row, `${path}.rows[${index}]`, context);
      }
      return;
    }
    case "blockSdt": {
      const content = tableCellParagraphSourceArray(block["content"], `${path}.content`);
      for (const [index, child] of content.entries()) {
        visitDecodedTableCellBlock(child, `${path}.content[${index}]`, context);
      }
      return;
    }
    default: {
      const exhaustiveTraversal: never = traversal;
      return exhaustiveTraversal;
    }
  }
}

const decodedTableCellParagraphSourcePayloadBrand = Symbol(
  "decodedTableCellParagraphSourcePayload",
);

export type DecodedTableCellParagraphSourcePayload = Readonly<{
  [decodedTableCellParagraphSourcePayloadBrand]: true;
  cells: readonly TableCell[];
}>;

type DecodedTableCellParagraphSourcePayloadState = Readonly<{
  paragraphs: readonly Readonly<{
    binding: TableCellParagraphPropertySourceBinding;
    paragraph: Paragraph;
    path: string;
  }>[];
}>;

const decodedTableCellParagraphSourcePayloadStates = new WeakMap<
  DecodedTableCellParagraphSourcePayload,
  DecodedTableCellParagraphSourcePayloadState
>();
const decodedTableCellParagraphSourcePayloads = new WeakMap<
  object,
  DecodedTableCellParagraphSourcePayload
>();

const decodedTableCellParagraphSourcePayloadState = (
  payload: DecodedTableCellParagraphSourcePayload,
): DecodedTableCellParagraphSourcePayloadState => {
  const state = decodedTableCellParagraphSourcePayloadStates.get(payload);
  if (!state) {
    panic("A collapsed-cell paragraph source payload bypassed its decoder.");
  }
  return state;
};

const decodeTableCellParagraphSourcePayloadUnchecked = (
  value: unknown,
): DecodedTableCellParagraphSourcePayload => {
  const graphContext: TableCellParagraphSourceGraphContext = {
    objects: [],
    seen: new WeakSet(),
    values: 0,
  };
  inspectTableCellParagraphSourceGraph(value, "continuationCells", 0, graphContext);

  const context: TableCellParagraphSourceDecodeContext = {
    paragraphs: [],
  };
  const rawCells = tableCellParagraphSourceArray(value, "continuationCells");
  const cells: TableCell[] = [];
  for (const [index, rawCell] of rawCells.entries()) {
    cells.push(visitDecodedTableCell(rawCell, `continuationCells[${index}]`, context));
  }

  let object = graphContext.objects.pop();
  while (object) {
    Object.freeze(object);
    object = graphContext.objects.pop();
  }
  const payload = Object.freeze({
    [decodedTableCellParagraphSourcePayloadBrand]: true,
    cells: Object.freeze(cells),
  } satisfies DecodedTableCellParagraphSourcePayload);
  decodedTableCellParagraphSourcePayloadStates.set(
    payload,
    Object.freeze({
      paragraphs: Object.freeze(context.paragraphs.map((entry) => Object.freeze(entry))),
    } satisfies DecodedTableCellParagraphSourcePayloadState),
  );
  decodedTableCellParagraphSourcePayloads.set(payload.cells, payload);
  return payload;
};

/** Decode and freeze the opaque collapsed-cell wire payload before typed use. */
export const decodeTableCellParagraphSourcePayload = (
  value: unknown,
): DecodedTableCellParagraphSourcePayload => {
  if (typeof value === "object" && value !== null) {
    const decoded = decodedTableCellParagraphSourcePayloads.get(value);
    if (decoded) {
      return decoded;
    }
  }
  try {
    const decoded = decodeTableCellParagraphSourcePayloadUnchecked(value);
    if (typeof value === "object" && value !== null) {
      decodedTableCellParagraphSourcePayloads.set(value, decoded);
    }
    return decoded;
  } catch (error) {
    if (
      error instanceof ParagraphPropertySourceValidationError &&
      tableCellParagraphSourceValidationErrors.has(error)
    ) {
      throw error;
    }
    return invalidTableCellParagraphSourcePayload("invalid_graph_value", "continuationCells");
  }
};

const setTableCellParagraphPropertySourceBinding = (
  paragraph: Paragraph,
  binding: TableCellParagraphPropertySourceBinding,
): void => {
  if (!Reflect.set(paragraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR, Object.freeze(binding))) {
    panic("Cannot attach paragraph-property transport identity to a table cell.");
  }
};

const tableCellParagraphPropertySourceBindingForTransport = (
  paragraph: Paragraph,
  inspection: TableCellParagraphPropertySourceBindingInspection,
): TableCellParagraphPropertySourceBinding => {
  const directToken = paragraphPropertySourceTokens.get(paragraph);
  if (directToken !== undefined) {
    if (!isParagraphPropertySourceToken(directToken)) {
      panic("Cannot transport a malformed private paragraph-property source token.");
    }
    switch (inspection.status) {
      case "absent":
        return { token: directToken, type: "source" };
      case "authored":
        panic("A hidden paragraph-property source binding conflicts with its private owner.");
      case "source":
        if (inspection.binding.token !== directToken) {
          panic("A hidden paragraph-property source binding conflicts with its private owner.");
        }
        return inspection.binding;
      case "invalid":
        panic("Cannot transport a malformed hidden paragraph-property source binding.");
      default: {
        const exhaustiveInspection: never = inspection;
        return exhaustiveInspection;
      }
    }
  }
  switch (inspection.status) {
    case "absent":
      // Transport is the construction boundary where an unbound paragraph is
      // explicitly classified as newly authored.
      return { type: "authored" };
    case "authored":
      return inspection.binding;
    case "source":
      if (!isParagraphPropertySourceToken(inspection.binding.token)) {
        panic("Cannot transport a malformed hidden paragraph-property source token.");
      }
      return inspection.binding;
    case "invalid":
      panic("Cannot transport a malformed hidden paragraph-property source binding.");
    default: {
      const exhaustiveInspection: never = inspection;
      return exhaustiveInspection;
    }
  }
};

/** Prepare opaque continuation cells for ProseMirror and Yjs transport. */
export const transportTableCellsWithParagraphPropertySources = (
  cells: TableCell[],
): TableCell[] => {
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
    const inspection = inspectTableCellParagraphPropertySourceBinding(source);
    setTableCellParagraphPropertySourceBinding(
      target,
      tableCellParagraphPropertySourceBindingForTransport(source, inspection),
    );
  }
  return cloned;
};

/** Restore decoded identities privately and remove them from the document model. */
export const restoreTableCellsWithParagraphPropertySources = (
  payload: DecodedTableCellParagraphSourcePayload,
): TableCell[] => {
  const sourceEntries = decodedTableCellParagraphSourcePayloadState(payload).paragraphs;
  const cloned = structuredClone([...payload.cells]);
  const targets = paragraphsInTableCells(cloned);
  if (sourceEntries.length !== targets.length) {
    panic("The cloned table cells changed paragraph graph ownership.");
  }
  for (const [index, sourceEntry] of sourceEntries.entries()) {
    const target = targets.at(index);
    if (!target) {
      panic("The cloned table cells lost a paragraph owner.");
    }
    const { paragraph: source } = sourceEntry;
    copyParagraphPropertyCapture(target, source);
    if (!Reflect.deleteProperty(target, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR)) {
      panic("Cannot detach paragraph-property transport identity from a table cell.");
    }
    switch (sourceEntry.binding.type) {
      case "authored":
        break;
      case "source":
        paragraphPropertySourceTokens.set(target, sourceEntry.binding.token);
        break;
      default: {
        const exhaustiveBinding: never = sourceEntry.binding;
        return exhaustiveBinding;
      }
    }
  }
  return cloned;
};

/** Visit every transported hidden paragraph in canonical cell-story order. */
export const visitTableCellParagraphPropertySourceBindings = (
  payload: DecodedTableCellParagraphSourcePayload,
  visit: (binding: TableCellParagraphPropertySourceBinding, path: string) => void,
): void => {
  for (const { binding, path } of decodedTableCellParagraphSourcePayloadState(payload).paragraphs) {
    visit(binding, path);
  }
};

/**
 * Clone package-crossing vertical-merge payloads with their captured `w:pPr`,
 * but without the durable paragraph tokens owned by the source package.
 */
export const cloneTableCellsWithParagraphPropertyCaptures = (
  payload: DecodedTableCellParagraphSourcePayload,
): TableCell[] => {
  const sourceEntries = decodedTableCellParagraphSourcePayloadState(payload).paragraphs;
  const cloned = structuredClone([...payload.cells]);
  const targets = paragraphsInTableCells(cloned);
  if (sourceEntries.length !== targets.length) {
    panic("The cloned table cells changed paragraph graph ownership.");
  }
  for (const [index, sourceEntry] of sourceEntries.entries()) {
    const target = targets.at(index);
    if (!target) {
      panic("The cloned table cells lost a paragraph owner.");
    }
    copyParagraphPropertyCapture(target, sourceEntry.paragraph);
    setTableCellParagraphPropertySourceBinding(target, { type: "authored" });
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
  linkProseParagraphProjectionOwner(paragraph, sourceParagraph);
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

/** Bind every projected PM paragraph to its exact Document paragraph owner. */
export const linkProseParagraphProjectionOwner = (
  proseParagraph: PMNode,
  sourceParagraph: Paragraph,
): void => {
  proseParagraphProjectionOwners.set(proseParagraph, sourceParagraph);
};
/** Carry a parser-linked paragraph owner across an immutable PM node rebuild. */
const copyProseParagraphSources = (target: PMNode, source: PMNode): void => {
  if (target.type.name !== "paragraph" || source.type.name !== "paragraph") {
    return;
  }
  const sourceOwner = proseParagraphSourceOwners.get(source);
  if (sourceOwner) {
    proseParagraphSourceOwners.set(target, sourceOwner);
  }
  const projectionOwner = proseParagraphProjectionOwners.get(source);
  if (projectionOwner) {
    proseParagraphProjectionOwners.set(target, projectionOwner);
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
  copyProseParagraphSources(target, source);
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
    copyProseParagraphSources(target, source);
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

/**
 * Record the live-only identity allocated while this PM paragraph still has
 * an exact Document owner. Reprojection restores it during its existing
 * conversion walk, without a positional join or another pass.
 */
export const transferSynthesizedParagraphIdentity = (
  target: PMNode,
  source: PMNode,
  paraId: string,
): void => {
  const sourceOwner = proseParagraphProjectionOwners.get(source);
  if (!sourceOwner) {
    return;
  }
  proseParagraphProjectionOwners.set(target, sourceOwner);
  synthesizedIdentityByParagraph.set(sourceOwner, Object.freeze({ type: "synthesized", paraId }));
};

/** Retain positional identity while PM content is materialized as a Document paragraph. */
export const captureSynthesizedParagraphIdentity = (target: Paragraph, source: PMNode): void => {
  const idStability = source.attrs[PARAGRAPH_ID_STABILITY_ATTR];
  if (idStability === undefined || idStability === null) {
    return;
  }
  if (idStability !== POSITIONAL_PARAGRAPH_ID_STABILITY) {
    panic("A paragraph carried unsupported live identity provenance.", { idStability });
  }
  const paraId = source.attrs["paraId"];
  if (typeof paraId !== "string" || paraId.length === 0) {
    panic("A positional paragraph identity requires a generated paraId.");
  }
  synthesizedIdentityByParagraph.set(target, Object.freeze({ type: "synthesized", paraId }));
};

/** Resolve the same paragraph id the private Document-to-PM projection emits. */
export const paragraphProjectionParaId = (paragraph: Paragraph): string | undefined => {
  const identity = synthesizedIdentityByParagraph.get(paragraph);
  if (!identity) {
    return paragraph.paraId;
  }
  if (paragraph.paraId !== undefined && paragraph.paraId !== identity.paraId) {
    panic("A paragraph's Document identity conflicts with its live projection provenance.", {
      documentParaId: paragraph.paraId,
      projectedParaId: identity.paraId,
    });
  }
  return identity.paraId;
};

/** Restore private identity provenance into an existing Document-to-PM conversion. */
export const applySynthesizedParagraphIdentity = (source: Paragraph, targetAttrs: object): void => {
  const identity = synthesizedIdentityByParagraph.get(source);
  if (!identity) {
    return;
  }
  Reflect.set(targetAttrs, "paraId", paragraphProjectionParaId(source));
  Reflect.set(targetAttrs, PARAGRAPH_ID_STABILITY_ATTR, POSITIONAL_PARAGRAPH_ID_STABILITY);
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
