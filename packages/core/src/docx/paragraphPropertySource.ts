import { panic } from "better-result";
import type { Fragment, Mark, Node as PMNode, NodeType } from "prosemirror-model";
import { PluginKey, type Transaction } from "prosemirror-state";

import type { BlockContent, Document, Paragraph, TableCell } from "../types/document";
import { visitDocxParagraphs } from "./paragraphTraversal";
import {
  canonicalParagraphPropertySourceFingerprintJson,
  paragraphPropertySourceFingerprintFromParts,
  type AuthoredParagraphProperties,
  type ParagraphPropertySourceFingerprint,
} from "@stll/docx-core/model";
import {
  ParagraphPropertySourceContract,
  ParagraphPropertySourceToken,
  ParagraphPropertyTransientTemplateResolution,
  ParagraphPropertyTransientTemplateStore,
  ParagraphPropertyTransientTemplateHandle,
  ParagraphPropertySourceValidationError,
  paragraphPropertySourceStoryKey,
} from "./paragraphPropertySourceIdentity";
import type {
  ParagraphPropertySourceAttribute,
  ParagraphPropertySourceStory,
  ParagraphPropertySourceValidationCode,
} from "./paragraphPropertySourceIdentity";

export {
  PARAGRAPH_PROPERTY_SOURCE_VALIDATION_CODES,
  ParagraphPropertySourceContract,
  ParagraphPropertySourceToken,
  ParagraphPropertyTransientTemplateHandle,
  ParagraphPropertySourceValidationError,
} from "./paragraphPropertySourceIdentity";
export type {
  ParagraphPropertySourceAttribute,
  ParagraphPropertySourceStory,
  ParagraphPropertySourceValidationCode,
} from "./paragraphPropertySourceIdentity";

export type ParagraphPropertySource = {
  fingerprint: ParagraphPropertySourceFingerprint;
  fingerprintJson: string;
  type: "present";
  xml: string;
};

type ParagraphPropertyCapture =
  | { fingerprint: ParagraphPropertySourceFingerprint; fingerprintJson: string; type: "absent" }
  | ParagraphPropertySource;

type ParagraphPropertySourceBinding =
  | { capture: ParagraphPropertyCapture; type: "captured-unbound" }
  | { type: "editor-created" }
  | {
      capture: ParagraphPropertyCapture;
      token: ParagraphPropertySourceToken;
      type: "imported";
    }
  | { capture: ParagraphPropertyCapture; type: "resolved-template" };

// Enumerable symbols follow ordinary immutable `{ ...document }` derivations,
// while JSON and other string-key serialization cannot expose the contract.
// `structuredClone` deliberately drops symbols, so the one sanctioned deep
// clone path transfers this value explicitly below.
const documentParagraphPropertySourceContract = Symbol("paragraphPropertySourceContract");
// Enumerable so immutable `{ ...paragraph }` derivations retain the capture.
// Symbols are excluded from JSON and OOXML serialization. Structured-clone
// call sites must transfer this binding through the sanctioned helper below.
const paragraphPropertySource = Symbol("paragraphPropertySource");

const freezeDeep = (value: object): void => {
  for (const key of Reflect.ownKeys(value)) {
    const child = Reflect.get(value, key);
    if (typeof child === "object" && child !== null && !Object.isFrozen(child)) {
      freezeDeep(child);
    }
  }
  Object.freeze(value);
};

const immutableFingerprint = (
  fingerprint: ParagraphPropertySourceFingerprint,
): ParagraphPropertySourceFingerprint => {
  const cloned = structuredClone(fingerprint);
  freezeDeep(cloned);
  return cloned;
};

const isDeepFrozen = (value: object): boolean => {
  if (!Object.isFrozen(value)) {
    return false;
  }
  for (const key of Reflect.ownKeys(value)) {
    const child = Reflect.get(value, key);
    if (typeof child === "object" && child !== null && !isDeepFrozen(child)) {
      return false;
    }
  }
  return true;
};

const isParagraphPropertyCapture = (value: unknown): value is ParagraphPropertyCapture => {
  if (typeof value !== "object" || value === null || !("type" in value)) {
    return false;
  }
  if (value.type === "absent") {
    return (
      "fingerprint" in value &&
      typeof value.fingerprint === "object" &&
      value.fingerprint !== null &&
      "fingerprintJson" in value &&
      typeof value.fingerprintJson === "string" &&
      isDeepFrozen(value.fingerprint) &&
      canonicalParagraphPropertySourceFingerprintJson(value.fingerprint) ===
        value.fingerprintJson &&
      canonicalParagraphPropertySourceFingerprintJson(
        paragraphPropertySourceFingerprintFromParts({}, {}),
      ) === value.fingerprintJson &&
      Object.isFrozen(value)
    );
  }
  if (
    value.type !== "present" ||
    !("fingerprint" in value) ||
    typeof value.fingerprint !== "object" ||
    value.fingerprint === null ||
    !("fingerprintJson" in value) ||
    typeof value.fingerprintJson !== "string" ||
    !("xml" in value) ||
    typeof value.xml !== "string"
  ) {
    return false;
  }
  return (
    Object.isFrozen(value) &&
    isDeepFrozen(value.fingerprint) &&
    canonicalParagraphPropertySourceFingerprintJson(value.fingerprint) === value.fingerprintJson
  );
};

const isParagraphPropertySourceBinding = (
  value: unknown,
): value is ParagraphPropertySourceBinding => {
  if (
    typeof value !== "object" ||
    value === null ||
    !("type" in value) ||
    !Object.isFrozen(value)
  ) {
    return false;
  }
  if (value.type === "editor-created") {
    return Object.keys(value).length === 1;
  }
  if (!("capture" in value) || !isParagraphPropertyCapture(value.capture)) {
    return false;
  }
  switch (value.type) {
    case "captured-unbound":
    case "resolved-template":
      return true;
    case "imported":
      return "token" in value && value.token instanceof ParagraphPropertySourceToken;
    default:
      return false;
  }
};

export const PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR = "_docxParagraphSourceToken";
export const PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR = "_docxParagraphSourceContract";

const setDocumentParagraphPropertySourceContract = (
  document: Document,
  contract: ParagraphPropertySourceContract,
): void => {
  const existing = Object.hasOwn(document, documentParagraphPropertySourceContract)
    ? Reflect.get(document, documentParagraphPropertySourceContract)
    : undefined;
  if (
    !Reflect.defineProperty(document, documentParagraphPropertySourceContract, {
      configurable: false,
      enumerable: true,
      value: contract,
      writable: false,
    })
  ) {
    panic("Cannot attach the paragraph-property source contract to the document");
  }
};

const paragraphPropertySourceBinding = (
  paragraph: Paragraph,
): ParagraphPropertySourceBinding | undefined => {
  if (!Object.hasOwn(paragraph, paragraphPropertySource)) {
    return undefined;
  }
  const binding = Reflect.get(paragraph, paragraphPropertySource);
  if (!isParagraphPropertySourceBinding(binding)) {
    panic("A paragraph carries an invalid paragraph-property source binding");
  }
  const descriptor = Object.getOwnPropertyDescriptor(paragraph, paragraphPropertySource);
  if (
    !descriptor ||
    descriptor.enumerable !== true ||
    descriptor.writable === true ||
    descriptor.configurable !== (binding.type === "captured-unbound")
  ) {
    panic("A paragraph carries a weakened paragraph-property source binding");
  }
  return binding;
};

const setParagraphPropertySourceBinding = (
  paragraph: Paragraph,
  binding: ParagraphPropertySourceBinding,
): void => {
  if (
    !Reflect.defineProperty(paragraph, paragraphPropertySource, {
      configurable: binding.type === "captured-unbound",
      enumerable: true,
      value: Object.freeze(binding),
      writable: false,
    })
  ) {
    panic("Cannot attach paragraph-property source provenance to a paragraph");
  }
};

/** Visit one story in canonical OOXML paragraph order. */
export const visitDocumentStoryParagraphs = (
  content: Document["package"]["document"]["content"],
  visit: (paragraph: Paragraph) => void,
): void => {
  visitDocxParagraphs({ documentBody: { content } }, visit);
};

export const assignParagraphPropertySource = (
  paragraph: Paragraph,
  source: { fingerprint: ParagraphPropertySourceFingerprint; xml: string },
): void => {
  const existing = paragraphPropertySourceBinding(paragraph);
  const fingerprint = immutableFingerprint(source.fingerprint);
  const capture = Object.freeze({
    fingerprint,
    fingerprintJson: canonicalParagraphPropertySourceFingerprintJson(fingerprint),
    type: "present",
    xml: source.xml,
  }) satisfies ParagraphPropertyCapture;
  switch (existing?.type) {
    case undefined:
      setParagraphPropertySourceBinding(paragraph, {
        capture,
        type: "captured-unbound",
      });
      return;
    case "captured-unbound":
    case "editor-created":
    case "imported":
    case "resolved-template":
      panic("A paragraph-property capture can only be assigned once");
      return;
    default: {
      const exhaustive: never = existing;
      return exhaustive;
    }
  }
};

/** Record that a parsed paragraph had no authored `w:pPr` element. */
export const assignAbsentParagraphPropertySource = (paragraph: Paragraph): void => {
  if (paragraphPropertySourceBinding(paragraph)) {
    panic("A paragraph-property capture can only be assigned once");
  }
  const fingerprint = immutableFingerprint(
    paragraphPropertySourceFingerprintFromParts({}, {}),
  );
  setParagraphPropertySourceBinding(paragraph, {
    capture: Object.freeze({
      fingerprint,
      fingerprintJson: canonicalParagraphPropertySourceFingerprintJson(fingerprint),
      type: "absent",
    }),
    type: "captured-unbound",
  });
};

export const getParagraphPropertySource = (
  paragraph: Paragraph,
): ParagraphPropertySource | undefined => {
  const binding = paragraphPropertySourceBinding(paragraph);
  if (!binding || binding.type === "editor-created" || binding.capture.type === "absent") {
    return undefined;
  }
  return binding.capture;
};

/** Authored `w:pPr` before style and numbering defaults are materialized. */
export const getParagraphAuthoredPPr = (
  paragraph: Paragraph,
): AuthoredParagraphProperties | undefined => {
  const binding = paragraphPropertySourceBinding(paragraph);
  if (!binding || binding.type === "editor-created") {
    return undefined;
  }
  return structuredClone(binding.capture.fingerprint.pPrBase);
};

/** Exact modeled properties captured from the paragraph's authored `w:pPr`. */
export const getParagraphPropertySourceFingerprint = (
  paragraph: Paragraph,
): ParagraphPropertySourceFingerprint | undefined => {
  const binding = paragraphPropertySourceBinding(paragraph);
  return binding && binding.type !== "editor-created" ? binding.capture.fingerprint : undefined;
};

/** Mark a model paragraph as deliberately created outside the parsed source census. */
export const assignEditorCreatedParagraphPropertySource = (paragraph: Paragraph): void => {
  if (paragraphPropertySourceBinding(paragraph)) {
    panic("Paragraph-property provenance can only be assigned once");
  }
  setParagraphPropertySourceBinding(paragraph, { type: "editor-created" });
};

const tokenFromBinding = (
  binding: ParagraphPropertySourceBinding | undefined,
): ParagraphPropertySourceToken | null => {
  switch (binding?.type) {
    case "imported":
      return binding.token;
    case "captured-unbound":
    case "editor-created":
    case "resolved-template":
    case undefined:
      return null;
    default: {
      const exhaustive: never = binding;
      return exhaustive;
    }
  }
};

/** Copy the captured `w:pPr` without claiming the source paragraph's durable identity. */
export const copyParagraphPropertyCapture = (target: Paragraph, source: Paragraph): void => {
  const sourceBinding = paragraphPropertySourceBinding(source);
  if (!sourceBinding || sourceBinding.type === "editor-created") {
    return;
  }
  setParagraphPropertySourceBinding(target, {
    capture: sourceBinding.capture,
    type: "resolved-template",
  });
};

export const PARAGRAPH_PROPERTY_TEMPLATE_STORE_MAX_CAPTURES = 100_000;

const paragraphPropertyTemplateCaptureIssuer = Symbol("paragraphPropertyTemplateCaptureIssuer");
type ParagraphPropertyTemplateCaptureIssuer = typeof paragraphPropertyTemplateCaptureIssuer;

/** Immutable authority to transfer one parsed capture without exposing its paragraph owner. */
export class ParagraphPropertyTemplateCapture {
  readonly #capture: ParagraphPropertyCapture;

  constructor(issuer: ParagraphPropertyTemplateCaptureIssuer, capture: ParagraphPropertyCapture) {
    if (issuer !== paragraphPropertyTemplateCaptureIssuer) {
      panic("Only a bound paragraph story may issue template capture capabilities");
    }
    this.#capture = capture;
    Object.freeze(this);
  }

  capture(issuer: ParagraphPropertyTemplateCaptureIssuer): ParagraphPropertyCapture {
    if (issuer !== paragraphPropertyTemplateCaptureIssuer) {
      return panic("Only the paragraph-property source kernel may resolve a template capture");
    }
    return this.#capture;
  }
}

export class ParagraphPropertyTemplateResolutionRegistry {
  readonly #resolution: ParagraphPropertyTransientTemplateResolution<ParagraphPropertyCapture>;

  constructor(
    issuer: ParagraphPropertyTemplateCaptureIssuer,
    resolution: ParagraphPropertyTransientTemplateResolution<ParagraphPropertyCapture>,
  ) {
    if (issuer !== paragraphPropertyTemplateCaptureIssuer) {
      panic("Only a paragraph-property template store may create a resolution registry");
    }
    this.#resolution = resolution;
    Object.freeze(this);
  }

  assertFullyConsumed(): void {
    this.#resolution.assertFullyConsumed();
  }

  consume(handle: ParagraphPropertyTransientTemplateHandle, target: Paragraph): void {
    if (paragraphPropertySourceBinding(target)) {
      panic("A paragraph-property template target already has source provenance");
    }
    const capture = this.#resolution.consume(handle);
    setParagraphPropertySourceBinding(target, { capture, type: "resolved-template" });
  }
}

/** Bounded session owner for opaque captures retained by edit and undo history. */
export class ParagraphPropertyTemplateStore {
  readonly #store = new ParagraphPropertyTransientTemplateStore<ParagraphPropertyCapture>(
    PARAGRAPH_PROPERTY_TEMPLATE_STORE_MAX_CAPTURES,
  );

  beginResolution(
    handles: readonly ParagraphPropertyTransientTemplateHandle[],
  ): ParagraphPropertyTemplateResolutionRegistry {
    return new ParagraphPropertyTemplateResolutionRegistry(
      paragraphPropertyTemplateCaptureIssuer,
      this.#store.beginResolution(handles),
    );
  }

  registerAll(
    capabilities: readonly ParagraphPropertyTemplateCapture[],
  ): readonly ParagraphPropertyTransientTemplateHandle[] {
    return this.#store.registerAll(
      capabilities.map((capability) =>
        capability.capture(paragraphPropertyTemplateCaptureIssuer),
      ),
    );
  }
}

export const createParagraphPropertyTemplateStore = (): ParagraphPropertyTemplateStore =>
  new ParagraphPropertyTemplateStore();

export const copyParagraphPropertySource = (target: Paragraph, source: Paragraph): void => {
  const binding = paragraphPropertySourceBinding(source);
  if (!binding) {
    return;
  }
  switch (binding.type) {
    case "imported":
      setParagraphPropertySourceBinding(target, binding);
      return;
    case "captured-unbound":
      setParagraphPropertySourceBinding(target, binding);
      return;
    case "editor-created":
      setParagraphPropertySourceBinding(target, binding);
      return;
    case "resolved-template":
      setParagraphPropertySourceBinding(target, binding);
      return;
    default: {
      const exhaustive: never = binding;
      return exhaustive;
    }
  }
};

/** Bind every parsed story paragraph to one exact source package. */
export const assignDocumentParagraphPropertySourceContract = (
  document: Document,
  sourceDigest: string,
): void => {
  if (Object.hasOwn(document, documentParagraphPropertySourceContract)) {
    panic("A document paragraph-property source contract can only be assigned once");
  }
  const contract = ParagraphPropertySourceContract.fromDigest(sourceDigest);
  const seenStories = new Set<string>();
  const seenParagraphs = new WeakSet<Paragraph>();
  const bodyParagraphs = new WeakSet<Paragraph>();
  const census: {
    capture: ParagraphPropertyCapture;
    paragraph: Paragraph;
    token: ParagraphPropertySourceToken;
  }[] = [];
  if (!Object.isExtensible(document)) {
    panic("A document must be extensible before binding paragraph-property provenance");
  }
  // The traversal is part of the v2 durable identity contract. Any ordering
  // change requires a token-version bump and collaboration reseed.
  for (const { content, story } of documentSourceStories(document)) {
    const storyKey = paragraphPropertySourceStoryKey(story);
    if (seenStories.has(storyKey)) {
      panic("A document contains duplicate paragraph-property story identity", { storyKey });
    }
    seenStories.add(storyKey);
    const paragraphs: Paragraph[] = [];
    visitDocumentStoryParagraphs(content, (paragraph) => paragraphs.push(paragraph));
    const tokenCensus = contract.bindStoryCensus(story, paragraphs.length);
    for (const [ordinal, paragraph] of paragraphs.entries()) {
      if (seenParagraphs.has(paragraph)) {
        panic("A paragraph cannot belong to more than one source story", { storyKey });
      }
      seenParagraphs.add(paragraph);
      if (story.type === "document") {
        bodyParagraphs.add(paragraph);
      }
      const binding = paragraphPropertySourceBinding(paragraph);
      if (binding?.type !== "captured-unbound") {
        panic("A document paragraph must have exactly one parsed property capture before binding");
      }
      const descriptor = Object.getOwnPropertyDescriptor(paragraph, paragraphPropertySource);
      if (!descriptor?.configurable) {
        panic("A parsed paragraph capture cannot transition to its durable source identity");
      }
      census.push({ capture: binding.capture, paragraph, token: tokenCensus.tokenAt(ordinal) });
    }
  }
  for (const section of document.package.document.sections ?? []) {
    visitDocumentStoryParagraphs(section.content, (paragraph) => {
      if (!bodyParagraphs.has(paragraph)) {
        panic("A derived document section must alias paragraphs from the body story");
      }
    });
  }

  setDocumentParagraphPropertySourceContract(document, contract);
  for (const { capture, paragraph, token } of census) {
    setParagraphPropertySourceBinding(paragraph, {
      capture,
      token,
      type: "imported",
    });
  }
};

export const getDocumentParagraphPropertySourceContract = (
  document: Document,
): ParagraphPropertySourceContract | undefined => {
  if (!Object.hasOwn(document, documentParagraphPropertySourceContract)) {
    return undefined;
  }
  const contract = Reflect.get(document, documentParagraphPropertySourceContract);
  if (!(contract instanceof ParagraphPropertySourceContract)) {
    panic("The document carries an invalid paragraph-property source contract");
  }
  const descriptor = Object.getOwnPropertyDescriptor(
    document,
    documentParagraphPropertySourceContract,
  );
  if (
    !descriptor ||
    descriptor.enumerable !== true ||
    descriptor.writable === true ||
    descriptor.configurable === true
  ) {
    panic("The document carries a weakened paragraph-property source contract");
  }
  return contract;
};

export const readProseDocumentParagraphPropertySourceContract = (
  document: PMNode,
): ParagraphPropertySourceAttribute<ParagraphPropertySourceContract> =>
  ParagraphPropertySourceContract.read(document.attrs[PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]);

export const readProseParagraphPropertySourceToken = (
  paragraph: PMNode,
): ParagraphPropertySourceAttribute<ParagraphPropertySourceToken> =>
  ParagraphPropertySourceToken.read(paragraph.attrs[PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]);

export const copyDocumentParagraphPropertySourceContract = (
  target: Document,
  source: Document,
): void => {
  const contract = getDocumentParagraphPropertySourceContract(source);
  if (contract) {
    setDocumentParagraphPropertySourceContract(target, contract);
  }
};

type DocumentDerivationOverrides = Partial<Document>;

/** Shallow document derivation that re-hardens its private source contract. */
export const deriveDocumentWithParagraphPropertySources = (
  document: Document,
  overrides: DocumentDerivationOverrides,
): Document => {
  const derived: Document = { ...document, ...overrides };
  copyDocumentParagraphPropertySourceContract(derived, document);
  return derived;
};

export const getParagraphPropertySourceToken = (
  paragraph: Paragraph,
): ParagraphPropertySourceToken | undefined => tokenFromBinding(paragraphPropertySourceBinding(paragraph)) ?? undefined;

const indexParagraphPropertySources = (
  content: BlockContent[],
  story: ParagraphPropertySourceStory,
  contract: ParagraphPropertySourceContract,
): ReadonlyMap<string, ParagraphPropertyTemplateCapture> => {
  const sources = new Map<string, ParagraphPropertyTemplateCapture>();
  visitDocumentStoryParagraphs(content, (paragraph) => {
    const binding = paragraphPropertySourceBinding(paragraph);
    if (binding?.type === "editor-created") {
      return;
    }
    const token = getParagraphPropertySourceToken(paragraph);
    if (!token || !token.belongsToStory(story) || !contract.owns(token)) {
      throw new ParagraphPropertySourceValidationError({
        code: "invalid_token",
        message: "The source story contains an invalid paragraph-property token.",
        token,
      });
    }
    if (sources.has(token.serialized)) {
      throw new ParagraphPropertySourceValidationError({
        code: "duplicate_token",
        message: "The source story contains a duplicate paragraph-property token.",
        token: token.serialized,
      });
    }
    const capture = binding?.type === "imported" ? binding.capture : undefined;
    if (!capture) {
      panic("A validated imported paragraph lost its source capture");
    }
    sources.set(
      token.serialized,
      new ParagraphPropertyTemplateCapture(paragraphPropertyTemplateCaptureIssuer, capture),
    );
  });
  return sources;
};

/** Bound source story whose content, identity, and package contract cannot diverge. */
export class ParagraphPropertyStorySource {
  readonly #contract: ParagraphPropertySourceContract;
  readonly #captures: ReadonlyMap<string, ParagraphPropertyTemplateCapture>;
  readonly #story: ParagraphPropertySourceStory;

  private constructor(
    contract: ParagraphPropertySourceContract,
    story: ParagraphPropertySourceStory,
    captures: ReadonlyMap<string, ParagraphPropertyTemplateCapture>,
  ) {
    this.#contract = contract;
    this.#story = story;
    this.#captures = captures;
    Object.freeze(this);
  }

  static fromDocument(
    document: Document,
    requestedStory: ParagraphPropertySourceStory,
  ): ParagraphPropertyStorySource {
    const contract = getDocumentParagraphPropertySourceContract(document);
    if (!contract) {
      throw new ParagraphPropertySourceValidationError({
        code: "contract_mismatch",
        message: "A paragraph-property story source requires a bound document contract.",
      });
    }
    const requestedKey = paragraphPropertySourceStoryKey(requestedStory);
    const matches = documentSourceStories(document).filter(
      ({ story }) => paragraphPropertySourceStoryKey(story) === requestedKey,
    );
    if (matches.length !== 1) {
      throw new ParagraphPropertySourceValidationError({
        code: "ambiguous_source",
        message: "The document does not contain exactly one requested paragraph-property story.",
      });
    }
    const match = matches.at(0);
    if (!match) {
      return panic("A unique paragraph-property story match disappeared");
    }
    return new ParagraphPropertyStorySource(
      contract,
      match.story,
      indexParagraphPropertySources(match.content, match.story, contract),
    );
  }

  owns(token: ParagraphPropertySourceToken): boolean {
    return (
      this.#contract.owns(token) &&
      token.belongsToStory(this.#story) &&
      this.#captures.has(token.serialized)
    );
  }

  readToken(raw: unknown): ParagraphPropertySourceToken {
    const token = this.#contract.readToken(raw);
    if (token.status !== "valid") {
      throw new ParagraphPropertySourceValidationError({
        code: "invalid_token",
        message: "A paragraph-property source token is malformed.",
        token: token.status === "invalid" ? token.raw : raw,
      });
    }
    if (!this.owns(token.value)) {
      throw new ParagraphPropertySourceValidationError({
        code: "unknown_token",
        message: "A paragraph-property source token is not owned by this story.",
        token: token.value.serialized,
      });
    }
    return token.value;
  }

  bindImportedParagraph(target: Paragraph, token: ParagraphPropertySourceToken): void {
    if (paragraphPropertySourceBinding(target)) {
      panic("An imported paragraph target already has source provenance");
    }
    const capability = this.owns(token) ? this.#captures.get(token.serialized) : undefined;
    if (!capability) {
      throw new ParagraphPropertySourceValidationError({
        code: "unknown_token",
        message: "A paragraph-property token is not owned by this source story.",
        token: token.serialized,
      });
    }
    setParagraphPropertySourceBinding(target, {
      capture: capability.capture(paragraphPropertyTemplateCaptureIssuer),
      token,
      type: "imported",
    });
  }

  templateCapture(token: ParagraphPropertySourceToken): ParagraphPropertyTemplateCapture {
    const capability = this.owns(token) ? this.#captures.get(token.serialized) : undefined;
    if (!capability) {
      throw new ParagraphPropertySourceValidationError({
        code: "unknown_token",
        message: "A paragraph-property token is not owned by this source story.",
        token: token.serialized,
      });
    }
    return capability;
  }
}

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
): Paragraph => {
  const cloned: Paragraph = { ...paragraph, ...overrides };
  if (!Reflect.deleteProperty(cloned, paragraphPropertySource)) {
    panic("Cannot detach paragraph-property source provenance from a paragraph clone");
  }
  return cloned;
};

type DocumentSourceStory = {
  content: Document["package"]["document"]["content"];
  story: ParagraphPropertySourceStory;
};

const compareCanonicalStoryKeys = (left: string, right: string): number => {
  if (left < right) {
    return -1;
  }
  return left > right ? 1 : 0;
};

const documentSourceStories = (document: Document): DocumentSourceStory[] => {
  const stories: DocumentSourceStory[] = [
    { content: document.package.document.content, story: Object.freeze({ type: "document" }) },
  ];
  for (const [relationshipId, story] of [...(document.package.headers ?? [])].toSorted(
    ([left], [right]) => compareCanonicalStoryKeys(left, right),
  )) {
    stories.push({
      content: story.content,
      story: Object.freeze({ relationshipId, type: "header" }),
    });
  }
  for (const [relationshipId, story] of [...(document.package.footers ?? [])].toSorted(
    ([left], [right]) => compareCanonicalStoryKeys(left, right),
  )) {
    stories.push({
      content: story.content,
      story: Object.freeze({ relationshipId, type: "footer" }),
    });
  }
  for (const story of [...(document.package.footnotes ?? [])].toSorted(
    (left, right) => left.id - right.id,
  )) {
    stories.push({
      content: story.content,
      story: Object.freeze({ noteId: story.id, type: "footnote" }),
    });
  }
  for (const story of [...(document.package.endnotes ?? [])].toSorted(
    (left, right) => left.id - right.id,
  )) {
    stories.push({
      content: story.content,
      story: Object.freeze({ noteId: story.id, type: "endnote" }),
    });
  }
  for (const comment of [...(document.package.document.comments ?? [])].toSorted(
    (left, right) => left.id - right.id,
  )) {
    stories.push({
      content: comment.content,
      story: Object.freeze({ commentId: comment.id, type: "comment" }),
    });
  }
  return stories;
};

const paragraphsIn = (document: Document): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  for (const { content } of documentSourceStories(document)) {
    visitDocumentStoryParagraphs(content, (paragraph) => paragraphs.push(paragraph));
  }
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
  return nodeType.create(
    {
      ...options.attrs,
      [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]:
        getParagraphPropertySourceToken(sourceParagraph)?.serialized ?? null,
    },
    options.content,
    options.marks,
  );
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
  const leftToken = readProseParagraphPropertySourceToken(left);
  const rightToken = readProseParagraphPropertySourceToken(right);
  if (leftToken.status === "invalid" || rightToken.status === "invalid") {
    throw new ParagraphPropertySourceValidationError({
      code: "invalid_token",
      message: "Cannot join a paragraph carrying a malformed paragraph-property token.",
      token: leftToken.status === "invalid" ? leftToken.raw : rightToken.raw,
    });
  }
  const leftPos = pos - left.nodeSize;
  transaction.join(pos);
  const joined = transaction.doc.nodeAt(leftPos);
  if (!joined || joined.type.name !== "paragraph") {
    panic("Joining paragraphs did not produce a paragraph");
  }
  transaction.setNodeMarkup(
    leftPos,
    undefined,
    {
      ...attrs,
      [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]:
        rightToken.status === "valid" ? rightToken.value.serialized : null,
    },
    joined.marks,
  );
  const transfers = transaction.getMeta(paragraphPropertySourceTransfersKey) ?? [];
  transaction.setMeta(paragraphPropertySourceTransfersKey, [
    ...transfers,
    {
      displacedToken: leftToken.status === "valid" ? leftToken.value.serialized : null,
      selectedToken: rightToken.status === "valid" ? rightToken.value.serialized : null,
    },
  ]);
  return transaction;
};

export const getExplicitParagraphPropertySourceTransfers = (
  transaction: Transaction,
): readonly ParagraphPropertySourceTransfer[] =>
  transaction.getMeta(paragraphPropertySourceTransfersKey) ?? [];

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
  return source.type.create(
    options.attrs ?? source.attrs,
    options.content === undefined ? source.content : options.content,
    options.marks ?? source.marks,
  );
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

/** Replace paragraph markup while retaining the caller's explicit attrs. */
export const setProseParagraphMarkupWithPropertySource = ({
  attrs,
  pos,
  transaction,
}: SetProseParagraphMarkupOptions): void => transaction.setNodeMarkup(pos, undefined, attrs);
