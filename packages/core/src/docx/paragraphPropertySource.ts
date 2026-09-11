import { panic } from "better-result";
import type { Fragment, Mark, Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

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
  PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS,
  PARAGRAPH_PROPERTY_SOURCE_MAX_STORIES,
  ParagraphPropertySourceToken,
  ParagraphPropertyTransientTemplateResolution,
  ParagraphPropertyTransientTemplateStore,
  ParagraphPropertyTransientTemplateHandle,
  ParagraphPropertySourceValidationError,
  paragraphPropertySourceTokenWire,
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

type ParagraphPropertyCaptureAxis =
  | { type: "none" }
  | { capture: ParagraphPropertyCapture; type: "captured" };

type ParagraphPropertyOwnershipAxis =
  | { type: "none" }
  | { type: "parsed-unbound" }
  | { token: ParagraphPropertySourceToken; type: "imported" };

type ParagraphPropertySourceBinding = {
  capture: ParagraphPropertyCaptureAxis;
  ownership: ParagraphPropertyOwnershipAxis;
};

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
    !("capture" in value) ||
    !("ownership" in value) ||
    !Object.isFrozen(value)
  ) {
    return false;
  }
  const capture = value.capture;
  const ownership = value.ownership;
  if (
    typeof capture !== "object" ||
    capture === null ||
    typeof ownership !== "object" ||
    ownership === null ||
    !("type" in capture) ||
    !("type" in ownership) ||
    !Object.isFrozen(capture) ||
    !Object.isFrozen(ownership)
  ) {
    return false;
  }
  if (
    capture.type !== "none" &&
    (capture.type !== "captured" ||
      !("capture" in capture) ||
      !isParagraphPropertyCapture(capture.capture))
  ) {
    return false;
  }
  switch (ownership.type) {
    case "none":
      return true;
    case "parsed-unbound":
      return capture.type === "captured";
    case "imported":
      return (
        capture.type === "captured" &&
        "token" in ownership &&
        ownership.token instanceof ParagraphPropertySourceToken
      );
    default:
      return false;
  }
};

const bindingWithCapture = (
  capture: ParagraphPropertyCapture,
  ownership: ParagraphPropertyOwnershipAxis,
): ParagraphPropertySourceBinding =>
  Object.freeze({
    capture: Object.freeze({ capture, type: "captured" }),
    ownership: Object.freeze(ownership),
  });

const editorCreatedBinding = (): ParagraphPropertySourceBinding =>
  Object.freeze({
    capture: Object.freeze({ type: "none" }),
    ownership: Object.freeze({ type: "none" }),
  });

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
    descriptor.configurable !== (binding.ownership.type === "parsed-unbound")
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
      configurable: binding.ownership.type === "parsed-unbound",
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
  if (existing) {
    panic("A paragraph-property capture can only be assigned once");
  }
  setParagraphPropertySourceBinding(
    paragraph,
    bindingWithCapture(capture, { type: "parsed-unbound" }),
  );
};

/** Record that a parsed paragraph had no authored `w:pPr` element. */
export const assignAbsentParagraphPropertySource = (paragraph: Paragraph): void => {
  if (paragraphPropertySourceBinding(paragraph)) {
    panic("A paragraph-property capture can only be assigned once");
  }
  const fingerprint = immutableFingerprint(paragraphPropertySourceFingerprintFromParts({}, {}));
  const capture = Object.freeze({
    fingerprint,
    fingerprintJson: canonicalParagraphPropertySourceFingerprintJson(fingerprint),
    type: "absent",
  }) satisfies ParagraphPropertyCapture;
  setParagraphPropertySourceBinding(
    paragraph,
    bindingWithCapture(capture, { type: "parsed-unbound" }),
  );
};

export const getParagraphPropertySource = (
  paragraph: Paragraph,
): ParagraphPropertySource | undefined => {
  const binding = paragraphPropertySourceBinding(paragraph);
  if (!binding || binding.capture.type === "none" || binding.capture.capture.type === "absent") {
    return undefined;
  }
  return binding.capture.capture;
};

/** Authored `w:pPr` before style and numbering defaults are materialized. */
export const getParagraphAuthoredPPr = (
  paragraph: Paragraph,
): AuthoredParagraphProperties | undefined => {
  const binding = paragraphPropertySourceBinding(paragraph);
  if (!binding || binding.capture.type === "none") {
    return undefined;
  }
  return structuredClone(binding.capture.capture.fingerprint.pPrBase);
};

/** Exact modeled properties captured from the paragraph's authored `w:pPr`. */
export const getParagraphPropertySourceFingerprint = (
  paragraph: Paragraph,
): ParagraphPropertySourceFingerprint | undefined => {
  const binding = paragraphPropertySourceBinding(paragraph);
  return binding?.capture.type === "captured" ? binding.capture.capture.fingerprint : undefined;
};

/** Mark a model paragraph as deliberately created outside the parsed source census. */
export const assignEditorCreatedParagraphPropertySource = (paragraph: Paragraph): void => {
  if (paragraphPropertySourceBinding(paragraph)) {
    panic("Paragraph-property provenance can only be assigned once");
  }
  setParagraphPropertySourceBinding(paragraph, editorCreatedBinding());
};

const tokenFromBinding = (
  binding: ParagraphPropertySourceBinding | undefined,
): ParagraphPropertySourceToken | null => {
  switch (binding?.ownership.type) {
    case "imported":
      return binding.ownership.token;
    case "parsed-unbound":
    case "none":
    case undefined:
      return null;
    default: {
      const exhaustive: never = binding.ownership;
      return exhaustive;
    }
  }
};

/** Copy the captured `w:pPr` without claiming the source paragraph's durable identity. */
export const copyParagraphPropertyCapture = (target: Paragraph, source: Paragraph): void => {
  const sourceBinding = paragraphPropertySourceBinding(source);
  if (!sourceBinding || sourceBinding.capture.type === "none") {
    return;
  }
  setParagraphPropertySourceBinding(
    target,
    bindingWithCapture(sourceBinding.capture.capture, { type: "none" }),
  );
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
    setParagraphPropertySourceBinding(target, bindingWithCapture(capture, { type: "none" }));
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
      capabilities.map((capability) => capability.capture(paragraphPropertyTemplateCaptureIssuer)),
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
  setParagraphPropertySourceBinding(target, binding);
};

/** Bind every parsed story paragraph to one exact source package. */
export const assignDocumentParagraphPropertySourceContract = (
  document: Document,
  sourceDigest: string,
): void => {
  if (Object.hasOwn(document, documentParagraphPropertySourceContract)) {
    panic("A document paragraph-property source contract can only be assigned once");
  }
  const seenStories = new Set<string>();
  const seenParagraphs = new WeakSet<Paragraph>();
  const bodyParagraphs = new WeakSet<Paragraph>();
  const census: {
    capture: ParagraphPropertyCapture;
    ordinal: number;
    paragraph: Paragraph;
    story: ParagraphPropertySourceStory;
  }[] = [];
  const sourceCensus: { paragraphCount: number; story: ParagraphPropertySourceStory }[] = [];
  let totalParagraphCount = 0;
  if (!Object.isExtensible(document)) {
    panic("A document must be extensible before binding paragraph-property provenance");
  }
  // The traversal is part of the v3 durable identity contract. Any ordering
  // change requires a token-version bump and collaboration reseed.
  for (const { content, story } of documentSourceStories(document)) {
    if (seenStories.size >= PARAGRAPH_PROPERTY_SOURCE_MAX_STORIES) {
      throw new ParagraphPropertySourceValidationError({
        code: "source_capacity_exceeded",
        message: "Paragraph-property source story capacity was exceeded.",
      });
    }
    const storyKey = paragraphPropertySourceStoryKey(story);
    if (seenStories.has(storyKey)) {
      panic("A document contains duplicate paragraph-property story identity", { storyKey });
    }
    seenStories.add(storyKey);
    const paragraphs: Paragraph[] = [];
    visitDocumentStoryParagraphs(content, (paragraph) => paragraphs.push(paragraph));
    totalParagraphCount += paragraphs.length;
    if (totalParagraphCount > PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS) {
      throw new ParagraphPropertySourceValidationError({
        code: "source_capacity_exceeded",
        message: "Paragraph-property source paragraph capacity was exceeded.",
      });
    }
    sourceCensus.push({ paragraphCount: paragraphs.length, story });
    for (const [ordinal, paragraph] of paragraphs.entries()) {
      if (seenParagraphs.has(paragraph)) {
        panic("A paragraph cannot belong to more than one source story", { storyKey });
      }
      seenParagraphs.add(paragraph);
      if (story.type === "document") {
        bodyParagraphs.add(paragraph);
      }
      const binding = paragraphPropertySourceBinding(paragraph);
      if (binding?.ownership.type !== "parsed-unbound" || binding.capture.type !== "captured") {
        panic("A document paragraph must have exactly one parsed property capture before binding");
      }
      const descriptor = Object.getOwnPropertyDescriptor(paragraph, paragraphPropertySource);
      if (!descriptor?.configurable) {
        panic("A parsed paragraph capture cannot transition to its durable source identity");
      }
      census.push({
        capture: binding.capture.capture,
        ordinal,
        paragraph,
        story,
      });
    }
  }
  for (const section of document.package.document.sections ?? []) {
    visitDocumentStoryParagraphs(section.content, (paragraph) => {
      if (!bodyParagraphs.has(paragraph)) {
        panic("A derived document section must alias paragraphs from the body story");
      }
    });
  }

  const contract = ParagraphPropertySourceContract.fromSourceCensus(sourceDigest, sourceCensus);
  const boundCensus = census.map(({ capture, ordinal, paragraph, story }) => {
    const token = contract.readToken(paragraphPropertySourceTokenWire(story, ordinal));
    if (token.status !== "valid") {
      return panic("A canonical paragraph-property source token failed contract reification");
    }
    return { capture, paragraph, token: token.value };
  });
  setDocumentParagraphPropertySourceContract(document, contract);
  for (const { capture, paragraph, token } of boundCensus) {
    setParagraphPropertySourceBinding(
      paragraph,
      bindingWithCapture(capture, { token, type: "imported" }),
    );
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
): ParagraphPropertySourceToken | undefined =>
  tokenFromBinding(paragraphPropertySourceBinding(paragraph)) ?? undefined;

const indexParagraphPropertySources = (
  content: BlockContent[],
  story: ParagraphPropertySourceStory,
  contract: ParagraphPropertySourceContract,
): ReadonlyMap<string, ParagraphPropertyTemplateCapture> => {
  const sources = new Map<string, ParagraphPropertyTemplateCapture>();
  visitDocumentStoryParagraphs(content, (paragraph) => {
    const binding = paragraphPropertySourceBinding(paragraph);
    if (binding?.ownership.type === "none") {
      return;
    }
    if (binding?.ownership.type === "parsed-unbound") {
      throw new ParagraphPropertySourceValidationError({
        code: "invalid_token",
        message: "A bound source story contains unbound parsed provenance.",
      });
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
    const capture = binding?.capture.type === "captured" ? binding.capture.capture : undefined;
    if (!capture || binding.ownership.type !== "imported") {
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
    setParagraphPropertySourceBinding(
      target,
      bindingWithCapture(capability.capture(paragraphPropertyTemplateCaptureIssuer), {
        token,
        type: "imported",
      }),
    );
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
