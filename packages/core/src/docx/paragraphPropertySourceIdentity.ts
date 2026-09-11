import { panic, TaggedError } from "better-result";

const PARAGRAPH_SOURCE_TOKEN_VERSION = "folio-ppr-v3";
const PARAGRAPH_SOURCE_TOKEN_PREFIX = "p3s";
export const PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS = 100_000;
export const PARAGRAPH_PROPERTY_SOURCE_MAX_STORIES = 10_000;
export const PARAGRAPH_PROPERTY_SOURCE_MAX_STORY_IDENTIFIER_CODE_UNITS = 1_024;
export const PARAGRAPH_PROPERTY_SOURCE_MAX_CONTRACT_CODE_UNITS = 2_000_000;
const PARAGRAPH_PROPERTY_SOURCE_MAX_SERIALIZED_TOKEN_CODE_UNITS =
  PARAGRAPH_PROPERTY_SOURCE_MAX_STORY_IDENTIFIER_CODE_UNITS * 3 + 32;
const PARAGRAPH_PROPERTY_SOURCE_MAX_SERIALIZED_ORDINAL_CODE_UNITS = (
  PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS - 1
).toString(36).length;
const SHA256_HEX = /^[0-9a-f]{64}$/;
const SOURCE_TOKEN =
  /^p3s:(comment|document|header|footer|footnote|endnote):([^:]*):(0|[1-9a-z][0-9a-z]*)$/;

/** Result of reifying serialized paragraph-source metadata at a trust boundary. */
export type ParagraphPropertySourceAttribute<T> =
  | { status: "absent" }
  | { raw: unknown; status: "invalid" }
  | { status: "valid"; value: T };

/** Stable identity of one independently editable story in a source package. */
export type ParagraphPropertySourceStory =
  | { type: "document" }
  | { relationshipId: string; type: "footer" | "header" }
  | { noteId: number; type: "endnote" | "footnote" }
  | { commentId: number; type: "comment" };

const serializedStoryIdentifier = (story: ParagraphPropertySourceStory): string => {
  switch (story.type) {
    case "document":
      return "";
    case "comment":
      return String(story.commentId);
    case "footer":
    case "header": {
      if (
        story.relationshipId.length === 0 ||
        story.relationshipId.length > PARAGRAPH_PROPERTY_SOURCE_MAX_STORY_IDENTIFIER_CODE_UNITS
      ) {
        panic("Paragraph-property relationship identity exceeds its bounded wire contract");
      }
      return encodeURIComponent(story.relationshipId);
    }
    case "endnote":
    case "footnote":
      return String(story.noteId);
    default: {
      const exhaustive: never = story;
      return exhaustive;
    }
  }
};

const readStory = (type: string, identifier: string): ParagraphPropertySourceStory | null => {
  switch (type) {
    case "document":
      return identifier === "" ? Object.freeze({ type: "document" }) : null;
    case "comment": {
      if (!/^-?(0|[1-9][0-9]*)$/.test(identifier)) {
        return null;
      }
      const commentId = Number(identifier);
      return Number.isSafeInteger(commentId) && String(commentId) === identifier
        ? Object.freeze({ commentId, type })
        : null;
    }
    case "footer":
    case "header": {
      if (
        identifier.length === 0 ||
        identifier.length > PARAGRAPH_PROPERTY_SOURCE_MAX_STORY_IDENTIFIER_CODE_UNITS * 3
      ) {
        return null;
      }
      let relationshipId: string;
      try {
        relationshipId = decodeURIComponent(identifier);
      } catch {
        return null;
      }
      if (
        !relationshipId ||
        relationshipId.length > PARAGRAPH_PROPERTY_SOURCE_MAX_STORY_IDENTIFIER_CODE_UNITS ||
        encodeURIComponent(relationshipId) !== identifier
      ) {
        return null;
      }
      return Object.freeze({ relationshipId, type });
    }
    case "endnote":
    case "footnote": {
      if (!/^-?(0|[1-9][0-9]*)$/.test(identifier)) {
        return null;
      }
      const noteId = Number(identifier);
      return Number.isSafeInteger(noteId) && String(noteId) === identifier
        ? Object.freeze({ noteId, type })
        : null;
    }
    default:
      return null;
  }
};

export const canonicalParagraphPropertySourceStory = (
  story: ParagraphPropertySourceStory,
): ParagraphPropertySourceStory => {
  const identifier = serializedStoryIdentifier(story);
  const canonical = readStory(story.type, identifier);
  if (!canonical) {
    panic("Paragraph-property source story must have a canonical identity");
  }
  return canonical;
};

export const paragraphPropertySourceStoryKey = (story: ParagraphPropertySourceStory): string => {
  const canonical = canonicalParagraphPropertySourceStory(story);
  return `${canonical.type}:${serializedStoryIdentifier(canonical)}`;
};

export const paragraphPropertySourceStoriesEqual = (
  left: ParagraphPropertySourceStory,
  right: ParagraphPropertySourceStory,
): boolean => {
  if (left.type !== right.type) {
    return false;
  }
  switch (left.type) {
    case "document":
      return true;
    case "comment":
      return right.type === left.type && right.commentId === left.commentId;
    case "footer":
    case "header":
      return right.type === left.type && right.relationshipId === left.relationshipId;
    case "endnote":
    case "footnote":
      return right.type === left.type && right.noteId === left.noteId;
    default: {
      const exhaustive: never = left;
      return exhaustive;
    }
  }
};

export const paragraphPropertySourceTokenWire = (
  story: ParagraphPropertySourceStory,
  ordinal: number,
): string => {
  if (
    !Number.isSafeInteger(ordinal) ||
    ordinal < 0 ||
    ordinal >= PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS
  ) {
    panic("Paragraph-property source token ordinal is outside its bounded wire contract");
  }
  const canonicalStory = canonicalParagraphPropertySourceStory(story);
  return `${PARAGRAPH_SOURCE_TOKEN_PREFIX}:${canonicalStory.type}:${serializedStoryIdentifier(canonicalStory)}:${ordinal.toString(36)}`;
};

type ParagraphPropertySourceCensusEntry = {
  paragraphCount: number;
  story: ParagraphPropertySourceStory;
};

type SerializedParagraphPropertySourceCensusEntry = readonly [
  type: ParagraphPropertySourceStory["type"],
  identifier: string,
  paragraphCount: number,
];

const canonicalSourceCensus = (
  entries: readonly ParagraphPropertySourceCensusEntry[],
): readonly SerializedParagraphPropertySourceCensusEntry[] => {
  if (entries.length > PARAGRAPH_PROPERTY_SOURCE_MAX_STORIES) {
    throw new ParagraphPropertySourceValidationError({
      code: "source_capacity_exceeded",
      message: "Paragraph-property source story capacity was exceeded.",
    });
  }
  let totalParagraphCount = 0;
  const seen = new Set<string>();
  const serialized: SerializedParagraphPropertySourceCensusEntry[] = [];
  for (const { paragraphCount, story } of entries) {
    if (
      !Number.isSafeInteger(paragraphCount) ||
      paragraphCount < 0 ||
      paragraphCount > PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS
    ) {
      throw new ParagraphPropertySourceValidationError({
        code: "source_capacity_exceeded",
        message: "Paragraph-property source paragraph capacity was exceeded.",
      });
    }
    totalParagraphCount += paragraphCount;
    if (totalParagraphCount > PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS) {
      throw new ParagraphPropertySourceValidationError({
        code: "source_capacity_exceeded",
        message: "Paragraph-property source paragraph capacity was exceeded.",
      });
    }
    const canonicalStory = canonicalParagraphPropertySourceStory(story);
    const key = paragraphPropertySourceStoryKey(canonicalStory);
    if (seen.has(key)) {
      throw new ParagraphPropertySourceValidationError({
        code: "ambiguous_source",
        message: "Paragraph-property source census contains duplicate story identity.",
      });
    }
    seen.add(key);
    serialized.push(
      Object.freeze([
        canonicalStory.type,
        serializedStoryIdentifier(canonicalStory),
        paragraphCount,
      ]),
    );
  }
  serialized.sort((left, right) => {
    const leftKey = `${left[0]}:${left[1]}`;
    const rightKey = `${right[0]}:${right[1]}`;
    if (leftKey < rightKey) {
      return -1;
    }
    return leftKey > rightKey ? 1 : 0;
  });
  return Object.freeze(serialized);
};

const censusCounts = (
  entries: readonly SerializedParagraphPropertySourceCensusEntry[],
): ReadonlyMap<string, number> => {
  const counts = new Map<string, number>();
  for (const [type, identifier, paragraphCount] of entries) {
    const story = readStory(type, identifier);
    if (!story) {
      panic("Canonical paragraph-property source census became invalid");
    }
    counts.set(paragraphPropertySourceStoryKey(story), paragraphCount);
  }
  return counts;
};

/** A source-package contract that has passed the canonical grammar check. */
export class ParagraphPropertySourceContract {
  readonly #validated = true;
  readonly #storyParagraphCounts: ReadonlyMap<string, number>;

  private constructor(
    readonly serialized: string,
    readonly fingerprint: string,
    entries: readonly SerializedParagraphPropertySourceCensusEntry[],
  ) {
    this.#storyParagraphCounts = censusCounts(entries);
    Object.freeze(this);
  }

  static fromSourceCensus(
    sourceDigest: string,
    entries: readonly ParagraphPropertySourceCensusEntry[],
  ): ParagraphPropertySourceContract {
    if (!SHA256_HEX.test(sourceDigest)) {
      panic("Paragraph-property source digest must be lowercase SHA-256 hex");
    }
    const census = canonicalSourceCensus(entries);
    const serialized = `${PARAGRAPH_SOURCE_TOKEN_VERSION}:${sourceDigest}:${JSON.stringify(census)}`;
    if (serialized.length > PARAGRAPH_PROPERTY_SOURCE_MAX_CONTRACT_CODE_UNITS) {
      throw new ParagraphPropertySourceValidationError({
        code: "source_capacity_exceeded",
        message: "Paragraph-property source contract capacity was exceeded.",
      });
    }
    return new ParagraphPropertySourceContract(serialized, sourceDigest, census);
  }

  static read(raw: unknown): ParagraphPropertySourceAttribute<ParagraphPropertySourceContract> {
    if (raw === null || raw === undefined) {
      return { status: "absent" };
    }
    if (typeof raw !== "string" || raw.length > PARAGRAPH_PROPERTY_SOURCE_MAX_CONTRACT_CODE_UNITS) {
      return { raw, status: "invalid" };
    }
    const prefix = `${PARAGRAPH_SOURCE_TOKEN_VERSION}:`;
    if (!raw.startsWith(prefix)) {
      return { raw, status: "invalid" };
    }
    const digestStart = prefix.length;
    const digest = raw.slice(digestStart, digestStart + 64);
    if (!SHA256_HEX.test(digest) || raw.at(digestStart + 64) !== ":") {
      return { raw, status: "invalid" };
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw.slice(digestStart + 65));
    } catch {
      return { raw, status: "invalid" };
    }
    if (!Array.isArray(parsed) || parsed.length > PARAGRAPH_PROPERTY_SOURCE_MAX_STORIES) {
      return { raw, status: "invalid" };
    }
    const census: ParagraphPropertySourceCensusEntry[] = [];
    for (const entry of parsed) {
      if (
        !Array.isArray(entry) ||
        entry.length !== 3 ||
        typeof entry[0] !== "string" ||
        typeof entry[1] !== "string" ||
        typeof entry[2] !== "number"
      ) {
        return { raw, status: "invalid" };
      }
      const story = readStory(entry[0], entry[1]);
      if (!story) {
        return { raw, status: "invalid" };
      }
      census.push({ paragraphCount: entry[2], story });
    }
    let canonical: readonly SerializedParagraphPropertySourceCensusEntry[];
    try {
      canonical = canonicalSourceCensus(census);
    } catch {
      return { raw, status: "invalid" };
    }
    if (`${prefix}${digest}:${JSON.stringify(canonical)}` !== raw) {
      return { raw, status: "invalid" };
    }
    return {
      status: "valid",
      value: new ParagraphPropertySourceContract(raw, digest, canonical),
    };
  }

  readToken(raw: unknown): ParagraphPropertySourceAttribute<ParagraphPropertySourceToken> {
    if (raw === null || raw === undefined) {
      return { status: "absent" };
    }
    if (typeof raw !== "string") {
      return { raw, status: "invalid" };
    }
    if (raw.length > PARAGRAPH_PROPERTY_SOURCE_MAX_SERIALIZED_TOKEN_CODE_UNITS) {
      return { raw, status: "invalid" };
    }
    const match = SOURCE_TOKEN.exec(raw);
    if (!match) {
      return { raw, status: "invalid" };
    }
    const storyType = match.at(1);
    const storyIdentifier = match.at(2);
    const serializedOrdinal = match.at(3);
    if (!storyType || storyIdentifier === undefined || !serializedOrdinal) {
      return { raw, status: "invalid" };
    }
    if (
      storyIdentifier.length > PARAGRAPH_PROPERTY_SOURCE_MAX_STORY_IDENTIFIER_CODE_UNITS * 3 ||
      serializedOrdinal.length > PARAGRAPH_PROPERTY_SOURCE_MAX_SERIALIZED_ORDINAL_CODE_UNITS
    ) {
      return { raw, status: "invalid" };
    }
    const story = readStory(storyType, storyIdentifier);
    const ordinal = Number.parseInt(serializedOrdinal, 36);
    if (
      !story ||
      !Number.isSafeInteger(ordinal) ||
      ordinal < 0 ||
      ordinal >= PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS ||
      ordinal.toString(36) !== serializedOrdinal ||
      ordinal >= (this.#storyParagraphCounts.get(paragraphPropertySourceStoryKey(story)) ?? 0)
    ) {
      return { raw, status: "invalid" };
    }
    return {
      status: "valid",
      value: new ParagraphPropertySourceToken(
        paragraphPropertySourceTokenIssuer,
        this,
        raw,
        story,
        ordinal,
      ),
    };
  }

  owns(token: ParagraphPropertySourceToken): boolean {
    return this.#validated && token.belongsTo(this);
  }
}

const paragraphPropertySourceTokenIssuer = Symbol("paragraphPropertySourceTokenIssuer");
type ParagraphPropertySourceTokenIssuer = typeof paragraphPropertySourceTokenIssuer;

/** A paragraph token whose syntax has passed the canonical grammar check. */
export class ParagraphPropertySourceToken {
  readonly #validated = true;
  readonly #contract: ParagraphPropertySourceContract;

  constructor(
    issuer: ParagraphPropertySourceTokenIssuer,
    contract: ParagraphPropertySourceContract,
    readonly serialized: string,
    readonly story: ParagraphPropertySourceStory,
    readonly ordinal: number,
  ) {
    if (issuer !== paragraphPropertySourceTokenIssuer) {
      panic("Only a paragraph-property source census may issue tokens");
    }
    this.#contract = contract;
    Object.freeze(this);
  }

  belongsTo(contract: ParagraphPropertySourceContract): boolean {
    return this.#validated && this.#contract === contract;
  }

  belongsToStory(story: ParagraphPropertySourceStory): boolean {
    return this.#validated && paragraphPropertySourceStoriesEqual(this.story, story);
  }
}

const transientTemplateHandleIssuer = Symbol("paragraphPropertyTransientTemplateHandleIssuer");
type TransientTemplateHandleIssuer = typeof transientTemplateHandleIssuer;

export class ParagraphPropertyTransientTemplateHandle {
  readonly #owner: object;

  constructor(
    issuer: TransientTemplateHandleIssuer,
    owner: object,
    readonly ordinal: number,
  ) {
    if (issuer !== transientTemplateHandleIssuer) {
      panic("Only a paragraph-property template store may issue transient handles");
    }
    this.#owner = owner;
    Object.freeze(this);
  }

  isValidated(): boolean {
    return this.#owner !== undefined;
  }

  belongsTo(owner: object): boolean {
    return this.#owner === owner;
  }
}

export const PARAGRAPH_PROPERTY_SOURCE_VALIDATION_CODES = [
  "ambiguous_source",
  "contract_mismatch",
  "duplicate_template_handle",
  "duplicate_token",
  "invalid_state",
  "invalid_template_handle",
  "ownership_transition_mismatch",
  "template_capacity_exceeded",
  "transient_state",
  "invalid_token",
  "source_capacity_exceeded",
  "state_capacity_exceeded",
  "unconsumed_template_handle",
  "unknown_template_handle",
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

export class ParagraphPropertyTransientTemplateResolution<Capture> {
  readonly #owner: object;
  readonly #remaining: Map<ParagraphPropertyTransientTemplateHandle, Capture>;
  readonly #consumed = new Set<ParagraphPropertyTransientTemplateHandle>();

  constructor(
    issuer: TransientTemplateHandleIssuer,
    owner: object,
    selected: ReadonlyMap<ParagraphPropertyTransientTemplateHandle, Capture>,
  ) {
    if (issuer !== transientTemplateHandleIssuer) {
      panic("Only a paragraph-property template store may begin handle resolution");
    }
    this.#owner = owner;
    this.#remaining = new Map(selected);
    Object.freeze(this);
  }

  consume(handle: ParagraphPropertyTransientTemplateHandle): Capture {
    if (!handle.isValidated() || !handle.belongsTo(this.#owner)) {
      throw new ParagraphPropertySourceValidationError({
        code: "unknown_template_handle",
        message: "A paragraph-property template handle belongs to a different store.",
      });
    }
    if (this.#consumed.has(handle)) {
      throw new ParagraphPropertySourceValidationError({
        code: "duplicate_template_handle",
        message: "A paragraph-property template handle was consumed more than once.",
      });
    }
    const capture = this.#remaining.get(handle);
    if (capture === undefined) {
      throw new ParagraphPropertySourceValidationError({
        code: "unknown_template_handle",
        message: "A paragraph-property template handle is outside this resolution.",
      });
    }
    this.#remaining.delete(handle);
    this.#consumed.add(handle);
    return capture;
  }

  assertFullyConsumed(): void {
    if (this.#remaining.size !== 0) {
      throw new ParagraphPropertySourceValidationError({
        code: "unconsumed_template_handle",
        message: "A paragraph-property template was not consumed during conversion.",
      });
    }
  }
}

/** Bounded process-local owner for opaque template captures. */
export class ParagraphPropertyTransientTemplateStore<Capture> {
  readonly #capacity: number;
  readonly #captures = new Map<ParagraphPropertyTransientTemplateHandle, Capture>();
  readonly #owner = Object.freeze({});
  #nextOrdinal = 0;

  constructor(capacity: number) {
    if (!Number.isSafeInteger(capacity) || capacity < 1) {
      panic("A paragraph-property template store requires a positive safe capacity");
    }
    this.#capacity = capacity;
  }

  registerAll(captures: readonly Capture[]): readonly ParagraphPropertyTransientTemplateHandle[] {
    if (this.#captures.size + captures.length > this.#capacity) {
      throw new ParagraphPropertySourceValidationError({
        code: "template_capacity_exceeded",
        message: "Paragraph-property template capture capacity was exceeded.",
      });
    }
    const handles = captures.map((capture) => {
      const handle = new ParagraphPropertyTransientTemplateHandle(
        transientTemplateHandleIssuer,
        this.#owner,
        this.#nextOrdinal,
      );
      this.#nextOrdinal += 1;
      this.#captures.set(handle, capture);
      return handle;
    });
    return Object.freeze(handles);
  }

  beginResolution(
    handles: readonly ParagraphPropertyTransientTemplateHandle[],
  ): ParagraphPropertyTransientTemplateResolution<Capture> {
    const selected = new Map<ParagraphPropertyTransientTemplateHandle, Capture>();
    for (const handle of handles) {
      if (!handle.isValidated() || !handle.belongsTo(this.#owner)) {
        throw new ParagraphPropertySourceValidationError({
          code: "unknown_template_handle",
          message: "A paragraph-property template handle belongs to a different store.",
        });
      }
      if (selected.has(handle)) {
        throw new ParagraphPropertySourceValidationError({
          code: "duplicate_template_handle",
          message: "A paragraph-property template handle occurs more than once.",
        });
      }
      const capture = this.#captures.get(handle);
      if (capture === undefined) {
        throw new ParagraphPropertySourceValidationError({
          code: "unknown_template_handle",
          message: "A paragraph-property template handle is unknown to this store.",
        });
      }
      selected.set(handle, capture);
    }
    return new ParagraphPropertyTransientTemplateResolution(
      transientTemplateHandleIssuer,
      this.#owner,
      selected,
    );
  }
}
