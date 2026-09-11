import { panic, TaggedError } from "better-result";

const PARAGRAPH_SOURCE_TOKEN_VERSION = "folio-ppr-v2";
const PARAGRAPH_SOURCE_TOKEN_PREFIX = "p2s";
const SHA256_HEX = /^[0-9a-f]{64}$/;
const SOURCE_TOKEN =
  /^p2s:(comment|document|header|footer|footnote|endnote):([^:]*):(0|[1-9a-z][0-9a-z]*)$/;

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
    case "header":
      return encodeURIComponent(story.relationshipId);
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
      let relationshipId: string;
      try {
        relationshipId = decodeURIComponent(identifier);
      } catch {
        return null;
      }
      if (!relationshipId || encodeURIComponent(relationshipId) !== identifier) {
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

export const paragraphPropertySourceStoryKey = (
  story: ParagraphPropertySourceStory,
): string => {
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

/** A source-package contract that has passed the canonical grammar check. */
export class ParagraphPropertySourceContract {
  readonly #validated = true;

  private constructor(
    readonly serialized: string,
    readonly fingerprint: string,
  ) {
    Object.freeze(this);
  }

  static fromDigest(sourceDigest: string): ParagraphPropertySourceContract {
    if (!SHA256_HEX.test(sourceDigest)) {
      panic("Paragraph-property source digest must be lowercase SHA-256 hex");
    }
    return new ParagraphPropertySourceContract(
      `${PARAGRAPH_SOURCE_TOKEN_VERSION}:${sourceDigest}`,
      sourceDigest,
    );
  }

  static read(raw: unknown): ParagraphPropertySourceAttribute<ParagraphPropertySourceContract> {
    if (raw === null || raw === undefined) {
      return { status: "absent" };
    }
    if (typeof raw !== "string") {
      return { raw, status: "invalid" };
    }
    const prefix = `${PARAGRAPH_SOURCE_TOKEN_VERSION}:`;
    const digest = raw.startsWith(prefix) ? raw.slice(prefix.length) : "";
    if (!SHA256_HEX.test(digest)) {
      return { raw, status: "invalid" };
    }
    return {
      status: "valid",
      value: new ParagraphPropertySourceContract(raw, digest),
    };
  }

  readToken(raw: unknown): ParagraphPropertySourceAttribute<ParagraphPropertySourceToken> {
    if (raw === null || raw === undefined) {
      return { status: "absent" };
    }
    if (typeof raw !== "string") {
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
    const story = readStory(storyType, storyIdentifier);
    const ordinal = Number.parseInt(serializedOrdinal, 36);
    if (!story || !Number.isSafeInteger(ordinal) || ordinal.toString(36) !== serializedOrdinal) {
      return { raw, status: "invalid" };
    }
    return {
      status: "valid",
      value: new ParagraphPropertySourceToken(
        paragraphPropertySourceTokenIssuer,
        this.fingerprint,
        raw,
        story,
        ordinal,
      ),
    };
  }

  bindStoryCensus(
    story: ParagraphPropertySourceStory,
    paragraphCount: number,
  ): ParagraphPropertySourceTokenCensus {
    return ParagraphPropertySourceTokenCensus.bind(this, story, paragraphCount);
  }

  owns(token: ParagraphPropertySourceToken): boolean {
    return this.#validated && token.belongsToContract(this.fingerprint);
  }
}

const paragraphPropertySourceTokenIssuer = Symbol("paragraphPropertySourceTokenIssuer");
type ParagraphPropertySourceTokenIssuer = typeof paragraphPropertySourceTokenIssuer;

/** A paragraph token whose syntax has passed the canonical grammar check. */
export class ParagraphPropertySourceToken {
  readonly #validated = true;
  readonly #contractFingerprint: string;

  constructor(
    issuer: ParagraphPropertySourceTokenIssuer,
    contractFingerprint: string,
    readonly serialized: string,
    readonly story: ParagraphPropertySourceStory,
    readonly ordinal: number,
  ) {
    if (issuer !== paragraphPropertySourceTokenIssuer) {
      panic("Only a paragraph-property source census may issue tokens");
    }
    this.#contractFingerprint = contractFingerprint;
    Object.freeze(this);
  }

  belongsTo(contract: ParagraphPropertySourceContract): boolean {
    return this.#validated && contract.owns(this);
  }

  belongsToStory(story: ParagraphPropertySourceStory): boolean {
    return this.#validated && paragraphPropertySourceStoriesEqual(this.story, story);
  }

  belongsToContract(fingerprint: string): boolean {
    return this.#validated && this.#contractFingerprint === fingerprint;
  }
}

/** The only token issuer: one complete, canonical story census. */
class ParagraphPropertySourceTokenCensus {
  readonly #tokens: readonly ParagraphPropertySourceToken[];

  private constructor(tokens: readonly ParagraphPropertySourceToken[]) {
    this.#tokens = Object.freeze(tokens);
    Object.freeze(this);
  }

  static bind(
    contract: ParagraphPropertySourceContract,
    story: ParagraphPropertySourceStory,
    paragraphCount: number,
  ): ParagraphPropertySourceTokenCensus {
    if (!Number.isSafeInteger(paragraphCount) || paragraphCount < 0) {
      panic("A paragraph-property source census requires a non-negative safe paragraph count");
    }
    const canonicalStory = canonicalParagraphPropertySourceStory(story);
    const tokens = Array.from({ length: paragraphCount }, (_, ordinal) =>
      new ParagraphPropertySourceToken(
        paragraphPropertySourceTokenIssuer,
        contract.fingerprint,
        `${PARAGRAPH_SOURCE_TOKEN_PREFIX}:${canonicalStory.type}:${serializedStoryIdentifier(canonicalStory)}:${ordinal.toString(36)}`,
        canonicalStory,
        ordinal,
      ),
    );
    return new ParagraphPropertySourceTokenCensus(tokens);
  }

  tokenAt(ordinal: number): ParagraphPropertySourceToken {
    const token = this.#tokens.at(ordinal);
    if (!token) {
      return panic("Paragraph-property source census ordinal is outside its bound story");
    }
    return token;
  }
}

const transientTemplateHandleIssuer = Symbol("paragraphPropertyTransientTemplateHandleIssuer");
type TransientTemplateHandleIssuer = typeof transientTemplateHandleIssuer;

export class ParagraphPropertyTransientTemplateHandle {
  readonly #owner: object;

  constructor(issuer: TransientTemplateHandleIssuer, owner: object, readonly ordinal: number) {
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
  "template_capacity_exceeded",
  "transient_state",
  "invalid_token",
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
