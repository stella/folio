import { panic, TaggedError } from "better-result";

const PARAGRAPH_SOURCE_TOKEN_VERSION = "folio-ppr-v2";
const PARAGRAPH_SOURCE_TOKEN_PREFIX = "p2d";
const SHA256_HEX = /^[0-9a-f]{64}$/;
const SOURCE_TOKEN =
  /^p2d:([0-9a-f]{64}):(comment|document|header|footer|footnote|endnote):([^:]*):(0|[1-9a-z][0-9a-z]*)$/;
const TRANSIENT_TEMPLATE_HANDLE = /^folio-ppr-template-v1:(0|[1-9a-z][0-9a-z]*)$/;

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

export const paragraphPropertySourceStoryKey = (
  story: ParagraphPropertySourceStory,
): string => `${story.type}:${serializedStoryIdentifier(story)}`;

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

  owns(token: ParagraphPropertySourceToken): boolean {
    return this.#validated && this.fingerprint === token.fingerprint;
  }
}

/** A paragraph token whose syntax has passed the canonical grammar check. */
export class ParagraphPropertySourceToken {
  readonly #validated = true;

  private constructor(
    readonly serialized: string,
    readonly fingerprint: string,
    readonly story: ParagraphPropertySourceStory,
    readonly ordinal: number,
  ) {
    Object.freeze(this);
  }

  static forOrdinal(
    contract: ParagraphPropertySourceContract,
    story: ParagraphPropertySourceStory,
    ordinal: number,
  ): ParagraphPropertySourceToken {
    if (!Number.isSafeInteger(ordinal) || ordinal < 0) {
      panic("Paragraph-property source ordinal must be a non-negative safe integer");
    }
    return new ParagraphPropertySourceToken(
      `${PARAGRAPH_SOURCE_TOKEN_PREFIX}:${contract.fingerprint}:${story.type}:${serializedStoryIdentifier(story)}:${ordinal.toString(36)}`,
      contract.fingerprint,
      Object.freeze({ ...story }),
      ordinal,
    );
  }

  static read(raw: unknown): ParagraphPropertySourceAttribute<ParagraphPropertySourceToken> {
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
    const fingerprint = match.at(1);
    const storyType = match.at(2);
    const storyIdentifier = match.at(3);
    const serializedOrdinal = match.at(4);
    if (!fingerprint || !storyType || storyIdentifier === undefined || !serializedOrdinal) {
      return { raw, status: "invalid" };
    }
    const story = readStory(storyType, storyIdentifier);
    const ordinal = Number.parseInt(serializedOrdinal, 36);
    if (!story || !Number.isSafeInteger(ordinal) || ordinal.toString(36) !== serializedOrdinal) {
      return { raw, status: "invalid" };
    }
    return {
      status: "valid",
      value: new ParagraphPropertySourceToken(raw, fingerprint, story, ordinal),
    };
  }

  belongsTo(contract: ParagraphPropertySourceContract): boolean {
    return this.#validated && contract.owns(this);
  }

  belongsToStory(story: ParagraphPropertySourceStory): boolean {
    return this.#validated && paragraphPropertySourceStoriesEqual(this.story, story);
  }
}

/** One conversion-scoped reference to an opaque target-template capture. */
export class ParagraphPropertyTransientTemplateHandle {
  readonly #validated = true;

  private constructor(
    readonly serialized: string,
    readonly ordinal: number,
  ) {
    Object.freeze(this);
  }

  static forOrdinal(ordinal: number): ParagraphPropertyTransientTemplateHandle {
    if (!Number.isSafeInteger(ordinal) || ordinal < 0) {
      panic("Paragraph-property template ordinal must be a non-negative safe integer");
    }
    return new ParagraphPropertyTransientTemplateHandle(
      `folio-ppr-template-v1:${ordinal.toString(36)}`,
      ordinal,
    );
  }

  static read(
    raw: unknown,
  ): ParagraphPropertySourceAttribute<ParagraphPropertyTransientTemplateHandle> {
    if (raw === null || raw === undefined) {
      return { status: "absent" };
    }
    if (typeof raw !== "string") {
      return { raw, status: "invalid" };
    }
    const match = TRANSIENT_TEMPLATE_HANDLE.exec(raw);
    if (!match) {
      return { raw, status: "invalid" };
    }
    const serializedOrdinal = match.at(1);
    if (!serializedOrdinal) {
      return { raw, status: "invalid" };
    }
    const ordinal = Number.parseInt(serializedOrdinal, 36);
    if (!Number.isSafeInteger(ordinal) || ordinal.toString(36) !== serializedOrdinal) {
      return { raw, status: "invalid" };
    }
    return {
      status: "valid",
      value: new ParagraphPropertyTransientTemplateHandle(raw, ordinal),
    };
  }

  isValidated(): boolean {
    return this.#validated;
  }
}

export const PARAGRAPH_PROPERTY_SOURCE_VALIDATION_CODES = [
  "ambiguous_source",
  "contract_mismatch",
  "duplicate_template_handle",
  "duplicate_token",
  "invalid_template_handle",
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
