/** Remaining seams to retire before the canonical session becomes the only authority. */
export const CANONICAL_GAP = {
  commands: "command-descriptors",
  trackedTableStructure: "tracked-table-structure",
  storyContentProjection: "story-content-projection",
  dispatch: "unclassified-transactions",
  publicComments: "publicOps.comments",
  publicUnsupportedInline: "publicOps.unsupportedInline",
  commentReplacementAlignment: "comment-replacement-alignment",
  publicHeadlessSession: "publicOps.headlessSession",
} as const;

export type CanonicalGap = (typeof CANONICAL_GAP)[keyof typeof CANONICAL_GAP];

/**
 * Rules a document states for itself. A refusal under one is the correct
 * outcome and stays after the cutover; it is never a gap to retire.
 */
export const CANONICAL_RULE: {
  readonly contentControlLock: "contentControlLock";
  readonly linkSuggestionShape: "linkSuggestionShape";
} = {
  contentControlLock: "contentControlLock",
  /**
   * A suggested link change cannot leave a piece of a link holding only
   * bookmark or comment markers, nor copy content that names a record of its
   * own (preserved run markup, a marker inside a copied text box, or a note
   * holding a tracked change, comment or bookmark). The same change made
   * directly applies.
   */
  linkSuggestionShape: "linkSuggestionShape",
} as const;

export type CanonicalRule =
  (typeof CANONICAL_RULE)[keyof typeof CANONICAL_RULE];

/** The editor message, in the folio catalog, a host shows for a refusal under each rule. */
export const CANONICAL_RULE_MESSAGE_KEYS: {
  readonly contentControlLock: "editor.contentLocked";
  readonly linkSuggestionShape: "editor.linkSuggestionUnavailable";
} = {
  contentControlLock: "editor.contentLocked",
  linkSuggestionShape: "editor.linkSuggestionUnavailable",
} as const satisfies Record<CanonicalRule, string>;

/**
 * Conditions of the session or of the request, not of the document: the
 * session holds the capability, but cannot make this change now or was not
 * asked for one it can make. A refusal under one is correct and is never a gap
 * to retire.
 */
export const CANONICAL_CONDITION: {
  readonly notReady: "notReady";
  readonly notEditable: "notEditable";
  readonly composing: "composing";
  readonly invalidRequest: "invalidRequest";
  readonly idSpace: "idSpace";
} = {
  /** No session holds a document yet, or the editor showing it has closed. */
  notReady: "notReady",
  /** The editor is read-only. */
  notEditable: "notEditable",
  /** An input composition must finish first. */
  composing: "composing",
  /**
   * The request names nothing in the document or describes no change the
   * document allows, such as a reply to a missing comment.
   */
  invalidRequest: "invalidRequest",
  /** No id is left to mint, in the document or in this author's lease. */
  idSpace: "idSpace",
} as const;

export type CanonicalCondition =
  (typeof CANONICAL_CONDITION)[keyof typeof CANONICAL_CONDITION];

/**
 * The editor message, in the folio catalog, a host shows for a refusal under
 * any gap: the change cannot be made here yet.
 */
export const CANONICAL_GAP_MESSAGE_KEY = "editor.changeUnavailable" as const;

/**
 * The editor message, in the folio catalog, a host shows for a refusal under
 * any condition: the change was not made.
 */
export const CANONICAL_CONDITION_MESSAGE_KEY =
  "editor.changeNotApplied" as const;

/** A catalog key a refusal notice reads. */
export type CanonicalRefusalMessageKey =
  | (typeof CANONICAL_RULE_MESSAGE_KEYS)[CanonicalRule]
  | typeof CANONICAL_GAP_MESSAGE_KEY
  | typeof CANONICAL_CONDITION_MESSAGE_KEY;

/**
 * What a host shows for a refusal, always from the folio catalog, which the
 * host translates: the editor's own message for a rule the document states,
 * that the change was not made for a condition, and that the change cannot be
 * made here yet for a gap. The session's own message describes the refusal
 * for logs and `onError`, never for the user.
 */
export const canonicalRefusalNotice = (
  ground: CanonicalRefusalGround,
  translate: (key: CanonicalRefusalMessageKey) => string,
): string => {
  if (ground.rule !== undefined)
    return translate(CANONICAL_RULE_MESSAGE_KEYS[ground.rule]);
  if (ground.condition !== undefined)
    return translate(CANONICAL_CONDITION_MESSAGE_KEY);
  return translate(CANONICAL_GAP_MESSAGE_KEY);
};

/**
 * The ground for a refusal: a capability the session does not have yet, a
 * rule the document or its review records state, or a condition of the
 * session or the request. A host tells "not supported yet" from "not allowed"
 * from "not now" by which one is present.
 */
export type CanonicalRefusalGround =
  | { gap: CanonicalGap; rule?: never; condition?: never }
  | { rule: CanonicalRule; gap?: never; condition?: never }
  | { condition: CanonicalCondition; gap?: never; rule?: never };

type CanonicalCapability = {
  kind: "routing" | "refusal";
  adapters: readonly ("react" | "vue")[];
  summary: string;
  remainingCommands?: { registry: "extension-commands" };
  owner: "document-operations" | "controller" | "adapters" | "ai-edits";
};

export const CANONICAL_CAPABILITIES: {
  readonly "command-descriptors": {
    readonly remainingCommands: {
      readonly registry: "extension-commands";
    };
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Commands without canonical descriptors retain PM probing and hit the dispatch boundary.";
  };
  readonly "tracked-table-structure": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Suggested row, column and table changes refuse content the tracked table records cannot express; a dragged table refuses to move.";
  };
  readonly "story-content-projection": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Unsupported inline, field and review content cannot be projected after canonical edits; pasted embedded blocks, comment and note references copied outside the editor, and compositions over such content refuse.";
  };
  readonly "unclassified-transactions": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "A host plugin or raw dispatch that changes the document is refused, naming the transaction's meta keys; hosts express edits as document operations (applyDocumentOperations).";
  };
  readonly "publicOps.comments": {
    readonly owner: "document-operations";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Public comment edits refuse when their anchors or revision structure cannot be represented.";
  };
  readonly "publicOps.unsupportedInline": {
    readonly owner: "document-operations";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Unsupported public payloads refuse with a typed compiler capability gap.";
  };
  readonly "comment-replacement-alignment": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Replacements refuse when preserving aligned comments needs anchor transitions or an unproven collapsed-range policy.";
  };
  readonly "publicOps.headlessSession": {
    readonly owner: "document-operations";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "A reviewer refuses to resolve one story of a suggestion that reaches several, which resolves only whole.";
  };
} = {
  [CANONICAL_GAP.commands]: {
    remainingCommands: { registry: "extension-commands" },
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Commands without canonical descriptors retain PM probing and hit the dispatch boundary.",
  },
  [CANONICAL_GAP.trackedTableStructure]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Suggested row, column and table changes refuse content the tracked table records cannot express; a dragged table refuses to move.",
  },
  [CANONICAL_GAP.storyContentProjection]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Unsupported inline, field and review content cannot be projected after canonical edits; pasted embedded blocks, comment and note references copied outside the editor, and compositions over such content refuse.",
  },
  [CANONICAL_GAP.dispatch]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "A host plugin or raw dispatch that changes the document is refused, naming the transaction's meta keys; hosts express edits as document operations (applyDocumentOperations).",
  },
  [CANONICAL_GAP.publicComments]: {
    owner: "document-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Public comment edits refuse when their anchors or revision structure cannot be represented.",
  },
  [CANONICAL_GAP.publicUnsupportedInline]: {
    owner: "document-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Unsupported public payloads refuse with a typed compiler capability gap.",
  },
  [CANONICAL_GAP.commentReplacementAlignment]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Replacements refuse when preserving aligned comments needs anchor transitions or an unproven collapsed-range policy.",
  },
  [CANONICAL_GAP.publicHeadlessSession]: {
    owner: "document-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "A reviewer refuses to resolve one story of a suggestion that reaches several, which resolves only whole.",
  },
} as const satisfies Record<CanonicalGap, CanonicalCapability>;

/**
 * The engine a headless reviewer holds its document in: the canonical
 * session, as the editors do.
 */
export type HeadlessReviewerSession = "canonical";

export const canonicalRefusalMessage = (
  gap: CanonicalGap,
  message: string,
): string => `[${gap}] ${message}`;
