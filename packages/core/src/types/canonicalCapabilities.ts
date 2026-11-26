/** Remaining seams to retire before the canonical session becomes the only authority. */
export const CANONICAL_GAP = {
  authorityRouting: "authority-routing",
  commands: "command-descriptors",
  trackedTableStructure: "tracked-table-structure",
  storyContentProjection: "story-content-projection",
  comments: "comment-model-edits",
  modelEdits: "direct-model-edits",
  sectionProperties: "section-properties",
  dispatch: "unclassified-transactions",
  saveSerialization: "save-serialization",
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
 * The editor message, in the folio catalog, a host shows for a refusal under
 * any gap: the change cannot be made here yet.
 */
export const CANONICAL_GAP_MESSAGE_KEY = "editor.changeUnavailable" as const;

/** A catalog key a refusal notice reads. */
export type CanonicalRefusalMessageKey =
  | (typeof CANONICAL_RULE_MESSAGE_KEYS)[CanonicalRule]
  | typeof CANONICAL_GAP_MESSAGE_KEY;

/**
 * What a host shows for a refusal, always from the folio catalog, which the
 * host translates: the editor's own message for a rule the document states,
 * and that the change cannot be made here yet for a gap. The session's own
 * message describes the refusal for logs and `onError`, never for the user.
 */
export const canonicalRefusalNotice = (
  ground: CanonicalRefusalGround,
  translate: (key: CanonicalRefusalMessageKey) => string,
): string =>
  translate(
    ground.rule === undefined
      ? CANONICAL_GAP_MESSAGE_KEY
      : CANONICAL_RULE_MESSAGE_KEYS[ground.rule],
  );

/**
 * The ground for a refusal: a capability the session does not have yet, or a
 * rule the document or its review records state. A host tells "not supported yet" from "not
 * allowed" by which one is present.
 */
export type CanonicalRefusalGround =
  { gap: CanonicalGap; rule?: never } | { rule: CanonicalRule; gap?: never };

type CanonicalCapability = {
  kind: "routing" | "refusal";
  adapters: readonly ("react" | "vue")[];
  summary: string;
  remainingCommands?: { registry: "extension-commands" };
  owner: "document-operations" | "controller" | "adapters" | "ai-edits";
};

export const CANONICAL_CAPABILITIES: {
  readonly "authority-routing": {
    readonly owner: "controller";
    readonly kind: "routing";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Explicit session selection and legacy authority routing.";
  };
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
  readonly "comment-model-edits": {
    readonly owner: "adapters";
    readonly kind: "routing";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Legacy sessions retain comment model edits; canonical comments use semantic operations.";
  };
  readonly "direct-model-edits": {
    readonly owner: "adapters";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Direct model and pending-suggestion snapshots bypass the canonical journal.";
  };
  readonly "section-properties": {
    readonly owner: "controller";
    readonly kind: "routing";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "The shared controller journals section properties; legacy sessions keep model changes.";
  };
  readonly "unclassified-transactions": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "A host plugin or raw dispatch that changes the document is refused, naming the transaction's meta keys; hosts express edits as document operations (applyDocumentOperations).";
  };
  readonly "save-serialization": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Canonical save waits for composition to finish and reports package serialization refusals and fidelity fallbacks.";
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
  [CANONICAL_GAP.authorityRouting]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    summary: "Explicit session selection and legacy authority routing.",
  },
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
  [CANONICAL_GAP.comments]: {
    owner: "adapters",
    kind: "routing",
    adapters: ["react", "vue"],
    summary:
      "Legacy sessions retain comment model edits; canonical comments use semantic operations.",
  },
  [CANONICAL_GAP.modelEdits]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Direct model and pending-suggestion snapshots bypass the canonical journal.",
  },
  [CANONICAL_GAP.sectionProperties]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    summary:
      "The shared controller journals section properties; legacy sessions keep model changes.",
  },
  [CANONICAL_GAP.dispatch]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "A host plugin or raw dispatch that changes the document is refused, naming the transaction's meta keys; hosts express edits as document operations (applyDocumentOperations).",
  },
  [CANONICAL_GAP.saveSerialization]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Canonical save waits for composition to finish and reports package serialization refusals and fidelity fallbacks.",
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
