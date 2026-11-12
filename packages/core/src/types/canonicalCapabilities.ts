/** Remaining seams to retire before the canonical session becomes the only authority. */
export const CANONICAL_GAP = {
  authorityRouting: "authority-routing",
  commands: "command-descriptors",
  trackedHyperlinkResolution: "tracked-hyperlink-resolution",
  trackedTableStructure: "tracked-table-structure",
  storyContentProjection: "story-content-projection",
  comments: "comment-model-edits",
  modelEdits: "direct-model-edits",
  sectionProperties: "section-properties",
  watermark: "watermark-model-edits",
  collaboration: "collaboration-session",
  dispatch: "unclassified-transactions",
  save: "pm-save-projection",
  resourceReplacement: "canonical-resource-replacement",
  history: "pm-history",
  suggestionPlugin: "pm-suggestion-plugin",
  paragraphTracker: "pm-paragraph-tracker",
  tableGeometry: "pm-table-geometry",
  paragraphIdentity: "pm-paragraph-identity",
  documentOperations: "pm-document-operations",
  aiSnapshots: "pm-ai-snapshots",
  publicComments: "publicOps.comments",
  publicUnsupportedInline: "publicOps.unsupportedInline",
  publicTableStructure: "publicOps.tableStructure",
  publicPendingParagraphMarkProperties:
    "publicOps.pendingParagraphMarkProperties",
  publicHeadlessSession: "publicOps.headlessSession",
} as const;

export type CanonicalGap = (typeof CANONICAL_GAP)[keyof typeof CANONICAL_GAP];

/**
 * Rules a document states for itself. A refusal under one is the correct
 * outcome and stays after the cutover; it is never a gap to retire.
 */
export const CANONICAL_RULE: {
  readonly contentControlLock: "contentControlLock";
} = {
  contentControlLock: "contentControlLock",
} as const;

export type CanonicalRule =
  (typeof CANONICAL_RULE)[keyof typeof CANONICAL_RULE];

/** The editor message, in the folio catalog, a host shows for a refusal under each rule. */
export const CANONICAL_RULE_MESSAGE_KEYS: {
  readonly contentControlLock: "editor.contentLocked";
} = {
  contentControlLock: "editor.contentLocked",
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
 * rule the document states. A host tells "not supported yet" from "not
 * allowed" by which one is present.
 */
export type CanonicalRefusalGround =
  { gap: CanonicalGap; rule?: never } | { rule: CanonicalRule; gap?: never };

type CanonicalCapability = {
  kind: "routing" | "refusal" | "mutation-source";
  adapters: readonly ("react" | "vue")[];
  summary: string;
  remainingCommands?: { registry: "extension-commands" };
} & (
  | {
      owner: "document-operations";
      legacySessionMutation:
        "pm-public-operations" | "pm-headless-reviewer" | "canonical";
    }
  | { owner: "controller" | "adapters" | "prosemirror" | "ai-edits" }
);

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
  readonly "tracked-hyperlink-resolution": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Hyperlink and TOC suggestions and edits intersecting review identities require serializable wrapper review provenance.";
  };
  readonly "tracked-table-structure": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Suggested row, column and table changes refuse content the tracked table records cannot express.";
  };
  readonly "story-content-projection": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Unsupported inline, field and review content cannot be projected after canonical edits.";
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
  readonly "watermark-model-edits": {
    readonly owner: "adapters";
    readonly kind: "refusal";
    readonly adapters: readonly ["vue"];
    readonly summary: "Vue watermark mutations still change the model directly.";
  };
  readonly "collaboration-session": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Collaboration cannot activate a canonical session.";
  };
  readonly "unclassified-transactions": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Unclassified native and plugin mutations are refused at the projection boundary.";
  };
  readonly "pm-save-projection": {
    readonly owner: "controller";
    readonly kind: "routing";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Legacy save and read paths reconstruct documents from PM.";
  };
  readonly "canonical-resource-replacement": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Canonical save refuses existing style, style-default, and media replacements the package writers cannot express.";
  };
  readonly "pm-history": {
    readonly owner: "prosemirror";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "PM history remains installed for legacy sessions.";
  };
  readonly "pm-suggestion-plugin": {
    readonly owner: "prosemirror";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "The legacy suggestion plugin produces PM transactions.";
  };
  readonly "pm-paragraph-tracker": {
    readonly owner: "prosemirror";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Dirty tracking and section bookkeeping still observe PM transactions.";
  };
  readonly "pm-table-geometry": {
    readonly owner: "prosemirror";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Table geometry, resize and repair still produce PM transactions.";
  };
  readonly "pm-paragraph-identity": {
    readonly owner: "prosemirror";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Legacy paragraph identity repair remains a PM append transaction.";
  };
  readonly "pm-document-operations": {
    readonly owner: "document-operations";
    readonly legacySessionMutation: "pm-public-operations";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Public document operations still execute against a held PM view.";
  };
  readonly "publicOps.comments": {
    readonly owner: "document-operations";
    readonly legacySessionMutation: "pm-public-operations";
    readonly kind: "routing";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Canonical public comments compile to semantic operations; legacy sessions retain PM comments.";
  };
  readonly "publicOps.unsupportedInline": {
    readonly owner: "document-operations";
    readonly legacySessionMutation: "pm-public-operations";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Unsupported public payloads refuse with a typed compiler capability gap.";
  };
  readonly "publicOps.tableStructure": {
    readonly owner: "document-operations";
    readonly legacySessionMutation: "pm-public-operations";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Public table structure operations require canonical table operation descriptors.";
  };
  readonly "publicOps.pendingParagraphMarkProperties": {
    readonly owner: "document-operations";
    readonly legacySessionMutation: "pm-public-operations";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Separate paragraph property edits refuse while the paragraph mark is pending deletion; atomic merge properties remain supported.";
  };
  readonly "publicOps.headlessSession": {
    readonly owner: "document-operations";
    readonly legacySessionMutation: "pm-headless-reviewer";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "The headless reviewer has no canonical session selector while its mutations remain PM-owned.";
  };
  readonly "pm-ai-snapshots": {
    readonly owner: "ai-edits";
    readonly kind: "mutation-source";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Headless AI edits still use PM execution and snapshot undo.";
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
  [CANONICAL_GAP.trackedHyperlinkResolution]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Hyperlink and TOC suggestions and edits intersecting review identities require serializable wrapper review provenance.",
  },
  [CANONICAL_GAP.trackedTableStructure]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Suggested row, column and table changes refuse content the tracked table records cannot express.",
  },
  [CANONICAL_GAP.storyContentProjection]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Unsupported inline, field and review content cannot be projected after canonical edits.",
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
  [CANONICAL_GAP.watermark]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["vue"],
    summary: "Vue watermark mutations still change the model directly.",
  },
  [CANONICAL_GAP.collaboration]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary: "Collaboration cannot activate a canonical session.",
  },
  [CANONICAL_GAP.dispatch]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Unclassified native and plugin mutations are refused at the projection boundary.",
  },
  [CANONICAL_GAP.save]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    summary: "Legacy save and read paths reconstruct documents from PM.",
  },
  [CANONICAL_GAP.resourceReplacement]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Canonical save refuses existing style, style-default, and media replacements the package writers cannot express.",
  },
  [CANONICAL_GAP.history]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "PM history remains installed for legacy sessions.",
  },
  [CANONICAL_GAP.suggestionPlugin]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "The legacy suggestion plugin produces PM transactions.",
  },
  [CANONICAL_GAP.paragraphTracker]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary:
      "Dirty tracking and section bookkeeping still observe PM transactions.",
  },
  [CANONICAL_GAP.tableGeometry]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Table geometry, resize and repair still produce PM transactions.",
  },
  [CANONICAL_GAP.paragraphIdentity]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary:
      "Legacy paragraph identity repair remains a PM append transaction.",
  },
  [CANONICAL_GAP.documentOperations]: {
    owner: "document-operations",
    legacySessionMutation: "pm-public-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Public document operations still execute against a held PM view.",
  },
  [CANONICAL_GAP.publicComments]: {
    owner: "document-operations",
    legacySessionMutation: "pm-public-operations",
    kind: "routing",
    adapters: ["react", "vue"],
    summary:
      "Canonical public comments compile to semantic operations; legacy sessions retain PM comments.",
  },
  [CANONICAL_GAP.publicUnsupportedInline]: {
    owner: "document-operations",
    legacySessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Unsupported public payloads refuse with a typed compiler capability gap.",
  },
  [CANONICAL_GAP.publicTableStructure]: {
    owner: "document-operations",
    legacySessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Public table structure operations require canonical table operation descriptors.",
  },
  [CANONICAL_GAP.publicPendingParagraphMarkProperties]: {
    owner: "document-operations",
    legacySessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Separate paragraph property edits refuse while the paragraph mark is pending deletion; atomic merge properties remain supported.",
  },
  [CANONICAL_GAP.publicHeadlessSession]: {
    owner: "document-operations",
    legacySessionMutation: "pm-headless-reviewer",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary:
      "The headless reviewer has no canonical session selector while its mutations remain PM-owned.",
  },
  [CANONICAL_GAP.aiSnapshots]: {
    owner: "ai-edits",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Headless AI edits still use PM execution and snapshot undo.",
  },
} as const satisfies Record<CanonicalGap, CanonicalCapability>;

/**
 * The engine an editor edits with. The canonical session is the default;
 * `"legacy"` opts out to the ProseMirror-owned engine.
 */
export type EditorSession = "canonical" | "legacy";

/**
 * Whether an editor runs the canonical session: every editor does unless its
 * host selected `"legacy"`. This is the one place the default is decided.
 * The gap argument is a typed tag counted by the source guard, not a runtime policy.
 * Retire a gap by removing its sites and ledger entry.
 */
export const usesCanonicalSession = (
  session: EditorSession | undefined,
  _gap: CanonicalGap,
): boolean => session !== "legacy";

export const canonicalRefusalMessage = (
  gap: CanonicalGap,
  message: string,
): string => `[${gap}] ${message}`;
