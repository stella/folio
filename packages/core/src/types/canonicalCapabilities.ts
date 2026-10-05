/** Remaining seams to retire before the canonical session becomes the only authority. */
export const CANONICAL_GAP = {
  authorityRouting: "authority-routing",
  commands: "command-descriptors",
  trackedHyperlinkResolution: "tracked-hyperlink-resolution",
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
  publicSuggestedMode: "publicOps.suggestedMode",
  publicTableProjection: "publicOps.tableProjection",
  publicUnsupportedInline: "publicOps.unsupportedInline",
  publicHeadlessSession: "publicOps.headlessSession",
  publicSecondaryStories: "publicOps.secondaryStories",
} as const;

export type CanonicalGap = (typeof CANONICAL_GAP)[keyof typeof CANONICAL_GAP];

type CanonicalCapability = {
  kind: "routing" | "refusal" | "mutation-source";
  adapters: readonly ("react" | "vue")[];
  summary: string;
  remainingCommands?: { registry: "extension-commands" };
} & (
  | {
      owner: "document-operations";
      defaultSessionMutation: "pm-public-operations" | "pm-headless-reviewer" | "canonical";
    }
  | { owner: "controller" | "adapters" | "prosemirror" | "ai-edits" }
);

export const CANONICAL_CAPABILITIES = {
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
    summary: "Hyperlink and TOC suggestions require serializable wrapper review provenance.",
  },
  [CANONICAL_GAP.comments]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary: "Comment mutations still use direct model changes.",
  },
  [CANONICAL_GAP.modelEdits]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary: "Direct model and pending-suggestion snapshots bypass the canonical journal.",
  },
  [CANONICAL_GAP.sectionProperties]: {
    owner: "adapters",
    kind: "routing",
    adapters: ["react", "vue"],
    summary: "Section properties select canonical operations or legacy model changes.",
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
    summary: "Unclassified native and plugin mutations are refused at the projection boundary.",
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
    summary: "Dirty tracking and section bookkeeping still observe PM transactions.",
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
    summary: "Legacy paragraph identity repair remains a PM append transaction.",
  },
  [CANONICAL_GAP.documentOperations]: {
    owner: "document-operations",
    defaultSessionMutation: "pm-public-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Public document operations still execute against a held PM view.",
  },
  [CANONICAL_GAP.publicComments]: {
    owner: "document-operations",
    defaultSessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Public comment operations refuse in canonical sessions; legacy sessions retain PM comments.",
  },
  [CANONICAL_GAP.publicSuggestedMode]: {
    owner: "document-operations",
    defaultSessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Public suggested mode refuses in canonical sessions until pending suggestions use the journal.",
  },
  [CANONICAL_GAP.publicTableProjection]: {
    owner: "document-operations",
    defaultSessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Public table operations refuse while canonical projections require paragraph stories.",
  },
  [CANONICAL_GAP.publicUnsupportedInline]: {
    owner: "document-operations",
    defaultSessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary: "Unsupported public payloads refuse with a typed compiler capability gap.",
  },
  [CANONICAL_GAP.publicSecondaryStories]: {
    owner: "document-operations",
    defaultSessionMutation: "pm-public-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary: "Public secondary-story batches refuse until canonical story routing is available.",
  },
  [CANONICAL_GAP.publicHeadlessSession]: {
    owner: "document-operations",
    defaultSessionMutation: "pm-headless-reviewer",
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
 * The gap argument is a typed tag counted by the source guard, not a runtime policy.
 * Retire a gap by removing its sites and ledger entry.
 */
export const usesCanonicalSession = (
  session: "canonical" | undefined,
  _gap: CanonicalGap,
): boolean => session === "canonical";

export const canonicalRefusalMessage = (gap: CanonicalGap, message: string): string =>
  `[${gap}] ${message}`;
