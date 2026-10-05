/** Remaining seams to retire before the canonical session becomes the only authority. */
export const CANONICAL_GAP = {
  authorityRouting: "authority-routing",
  commands: "command-descriptors",
  suggesting: "adapter-suggesting",
  trackedHyperlinkResolution: "tracked-hyperlink-resolution",
  comments: "comment-model-edits",
  modelEdits: "direct-model-edits",
  sectionProperties: "section-properties",
  watermark: "watermark-model-edits",
  secondaryStories: "secondary-story-routing",
  collaboration: "collaboration-session",
  dispatch: "unclassified-transactions",
  tableActivation: "table-session-activation",
  save: "pm-save-projection",
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
} as const;

export type CanonicalGap = (typeof CANONICAL_GAP)[keyof typeof CANONICAL_GAP];

type CanonicalCapability = {
  owner: "controller" | "adapters" | "prosemirror" | "document-operations" | "ai-edits";
  kind: "routing" | "refusal" | "mutation-source";
  adapters: readonly ("react" | "vue")[];
  summary: string;
  remainingCommands?: { registry: "extension-commands" };
};

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
    summary: "Hyperlink suggestions require serializable wrapper review provenance.",
  },
  [CANONICAL_GAP.suggesting]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary: "Adapter suggesting controls and tracked secondary-story edits remain gated.",
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
  [CANONICAL_GAP.secondaryStories]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    summary: "Secondary-story creation, views and history still select their session authority.",
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
  [CANONICAL_GAP.tableActivation]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary: "Canonical session activation requires table projection and cell addressing.",
  },
  [CANONICAL_GAP.save]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    summary: "Legacy save and read paths reconstruct documents from PM.",
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
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Public document operations still execute against a held PM view.",
  },
  [CANONICAL_GAP.publicComments]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Public comment operations still use the PM executor.",
  },
  [CANONICAL_GAP.publicSuggestedMode]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Public suggested-mode receipts still use the PM executor.",
  },
  [CANONICAL_GAP.publicTableProjection]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Public table operations still depend on PM projection geometry.",
  },
  [CANONICAL_GAP.publicUnsupportedInline]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    summary: "Public inline payloads still require canonical compiler coverage.",
  },
  [CANONICAL_GAP.publicHeadlessSession]: {
    owner: "document-operations",
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
