/** Remaining seams to retire before the canonical session becomes the only authority. */
export const CANONICAL_GAP = {
  authorityRouting: "authority-routing",
  commands: "command-descriptors",
  suggesting: "adapter-suggesting",
  comments: "comment-model-edits",
  modelEdits: "direct-model-edits",
  sectionProperties: "section-properties",
  watermark: "watermark-model-edits",
  secondaryStories: "secondary-story-routing",
  collaboration: "collaboration-session",
  dispatch: "unclassified-transactions",
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
  status: "remaining";
  summary: string;
};

export const CANONICAL_CAPABILITIES = {
  [CANONICAL_GAP.authorityRouting]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Explicit session selection and legacy authority routing.",
  },
  [CANONICAL_GAP.commands]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    status: "remaining",
    summary:
      "Commands without canonical descriptors retain PM probing and hit the dispatch boundary.",
  },
  [CANONICAL_GAP.suggesting]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Adapter suggesting controls and tracked secondary-story edits remain gated.",
  },
  [CANONICAL_GAP.comments]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Comment mutations still use direct model changes.",
  },
  [CANONICAL_GAP.modelEdits]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Direct model and pending-suggestion snapshots bypass the canonical journal.",
  },
  [CANONICAL_GAP.sectionProperties]: {
    owner: "adapters",
    kind: "routing",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Section properties select canonical operations or legacy model changes.",
  },
  [CANONICAL_GAP.watermark]: {
    owner: "adapters",
    kind: "refusal",
    adapters: ["vue"],
    status: "remaining",
    summary: "Vue watermark mutations still change the model directly.",
  },
  [CANONICAL_GAP.secondaryStories]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Secondary-story creation, views and history still select their session authority.",
  },
  [CANONICAL_GAP.collaboration]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Collaboration cannot activate a canonical session.",
  },
  [CANONICAL_GAP.dispatch]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Unclassified native and plugin mutations are refused at the projection boundary.",
  },
  [CANONICAL_GAP.save]: {
    owner: "controller",
    kind: "routing",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Legacy save and read paths reconstruct documents from PM.",
  },
  [CANONICAL_GAP.history]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "PM history remains installed for legacy sessions.",
  },
  [CANONICAL_GAP.suggestionPlugin]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "The legacy suggestion plugin produces PM transactions.",
  },
  [CANONICAL_GAP.paragraphTracker]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Dirty tracking and section bookkeeping still observe PM transactions.",
  },
  [CANONICAL_GAP.tableGeometry]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Table geometry, resize and repair still produce PM transactions.",
  },
  [CANONICAL_GAP.paragraphIdentity]: {
    owner: "prosemirror",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Legacy paragraph identity repair remains a PM append transaction.",
  },
  [CANONICAL_GAP.documentOperations]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Public document operations still execute against a held PM view.",
  },
  [CANONICAL_GAP.publicComments]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Public comment operations still use the PM executor.",
  },
  [CANONICAL_GAP.publicSuggestedMode]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Public suggested-mode receipts still use the PM executor.",
  },
  [CANONICAL_GAP.publicTableProjection]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Public table operations still depend on PM projection geometry.",
  },
  [CANONICAL_GAP.publicUnsupportedInline]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Public inline payloads still require canonical compiler coverage.",
  },
  [CANONICAL_GAP.publicHeadlessSession]: {
    owner: "document-operations",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary:
      "The headless reviewer has no canonical session selector while its mutations remain PM-owned.",
  },
  [CANONICAL_GAP.aiSnapshots]: {
    owner: "ai-edits",
    kind: "mutation-source",
    adapters: ["react", "vue"],
    status: "remaining",
    summary: "Headless AI edits still use PM execution and snapshot undo.",
  },
} as const satisfies Record<CanonicalGap, CanonicalCapability>;

/** Both adapters identify session branches through the same typed capability ledger. */
export const usesCanonicalSession = (
  session: "canonical" | undefined,
  gap: CanonicalGap,
): boolean => CANONICAL_CAPABILITIES[gap].status === "remaining" && session === "canonical";

export const canonicalRefusalMessage = (gap: CanonicalGap, message: string): string =>
  `[${gap}] ${message}`;
