/** Remaining seams to retire before the canonical session becomes the only authority. */
export const CANONICAL_GAP = {
  trackedTableStructure: "tracked-table-structure",
  storyContentProjection: "story-content-projection",
  trackedChangeRecords: "tracked-change-records",
} as const;

export type CanonicalGap = (typeof CANONICAL_GAP)[keyof typeof CANONICAL_GAP];

/**
 * Rules a document states for itself. A refusal under one is the correct
 * outcome and stays after the cutover; it is never a gap to retire.
 */
export const CANONICAL_RULE: {
  readonly contentControlLock: "contentControlLock";
  readonly linkSuggestionShape: "linkSuggestionShape";
  readonly revisionContentShape: "revisionContentShape";
  readonly tableSnapshotPresence: "tableSnapshotPresence";
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
  /**
   * OOXML has no revision record for the content a suggestion would track:
   * a `w:ins`, `w:del`, `w:moveFrom` or `w:moveTo` holds run-level content
   * only (CT_RunTrackChange), and a block-level `w:sdt` takes no revision of
   * its own (CT_SdtBlock records none), so a suggested paste or move of a
   * block content control is refused. The same change made directly applies.
   */
  revisionContentShape: "revisionContentShape",
  /** Review snapshots cannot store JavaScript's distinction between omitted and own undefined fields. */
  tableSnapshotPresence: "tableSnapshotPresence",
} as const;

export type CanonicalRule =
  (typeof CANONICAL_RULE)[keyof typeof CANONICAL_RULE];

/** The editor message, in the folio catalog, a host shows for a refusal under each rule. */
export const CANONICAL_RULE_MESSAGE_KEYS: {
  readonly contentControlLock: "editor.contentLocked";
  readonly linkSuggestionShape: "editor.linkSuggestionUnavailable";
  readonly revisionContentShape: "editor.suggestionUnavailable";
  readonly tableSnapshotPresence: "editor.changeNotApplied";
} = {
  contentControlLock: "editor.contentLocked",
  linkSuggestionShape: "editor.linkSuggestionUnavailable",
  revisionContentShape: "editor.suggestionUnavailable",
  tableSnapshotPresence: "editor.changeNotApplied",
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
  readonly stale: "stale";
  readonly hostTransaction: "hostTransaction";
  readonly internal: "internal";
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
  /** The projection, version or lease the change was prepared against moved on. */
  stale: "stale",
  /**
   * A plugin or raw dispatch changed the document outside document
   * operations; hosts express edits through `applyDocumentOperations`.
   */
  hostTransaction: "hostTransaction",
  /**
   * The session failed to project or stage the change: a defect to report,
   * not a missing capability.
   */
  internal: "internal",
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
  owner: "document-operations" | "controller" | "adapters" | "ai-edits";
};

export const CANONICAL_CAPABILITIES: {
  readonly "tracked-table-structure": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "Suggested table changes refuse markup the tracked table records cannot express: a row insertion or deletion in a table holding captured markup, wrapper stacks or row content controls; a column or cell edit in a table whose rows keep captured markup between their cells; a table inserted or deleted with captured markup, bookmarks, wrapper stacks or property revisions; a column deletion through a cell spanning other columns.";
  };
  readonly "story-content-projection": {
    readonly owner: "controller";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "A paste of content whose package parts the model does not carry refuses where those parts cannot be read: an opaque drawing, embedded object or kept block markup from another package whose copy did not read its source package's parts (a clipboard from outside the editor), or that sits in a copied note or comment.";
  };
  readonly "tracked-change-records": {
    readonly owner: "document-operations";
    readonly kind: "refusal";
    readonly adapters: readonly ["react", "vue"];
    readonly summary: "A suggested change whose content the tracked-change records cannot express yet is refused: a section-ending mark resolved across a standing range boundary, a change inside a simple field holding content the field's model cannot hold, a custom XML element or kept block markup pasted or moved, a copied paragraph mark or review carrier a tracked paste cannot attach, copy-origin records nested past their limit, a tracked deletion without joins of a record holding a comment anchor, an insertion after a pending paragraph mark without its former snapshot, and the removal of a block content control's only paragraph.";
  };
} = {
  [CANONICAL_GAP.trackedTableStructure]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "Suggested table changes refuse markup the tracked table records cannot express: a row insertion or deletion in a table holding captured markup, wrapper stacks or row content controls; a column or cell edit in a table whose rows keep captured markup between their cells; a table inserted or deleted with captured markup, bookmarks, wrapper stacks or property revisions; a column deletion through a cell spanning other columns.",
  },
  [CANONICAL_GAP.storyContentProjection]: {
    owner: "controller",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "A paste of content whose package parts the model does not carry refuses where those parts cannot be read: an opaque drawing, embedded object or kept block markup from another package whose copy did not read its source package's parts (a clipboard from outside the editor), or that sits in a copied note or comment.",
  },
  [CANONICAL_GAP.trackedChangeRecords]: {
    owner: "document-operations",
    kind: "refusal",
    adapters: ["react", "vue"],
    summary:
      "A suggested change whose content the tracked-change records cannot express yet is refused: a section-ending mark resolved across a standing range boundary, a change inside a simple field holding content the field's model cannot hold, a custom XML element or kept block markup pasted or moved, a copied paragraph mark or review carrier a tracked paste cannot attach, copy-origin records nested past their limit, a tracked deletion without joins of a record holding a comment anchor, an insertion after a pending paragraph mark without its former snapshot, and the removal of a block content control's only paragraph.",
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
