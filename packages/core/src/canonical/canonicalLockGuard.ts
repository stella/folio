import {
  type ContentControlLock,
  contentControlLockAllows,
  type LockedControlViolation,
  locksContent,
  locksControl,
} from "@stll/docx-core/model";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  documentStories,
  findStoryBody,
  type OpStory,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";

/** A revision wrapper a control stands inside, outermost first. */
type StandingRevision = { kind: string; id: number };

type Control = {
  lock: ContentControlLock | undefined;
  signature: string;
  /** The revisions around the control: resolving one changes what it holds. */
  standing: readonly StandingRevision[];
};

const REVISION_WRAPPERS: ReadonlySet<string> = new Set([
  "insertion",
  "deletion",
  "moveFrom",
  "moveTo",
]);
/** The wrappers that, standing newly around a control, remove it once taken. */
const PENDING_REMOVALS: ReadonlySet<string> = new Set(["deletion", "moveFrom"]);

/** `standing`, with `node` added when it is a revision wrapper. */
const standingWithin = (
  node: { type?: unknown; info?: { id?: unknown } },
  standing: readonly StandingRevision[],
): readonly StandingRevision[] =>
  typeof node.type === "string" &&
  REVISION_WRAPPERS.has(node.type) &&
  typeof node.info?.id === "number"
    ? [...standing, { kind: node.type, id: node.info.id }]
    : standing;

/**
 * Whether a control kept by a change stands as it did: under the same
 * revisions, or newly under a pending removal, which is judged once taken.
 * Resolving a revision around it changes what its content is.
 */
const standsAsBefore = (
  before: readonly StandingRevision[],
  now: readonly StandingRevision[],
): boolean => {
  const same = (left: StandingRevision, right: StandingRevision) =>
    left.kind === right.kind && left.id === right.id;
  if (!before.every((revision) => now.some((held) => same(held, revision))))
    return false;
  return now.every(
    (revision) =>
      before.some((held) => same(held, revision)) ||
      PENDING_REMOVALS.has(revision.kind),
  );
};

/** Comment marks, anchors and identities: what a lock does not hold still. */
const IGNORED_KEYS = new Set(["paraId", "textId"]);
const IGNORED_ITEMS = new Set([
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
]);
const isIgnoredItem = (item: unknown) =>
  typeof item === "object" &&
  item !== null &&
  "type" in item &&
  IGNORED_ITEMS.has(String(item.type));
const isRevisionIdentity = (key: string, value: unknown) =>
  key === "preservedAttributes" && Array.isArray(value);

/** What a control holds, as the lock judges it. */
const signatureOf = (content: unknown): string =>
  JSON.stringify(content, (key, value: unknown) => {
    if (IGNORED_KEYS.has(key)) return undefined;
    if (isRevisionIdentity(key, value))
      return (value as { name?: unknown }[]).filter(
        ({ name }) => typeof name !== "string" || !name.startsWith("rsid"),
      );
    if (Array.isArray(value))
      return value.filter((item) => !isIgnoredItem(item));
    return value;
  });

const storyKey = (story: OpStory) => {
  if (story === "main") return "main";
  return "rId" in story
    ? `${story.kind}:${story.rId}`
    : `${story.kind}:${String(story.id)}`;
};

type ControlEntry = { id: number | undefined; control: Control };

/** The content controls inside one value, in document order. */
const collectControls = (
  value: unknown,
  into: ControlEntry[],
  standing: readonly StandingRevision[] = [],
): void => {
  if (Array.isArray(value)) {
    for (const item of value) collectControls(item, into, standing);
    return;
  }
  if (typeof value !== "object" || value === null) return;
  const node = value as {
    type?: unknown;
    info?: { id?: unknown };
    properties?: { id?: number; lock?: ContentControlLock };
    content?: unknown;
  };
  if (node.type === "blockSdt" || node.type === "inlineSdt")
    into.push({
      id: node.properties?.id,
      control: {
        lock: node.properties?.lock,
        signature: signatureOf(node.content),
        standing,
      },
    });
  const within = standingWithin(node, standing);
  for (const child of Object.values(value))
    collectControls(child, into, within);
};

// An operation leaves every block it does not name the same object, and never
// changes a block in place, so a block's controls are read once. Typing then
// reads the one block it edits instead of the whole document.
const blockControls = new WeakMap<object, readonly ControlEntry[]>();
const controlsInBlock = (block: object): readonly ControlEntry[] => {
  let entries = blockControls.get(block);
  if (entries === undefined) {
    const collected: ControlEntry[] = [];
    collectControls(block, collected);
    entries = collected;
    blockControls.set(block, entries);
  }
  return entries;
};

/**
 * Every content control in the document, keyed by its story and `w:id`, or by
 * its order among the controls without one.
 */
export const controlsOf = (document: Document): Map<string, Control> => {
  const controls = new Map<string, Control>();
  for (const story of documentStories(document)) {
    const body = findStoryBody(document, story);
    if (body === undefined) continue;
    let anonymous = 0;
    for (const block of body.content) {
      for (const { id, control } of controlsInBlock(block)) {
        const key = `${storyKey(story)}#${id === undefined ? `anonymous ${String(anonymous++)}` : String(id)}`;
        controls.set(key, control);
      }
    }
  }
  return controls;
};

/**
 * The same map read from the whole document in one walk, with no reuse: the
 * reference the block reads above must equal.
 */
export const controlsOfWholeDocument = (
  document: Document,
): Map<string, Control> => {
  const controls = new Map<string, Control>();
  for (const story of documentStories(document)) {
    const body = findStoryBody(document, story);
    if (body === undefined) continue;
    let anonymous = 0;
    const visit = (
      value: unknown,
      standing: readonly StandingRevision[] = [],
    ): void => {
      if (Array.isArray(value)) {
        for (const item of value) visit(item, standing);
        return;
      }
      if (typeof value !== "object" || value === null) return;
      const node = value as {
        type?: unknown;
        info?: { id?: unknown };
        properties?: { id?: number; lock?: ContentControlLock };
        content?: unknown;
      };
      if (node.type === "blockSdt" || node.type === "inlineSdt") {
        const id = node.properties?.id;
        const key = `${storyKey(story)}#${id === undefined ? `anonymous ${String(anonymous++)}` : String(id)}`;
        controls.set(key, {
          lock: node.properties?.lock,
          signature: signatureOf(node.content),
          standing,
        });
      }
      const within = standingWithin(node, standing);
      for (const child of Object.values(value)) visit(child, within);
    };
    visit(body.content);
  }
  return controls;
};

const lockedControls = new WeakMap<Document, Map<string, Control>>();
const lockedIn = (document: Document) => {
  let locked = lockedControls.get(document);
  if (locked === undefined) {
    locked = new Map(
      [...controlsOf(document)].filter(
        ([, { lock }]) => locksContent(lock) || locksControl(lock),
      ),
    );
    lockedControls.set(document, locked);
  }
  return locked;
};

/** Only locks present in the source can hold a resolution. */
export const hasLockedControls = (document: Document): boolean =>
  lockedIn(document).size > 0;

/** What `after` does to the locked controls `locked` lists, read as it stands. */
const violationIn = (
  locked: ReadonlyMap<string, Control>,
  after: Document,
): LockedControlViolation | null => {
  const now = controlsOf(after);
  for (const [key, { lock, signature, standing }] of locked) {
    if (lock === undefined) continue;
    const kept = now.get(key);
    if (kept === undefined) {
      if (!contentControlLockAllows(lock, "removeControl"))
        return { lock, change: "removeControl" };
    } else if (
      (kept.signature !== signature ||
        !standsAsBefore(standing, kept.standing)) &&
      !contentControlLockAllows(lock, "editContent")
    )
      return { lock, change: "editContent" };
  }
  return null;
};

/** The ids of the revision records in one story. */
const revisionIdsIn = (document: Document, story: OpStory): Set<number> => {
  const ids = new Set<number>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    if (
      "info" in value &&
      typeof value.info === "object" &&
      value.info !== null &&
      "id" in value.info &&
      typeof value.info.id === "number" &&
      "author" in value.info
    )
      ids.add(value.info.id);
    for (const child of Object.values(value)) visit(child);
  };
  visit(findStoryBody(document, story)?.content);
  return ids;
};

/**
 * `after` with the revisions it records and `before` does not accepted: what
 * a tracked change does once it is taken. `undefined` when it records none,
 * or when the resolution refuses one, which leaves the change judged as it
 * stands.
 */
const acceptedNewRevisions = (
  before: Document,
  after: Document,
): Document | undefined => {
  const ops: DocumentOp[] = [];
  for (const story of documentStories(after)) {
    const earlier = revisionIdsIn(before, story);
    const added = [...revisionIdsIn(after, story)].filter(
      (id) => !earlier.has(id),
    );
    if (added.length > 0)
      ops.push({
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds: added,
        decision: REVISION_DECISIONS.ACCEPT,
      });
  }
  if (ops.length === 0) return undefined;
  const accepted = applyDocumentOps(after, ops);
  return accepted.isOk() ? accepted.value.document : undefined;
};

/**
 * What a change from `before` to `after` does to a locked content control,
 * judged on the result rather than the steps that produced it, so every path
 * into the session meets the same rule. A tracked change is judged as it
 * stands and once accepted: a pending deletion of a control's text edits it,
 * and one of the control removes it, when accepted.
 */
export const lockViolationBetween = (
  before: Document,
  after: Document,
): LockedControlViolation | null => {
  const locked = lockedIn(before);
  if (locked.size === 0 || before === after) return null;
  const standing = violationIn(locked, after);
  if (standing !== null) return standing;
  const accepted = acceptedNewRevisions(before, after);
  return accepted === undefined ? null : violationIn(locked, accepted);
};
