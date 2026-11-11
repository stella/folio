import {
  type ContentControlLock,
  contentControlLockAllows,
  type LockedControlViolation,
  locksContent,
  locksControl,
} from "@stll/docx-core/model";
import {
  documentStories,
  findStoryBody,
  type OpStory,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";

type Control = { lock: ContentControlLock | undefined; signature: string };

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
const collectControls = (value: unknown, into: ControlEntry[]): void => {
  if (Array.isArray(value)) {
    for (const item of value) collectControls(item, into);
    return;
  }
  if (typeof value !== "object" || value === null) return;
  const node = value as {
    type?: unknown;
    properties?: { id?: number; lock?: ContentControlLock };
    content?: unknown;
  };
  if (node.type === "blockSdt" || node.type === "inlineSdt")
    into.push({
      id: node.properties?.id,
      control: {
        lock: node.properties?.lock,
        signature: signatureOf(node.content),
      },
    });
  for (const child of Object.values(value)) collectControls(child, into);
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
    const visit = (value: unknown): void => {
      if (Array.isArray(value)) {
        for (const item of value) visit(item);
        return;
      }
      if (typeof value !== "object" || value === null) return;
      const node = value as {
        type?: unknown;
        properties?: { id?: number; lock?: ContentControlLock };
        content?: unknown;
      };
      if (node.type === "blockSdt" || node.type === "inlineSdt") {
        const id = node.properties?.id;
        const key = `${storyKey(story)}#${id === undefined ? `anonymous ${String(anonymous++)}` : String(id)}`;
        controls.set(key, {
          lock: node.properties?.lock,
          signature: signatureOf(node.content),
        });
      }
      for (const child of Object.values(value)) visit(child);
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

/**
 * What a change from `before` to `after` does to a locked content control,
 * judged on the result rather than the steps that produced it, so every path
 * into the session meets the same rule.
 */
export const lockViolationBetween = (
  before: Document,
  after: Document,
): LockedControlViolation | null => {
  const locked = lockedIn(before);
  if (locked.size === 0 || before === after) return null;
  const now = controlsOf(after);
  for (const [key, { lock, signature }] of locked) {
    if (lock === undefined) continue;
    const kept = now.get(key);
    if (kept === undefined) {
      if (!contentControlLockAllows(lock, "removeControl"))
        return { lock, change: "removeControl" };
    } else if (
      kept.signature !== signature &&
      !contentControlLockAllows(lock, "editContent")
    )
      return { lock, change: "editContent" };
  }
  return null;
};
