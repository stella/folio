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

/**
 * Every content control in the document, keyed by its story and `w:id`, or by
 * its order among the controls without one.
 */
const controlsOf = (document: Document): Map<string, Control> => {
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
