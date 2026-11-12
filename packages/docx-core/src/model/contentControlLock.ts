import type { SdtProperties } from "./content";
/**
 * What a content control's `w:lock` (ECMA-376 Part 1, §17.5.2.23) permits.
 *
 * - `contentLocked` and `sdtContentLocked` lock the content: no text,
 *   formatting or block inside the control may change.
 * - `sdtLocked` and `sdtContentLocked` lock the control: it may not be
 *   removed, neither on its own nor with the text or paragraph that holds it.
 *
 * Every editing path asks this module, in its own coordinates: a position
 * scale where the control spans `from`..`to` and its content
 * `contentFrom`..`contentTo`, and the stretch a change touches on the same
 * scale. A refused change applies nothing.
 */
import type { WritableRecord } from "./deepReadonly";

export type ContentControlLock = NonNullable<SdtProperties["lock"]>;

/** Whether `lock` forbids changing what the control holds. */
export const locksContent = (
  lock: ContentControlLock | null | undefined,
): lock is "contentLocked" | "sdtContentLocked" =>
  lock === "contentLocked" || lock === "sdtContentLocked";

/** Whether `lock` forbids removing the control. */
export const locksControl = (
  lock: ContentControlLock | null | undefined,
): lock is "sdtLocked" | "sdtContentLocked" =>
  lock === "sdtLocked" || lock === "sdtContentLocked";

/** What a change does to a control: changes what it holds, or removes it. */
export type ContentControlChange = "editContent" | "removeControl";

/**
 * Whether `lock` allows `change`: the one decision every editing path makes.
 * No lock, and `unlocked`, allow both.
 */
export const contentControlLockAllows = (
  lock: ContentControlLock | null | undefined,
  change: ContentControlChange,
): boolean => {
  switch (change) {
    case "editContent":
      return !locksContent(lock);
    case "removeControl":
      return !locksControl(lock);
    default: {
      const unhandled: never = change;
      throw new TypeError(
        `Unhandled content control change ${String(unhandled)}`,
      );
    }
  }
};

/** A content control, on the caller's position scale. */
export type LockedControlSpan = WritableRecord & {
  lock: ContentControlLock;
  /** Before the control. */
  from: number;
  /** After the control. */
  to: number;
  /** Where its content starts. */
  contentFrom: number;
  /** Where its content ends. */
  contentTo: number;
};

/**
 * The stretch a change touches. `remove` removes every control it wholly
 * covers, with its content, and changes what it partly covers; `edit`
 * changes the text or properties of what it reaches and keeps every control
 * standing (a text edit across a whole control rewrites its content). An
 * insertion is an `edit` whose `from` equals `to`.
 */
export type LockedControlTouch = WritableRecord & {
  from: number;
  to: number;
  effect: "remove" | "edit";
};

/** Why a lock refuses a change. */
export type LockedControlViolation = WritableRecord & {
  lock: ContentControlLock;
  change: ContentControlChange;
};

/** The skip or refusal reason every editing path reports for a lock. */
export const CONTENT_CONTROL_LOCK = "contentControlLock" as const;

/** What `control`'s lock says about `touch`: `null` when it allows it. */
export const lockedControlViolation = (
  control: LockedControlSpan,
  touch: LockedControlTouch,
): LockedControlViolation | null => {
  const { lock } = control;
  const removes =
    touch.effect === "remove" &&
    touch.from <= control.from &&
    touch.to >= control.to;
  const reachesContent =
    touch.from === touch.to
      ? touch.from >= control.contentFrom && touch.from <= control.contentTo
      : touch.from < control.contentTo && touch.to > control.contentFrom;
  let change: ContentControlChange | null = null;
  if (removes) change = "removeControl";
  else if (reachesContent) change = "editContent";
  return change === null || contentControlLockAllows(lock, change)
    ? null
    : { lock, change };
};

/** The first violation any of `controls` reports for any of `touches`. */
export const firstLockedControlViolation = (
  controls: Iterable<LockedControlSpan>,
  touches: readonly LockedControlTouch[],
): LockedControlViolation | null => {
  for (const control of controls)
    for (const touch of touches) {
      const violation = lockedControlViolation(control, touch);
      if (violation !== null) return violation;
    }
  return null;
};

/** A sentence naming what the lock refused. */
export const lockedControlMessage = ({
  lock,
  change,
}: LockedControlViolation): string =>
  change === "editContent"
    ? `the change reaches inside a content control whose w:lock is ${lock}.`
    : `the change removes a content control whose w:lock is ${lock}.`;
