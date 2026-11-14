import * as v from "valibot";

import {
  OP_STORIES,
  type SequencedBatch,
  type TextPosition,
} from "@stll/docx-core/ops";

import { remapSelection } from "../canonical/canonicalRemoteHistory";

/*
 * Presence: where each peer's selection is, in canonical positions
 * (story, paragraph id, offset), never in an editor's own coordinates. Each
 * peer publishes its selection with the revision it had applied; a viewer
 * moves every other peer's selection over each sequenced batch it applies
 * after that revision, with the transform the session moves its own
 * selection with, and shows it through its own projection.
 */

export type CanonicalPresenceUser = { name: string; color: string };

export type CanonicalPresenceSelection = {
  anchor: TextPosition;
  head: TextPosition;
};

/** One peer's presence as it travels between peers. */
export type CanonicalPresenceState = {
  actor: string;
  /**
   * Orders one sender's states: `instance` is fresh each time the actor
   * joins, `clock` grows within it. A viewer keeps the latest.
   */
  instance: string;
  clock: number;
  user: CanonicalPresenceUser;
  /**
   * The revision the sender had applied when it took `selection`. The
   * positions also count the sender's own edits not sequenced yet.
   */
  rev: number;
  selection: CanonicalPresenceSelection | null;
};

/**
 * How presence travels: an awareness channel beside a room transport. It
 * carries the latest state of each peer and nothing else; the controller
 * validates and orders what arrives.
 */
export type CanonicalPresenceChannel = {
  /** Publish this peer's state; null withdraws it. */
  publish(state: CanonicalPresenceState | null): void;
  /**
   * Called with every other peer's state as published, unvalidated, keyed by
   * the channel's own peer id; null when that peer left. The peers known
   * when subscribing are reported after `subscribe` returns. Returns the
   * unsubscribe.
   */
  subscribe(listener: (peer: string, state: unknown) => void): () => void;
};

/** Another peer's presence, mapped into this session's document. */
export type CanonicalRemotePresence = {
  actor: string;
  user: CanonicalPresenceUser;
  /** Undefined when an edit since the peer published removed an end of it. */
  selection: CanonicalPresenceSelection | undefined;
};

const COUNT = v.pipe(v.number(), v.safeInteger(), v.minValue(0));
const STORY = v.union([
  v.literal(OP_STORIES.MAIN),
  v.strictObject({
    kind: v.picklist(["header", "footer"]),
    rId: v.string(),
  }),
  v.strictObject({
    kind: v.picklist(["footnote", "endnote"]),
    id: COUNT,
  }),
]);
const POSITION = v.strictObject({
  story: STORY,
  blockId: v.pipe(v.string(), v.minLength(1)),
  offset: COUNT,
  zeroWidthBefore: v.exactOptional(COUNT),
});
const MAX_LABEL = 256;
const PRESENCE = v.strictObject({
  actor: v.pipe(v.string(), v.minLength(1)),
  instance: v.pipe(v.string(), v.minLength(1)),
  clock: COUNT,
  user: v.strictObject({
    name: v.pipe(v.string(), v.maxLength(MAX_LABEL)),
    color: v.pipe(v.string(), v.maxLength(MAX_LABEL)),
  }),
  rev: COUNT,
  selection: v.nullable(v.strictObject({ anchor: POSITION, head: POSITION })),
});

/** A peer's published presence, or undefined when it is not one. */
export const parseCanonicalPresence = (
  value: unknown,
): CanonicalPresenceState | undefined => {
  const parsed = v.safeParse(PRESENCE, value);
  return parsed.success ? parsed.output : undefined;
};

/** A selection moved over one sequenced batch, as the session moves its own. */
const remap = (
  selection: CanonicalPresenceSelection,
  over: SequencedBatch,
): CanonicalPresenceSelection | undefined => {
  const moved = remapSelection({ type: "text", ...selection }, over);
  return moved === undefined
    ? undefined
    : { anchor: moved.anchor, head: moved.head };
};

/**
 * Batches the sender had not applied when it published: after its revision,
 * and not its own (those are its pending edits, already counted).
 */
const unseenBy = (state: CanonicalPresenceState, batch: SequencedBatch) =>
  batch.revision > state.rev && batch.actor !== state.actor;

type Tracked = {
  peer: string;
  state: CanonicalPresenceState;
  selection: CanonicalPresenceSelection | undefined;
};

/** How many applied batches a viewer keeps to move a state that arrives late. */
const RECENT_BATCHES = 256;

/**
 * Every other peer's latest presence, moved over the batches this session
 * applies. Owned by the collaboration controller.
 * @internal
 */
export const createPresenceTracker = (self: string) => {
  const byActor = new Map<string, Tracked>();
  const recent: SequencedBatch[] = [];
  let headRev = 0;

  const moveOverRecent = (
    state: CanonicalPresenceState,
  ): CanonicalPresenceSelection | undefined => {
    if (state.selection === null) return undefined;
    if (state.rev >= headRev) return state.selection;
    // The batches between the sender's revision and ours are gone.
    const oldest = recent.at(0)?.revision ?? headRev + 1;
    if (oldest > state.rev + 1) return undefined;
    let selection: CanonicalPresenceSelection | undefined = state.selection;
    for (const batch of recent) {
      if (selection === undefined) return undefined;
      if (unseenBy(state, batch)) selection = remap(selection, batch);
    }
    return selection;
  };

  return {
    /** A state the channel delivered; true when what a host shows changed. */
    receive(peer: string, value: unknown): boolean {
      if (value === null || value === undefined) {
        let changed = false;
        for (const [actor, tracked] of byActor)
          if (tracked.peer === peer) {
            byActor.delete(actor);
            changed = true;
          }
        return changed;
      }
      const state = parseCanonicalPresence(value);
      if (state === undefined || state.actor === self) return false;
      const known = byActor.get(state.actor);
      if (
        known !== undefined &&
        known.state.instance === state.instance &&
        known.state.clock >= state.clock
      )
        return false;
      byActor.set(state.actor, {
        peer,
        state,
        selection: moveOverRecent(state),
      });
      return true;
    },
    /** A sequenced batch this session applied, in revision order. */
    applied(batch: SequencedBatch): boolean {
      headRev = batch.revision;
      recent.push(batch);
      if (recent.length > RECENT_BATCHES) recent.shift();
      let changed = false;
      for (const tracked of byActor.values()) {
        if (tracked.selection === undefined || !unseenBy(tracked.state, batch))
          continue;
        tracked.selection = remap(tracked.selection, batch);
        changed = true;
      }
      return changed;
    },
    peers(): CanonicalRemotePresence[] {
      return [...byActor.values()].map(({ state, selection }) => ({
        actor: state.actor,
        user: state.user,
        selection,
      }));
    },
  };
};
