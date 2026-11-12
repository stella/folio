import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  LOCAL_KINDS,
  type LocalKind,
  MAX_PEERS,
} from "./__tests__/canonicalPeerEdits";
import {
  createRoomPeers,
  emptyRoomCoverage,
  ROOM_BLOCK_SIZES,
  type RoomCoverage,
  type RoomPeer,
  settle,
} from "./__tests__/canonicalRoomPeers";
import { restoreCanonicalRoomAuthority } from "./canonicalRoomAuthority";
import { toWireJson } from "./canonicalRoomWire";
import {
  type CanonicalTabEnvironment,
  type CanonicalTabJournal,
  type CanonicalTabRoom,
  type CanonicalTabRoomError,
  openCanonicalTabRoom,
} from "./canonicalTabRoom";

setDefaultTimeout(propertyTestTimeout(180_000));

/*
 * 3 or 4 tabs of one origin share a room: each tab is a real canonical
 * session joined through the controller to a tab room. The simulated browser
 * gives every tab a BroadcastChannel endpoint and one exclusive lock; every
 * message between tabs waits in one network queue, so deliveries are delayed
 * and reordered, and one may run twice. The leading tab may be killed: its
 * messages in flight are lost, the lock passes to the next waiting tab, which
 * must rebuild the authority from the shared durable journal before it leads.
 * A killed tab comes back later as the same actor in a new tab; the last tab
 * may join late.
 *
 * Checked once the network is drained: every live tab follows the room and
 * holds the authority's document (the durable journal, replayed), saves the
 * same bytes, keeps every sequenced keystroke and every acknowledged edit,
 * and sees every other tab's selection where that tab holds it.
 */

type Command =
  | { type: "local"; peer: number; kind: LocalKind; seed: number }
  | { type: "deliver"; seed: number }
  | { type: "repeat"; seed: number }
  | { type: "kill" }
  | { type: "reload" }
  | { type: "join" };

const commandArbitrary: fc.Arbitrary<Command> = fc.oneof(
  {
    weight: 6,
    arbitrary: fc.record({
      type: fc.constant("local" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
      kind: fc.constantFrom(...LOCAL_KINDS),
      seed: fc.nat(),
    }),
  },
  {
    weight: 8,
    arbitrary: fc.record({
      type: fc.constant("deliver" as const),
      seed: fc.nat(),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("repeat" as const),
      seed: fc.nat(),
    }),
  },
  { weight: 1, arbitrary: fc.constant({ type: "kill" as const }) },
  { weight: 1, arbitrary: fc.constant({ type: "reload" as const }) },
  { weight: 1, arbitrary: fc.constant({ type: "join" as const }) },
);

type TabCoverage = RoomCoverage & { handoffs: number; reloads: number };

type Endpoint = { tab: SimTab; listeners: Set<(message: unknown) => void> };
type SimTab = { alive: boolean };
type Hop = { from: SimTab; to: SimTab; run: () => void };
type Waiter = {
  tab: SimTab;
  signal: AbortSignal;
  lead: () => Promise<void>;
  released: () => void;
};

/**
 * One origin's tabs: BroadcastChannel endpoints whose messages wait in
 * `hops`, one exclusive lock granted in request order, and the durable
 * journal they share. A killed tab sends, receives and stores nothing more,
 * and its lock passes on.
 */
const createBrowser = () => {
  const hops: Hop[] = [];
  const endpoints = new Set<Endpoint>();
  const waiters: Waiter[] = [];
  let holder: Waiter | undefined;
  const stored: unknown[] = [];
  let term = 0;

  const grantNext = (): void => {
    if (holder !== undefined) return;
    const index = waiters.findIndex(
      ({ tab, signal }) => tab.alive && !signal.aborted,
    );
    if (index < 0) return;
    const [next] = waiters.splice(index, 1);
    if (next === undefined) return;
    holder = next;
    const hold = async (): Promise<void> => {
      await next.lead();
      if (holder !== next) return;
      holder = undefined;
      next.released();
      grantNext();
    };
    void hold();
  };

  const environment = (tab: SimTab): CanonicalTabEnvironment => ({
    openChannel: () => {
      const endpoint: Endpoint = { tab, listeners: new Set() };
      endpoints.add(endpoint);
      return {
        post: (message) => {
          if (!tab.alive) return;
          for (const other of endpoints) {
            if (other === endpoint || !other.tab.alive) continue;
            hops.push({
              from: tab,
              to: other.tab,
              run: () => {
                if (!other.tab.alive || !endpoints.has(other)) return;
                for (const listener of other.listeners) listener(message);
              },
            });
          }
        },
        onMessage: (listener) => {
          endpoint.listeners.add(listener);
          return () => endpoint.listeners.delete(listener);
        },
        close: () => {
          endpoints.delete(endpoint);
        },
      };
    },
    requestLock: (_name, signal, lead) =>
      new Promise((released) => {
        waiters.push({ tab, signal, lead, released: () => released() });
        queueMicrotask(grantNext);
      }),
  });

  const journal = (tab: SimTab): CanonicalTabJournal => ({
    load: async () => ({ records: [...stored], term }),
    storeTerm: async (next) => {
      if (tab.alive) term = next;
    },
    append: async (records) => {
      if (tab.alive)
        stored.push(...records.map((record) => toWireJson(record)));
    },
  });

  const kill = (tab: SimTab): void => {
    tab.alive = false;
    for (let index = hops.length - 1; index >= 0; index -= 1) {
      const hop = hops[index];
      if (hop !== undefined && (hop.from === tab || hop.to === tab))
        hops.splice(index, 1);
    }
    if (holder?.tab === tab) {
      holder = undefined;
      grantNext();
    }
  };

  return {
    hops,
    environment,
    journal,
    kill,
    get leader() {
      return holder?.tab;
    },
    get stored(): readonly unknown[] {
      return stored;
    },
  };
};

/**
 * A deliberate defect in the journal every tab of one run reads: made once
 * per run, so the tabs share its state.
 */
type JournalFault = () => (journal: CanonicalTabJournal) => CanonicalTabJournal;

type Slot = {
  tab: SimTab;
  room: CanonicalTabRoom;
  incarnation: number;
  peer?: RoomPeer;
};

const simulate = async ({
  peersCount,
  lateJoin,
  commands,
  journalFault,
  coverage,
}: {
  peersCount: number;
  lateJoin: boolean;
  commands: readonly Command[];
  journalFault: JournalFault | undefined;
  coverage: TabCoverage;
}): Promise<void> => {
  const browser = createBrowser();
  const room = await createRoomPeers(coverage);
  const faulty = journalFault?.();
  const slots: (Slot | undefined)[] = [];
  const errors: CanonicalTabRoomError[] = [];
  const joining = new Set<Promise<void>>();

  const open = (index: number, incarnation: number): void => {
    const tab: SimTab = { alive: true };
    const journal = browser.journal(tab);
    const tabRoom = openCanonicalTabRoom({
      room: "property",
      snapshot: room.snapshot,
      document: room.base,
      journal: faulty === undefined ? journal : faulty(journal),
      idBlockSizes: ROOM_BLOCK_SIZES,
      environment: browser.environment(tab),
      presenceRefreshMs: 0,
      onError: (error) => errors.push(error),
    });
    const slot: Slot = { tab, room: tabRoom, incarnation };
    slots[index] = slot;
    const joinSlot = async (): Promise<void> => {
      const peer = await room.join({ index, incarnation, transport: tabRoom });
      if (tab.alive) slot.peer = peer;
    };
    const joined = joinSlot().finally(() => joining.delete(joined));
    joining.add(joined);
  };
  const live = (): RoomPeer[] =>
    slots.flatMap((slot) =>
      slot?.tab.alive === true && slot.peer !== undefined ? [slot.peer] : [],
    );

  const startCount = lateJoin ? peersCount - 1 : peersCount;
  for (let index = 0; index < startCount; index += 1) open(index, 0);
  await settle();

  for (const command of commands) {
    switch (command.type) {
      case "local": {
        const peer = slots.at(command.peer % peersCount)?.peer;
        if (peer === undefined || !slots[peer.index]?.tab.alive) break;
        room.edit(peer, live(), command.kind, command.seed);
        break;
      }
      case "deliver":
      case "repeat": {
        const { hops } = browser;
        if (hops.length === 0) break;
        const index = command.seed % hops.length;
        const hop =
          command.type === "deliver" ? hops.splice(index, 1)[0] : hops[index];
        hop?.run();
        break;
      }
      case "kill": {
        const leader = browser.leader;
        const slot = slots.find((candidate) => candidate?.tab === leader);
        // Keep at least one tab joined, so the room goes on.
        if (slot?.peer === undefined || live().length < 2) break;
        coverage.handoffs += 1;
        // Another tab's submission the dying leader leaves unanswered.
        if (
          live().some(
            (peer) =>
              peer !== slot.peer && peer.collaboration.pending.length > 0,
          )
        )
          coverage.resubmitted += 1;
        browser.kill(slot.tab);
        break;
      }
      case "reload": {
        const index = slots.findIndex((slot) => slot?.tab.alive === false);
        const slot = slots[index];
        if (slot === undefined) break;
        coverage.reloads += 1;
        open(index, slot.incarnation + 1);
        break;
      }
      case "join":
        if (slots.length < peersCount) {
          coverage.lateJoins += 1;
          open(peersCount - 1, 0);
        }
        break;
      default: {
        const exhaustive: never = command;
        return exhaustive;
      }
    }
    await settle();
  }
  if (slots.length < peersCount) open(peersCount - 1, 0);
  for (const [index, slot] of slots.entries())
    if (slot !== undefined && !slot.tab.alive)
      open(index, slot.incarnation + 1);
  await settle();
  for (let steps = 0; browser.hops.length > 0 || joining.size > 0; steps += 1) {
    expect(steps).toBeLessThan(40_000);
    browser.hops.shift()?.run();
    await settle();
  }

  expect(errors).toEqual([]);
  const authority = restoreCanonicalRoomAuthority({
    document: room.base,
    records: browser.stored,
    idBlockSizes: ROOM_BLOCK_SIZES,
  }).unwrap();
  const peers = live();
  expect(peers).toHaveLength(peersCount);
  await room.expectConverged({
    peers,
    peersCount,
    document: authority.document,
    journal: authority.journal,
    contiguous: () => true,
  });
  for (const slot of slots) slot?.room.close();
};

const law = (journalFault: JournalFault | undefined, coverage: TabCoverage) =>
  fc.asyncProperty(
    fc.integer({ min: 3, max: MAX_PEERS }),
    fc.boolean(),
    fc.array(commandArbitrary, { minLength: 10, maxLength: 40 }),
    async (peersCount, lateJoin, commands) =>
      simulate({ peersCount, lateJoin, commands, journalFault, coverage }),
  );

const deliveries = (count: number, seed = 0): Command[] =>
  Array.from({ length: count }, () => ({ type: "deliver" as const, seed }));
/** Enough first-in-first-out deliveries for every waiting tab to join. */
const JOINED = deliveries(40);

/**
 * Peer 1's batch reaches the leading tab (peer 0's), which sequences and
 * stores it, then dies with the broadcast and the answer still in flight.
 * The next leader must replay the journal: peer 1's resubmission is
 * answered from it, and peer 2's edit lands after it.
 */
const HANDOFF_AFTER_SEQUENCED: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 1, kind: "type", seed: 0 },
  ...deliveries(1),
  { type: "kill" },
  { type: "local", peer: 2, kind: "prefix", seed: 1 },
];

/** Two handoffs; each dead tab reloads as the same actor in a new tab. */
const TWO_HANDOFFS: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 1, kind: "type", seed: 0 },
  { type: "local", peer: 2, kind: "splitEnd", seed: 0 },
  ...deliveries(3),
  { type: "kill" },
  ...JOINED,
  { type: "reload" },
  ...JOINED,
  { type: "local", peer: 2, kind: "type", seed: 0 },
  ...deliveries(1),
  { type: "kill" },
  { type: "local", peer: 0, kind: "type", seed: 0 },
];

/** A late tab joins after a handoff and replays the journal. */
const LATE_JOIN_AFTER_HANDOFF: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 1, kind: "type", seed: 0 },
  ...deliveries(1),
  { type: "kill" },
  { type: "join" },
  ...JOINED,
  { type: "local", peer: 3, kind: "type", seed: 0 },
];

/** Peer 1 types, peer 2's edit reaches it, peer 1 undoes. */
const UNDO_AFTER_REMOTE: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 1, kind: "type", seed: 0 },
  { type: "local", peer: 2, kind: "type", seed: 0 },
  ...JOINED,
  { type: "local", peer: 1, kind: "undo", seed: 0 },
];

const EXAMPLES: [number, boolean, Command[]][] = [
  [3, false, [...HANDOFF_AFTER_SEQUENCED]],
  [3, false, [...TWO_HANDOFFS]],
  [4, true, [...LATE_JOIN_AFTER_HANDOFF]],
  [3, false, [...UNDO_AFTER_REMOTE]],
];

const emptyCoverage = (): TabCoverage => ({
  ...emptyRoomCoverage(),
  handoffs: 0,
  reloads: 0,
});

/** The defect: a new leader starts from an empty journal, not the stored one. */
const leaderSkipsReplay: JournalFault = () => {
  let loads = 0;
  return (journal) => ({
    ...journal,
    load: async () => {
      const loaded = await journal.load();
      loads += 1;
      return loads === 1 ? loaded : { records: [], term: loaded.term };
    },
  });
};

test("tabs sharing a room converge on the authority's document across leader handoffs", async () => {
  await assertProperty(law(undefined, emptyCoverage()), {
    numRuns: 12,
    examples: EXAMPLES,
    id: "canonical tab room convergence",
  });
  // On a fixed sample: remote edits arrive, undo follows them, the leading
  // tab dies with work in flight, dead tabs reload and a late tab joins.
  const coverage = emptyCoverage();
  await checkCoverageSample(law(undefined, coverage), 8, {
    map: ([peersCount, lateJoin, commands], index) => {
      const example = EXAMPLES[index % EXAMPLES.length];
      return example === undefined
        ? [peersCount, lateJoin, commands]
        : [example[0], example[1], [...example[2], ...commands]];
    },
  });
  expect(coverage.remoteApplied).toBeGreaterThan(0);
  expect(coverage.undoAfterRemote).toBeGreaterThan(0);
  expect(coverage.handoffs).toBeGreaterThan(1);
  expect(coverage.resubmitted).toBeGreaterThan(0);
  expect(coverage.reloads).toBeGreaterThan(0);
  expect(coverage.lateJoins).toBeGreaterThan(0);
});

test("the property catches a new leader that starts before replaying the journal", async () => {
  const result = await fc.check(
    law(leaderSkipsReplay, emptyCoverage()),
    propertyConfig({ numRuns: 30, endOnFailure: true, examples: EXAMPLES }),
  );
  expect(result.failed).toBe(true);
});

test("a leader with no journal does not lead", async () => {
  const browser = createBrowser();
  const room = await createRoomPeers(emptyRoomCoverage());
  const tab: SimTab = { alive: true };
  const errors: CanonicalTabRoomError[] = [];
  const opened = openCanonicalTabRoom({
    room: "unreadable",
    snapshot: room.snapshot,
    document: room.base,
    journal: {
      ...browser.journal(tab),
      load: () => Promise.reject(new Error("unavailable")),
    },
    environment: browser.environment(tab),
    presenceRefreshMs: 0,
    onError: (error) => errors.push(error),
  });
  for (let round = 0; round < 4; round += 1) await settle();
  expect(opened.leading).toBe(false);
  expect(opened.connected).toBe(false);
  expect(errors.map(({ reason }) => reason)).toEqual(["journalUnavailable"]);
  expect(browser.leader).toBeUndefined();
  opened.close();
});
