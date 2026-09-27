import { describe, expect, test } from "bun:test";

import { LAYOUT_MEASURE, type LayoutRunOptions } from "./layoutRunOptions";
import { createLayoutScheduler, type SchedulerClock } from "./layoutScheduler";

type TestState = { id: number };

const makeFakeClock = () => {
  let time = 0;
  let nextId = 1;
  const timers = new Map<number, () => void>();
  const frames = new Map<number, () => void>();

  const clock: SchedulerClock = {
    now: () => time,
    setTimer: (cb) => {
      const id = nextId++;
      timers.set(id, cb);
      return id;
    },
    clearTimer: (id) => {
      timers.delete(id);
    },
    requestFrame: (cb) => {
      const id = nextId++;
      frames.set(id, cb);
      return id;
    },
    cancelFrame: (id) => {
      frames.delete(id);
    },
  };

  const fireAll = (map: Map<number, () => void>): void => {
    const callbacks = [...map.values()];
    map.clear();
    for (const run of callbacks) {
      run();
    }
  };

  return {
    clock,
    advance: (ms: number) => {
      time += ms;
    },
    fireTimers: () => fireAll(timers),
    fireFrames: () => fireAll(frames),
    pendingTimers: () => timers.size,
    pendingFrames: () => frames.size,
  };
};

type SchedulerRigOptions = {
  leadingFrame?: boolean;
};

/** A scheduler over a mutable "view" state, recording every run. */
const makeRig = ({ leadingFrame = false }: SchedulerRigOptions = {}) => {
  const fake = makeFakeClock();
  const view: { state: TestState | null } = { state: { id: 0 } };
  const runs: { state: TestState; options: LayoutRunOptions }[] = [];
  const scheduler = createLayoutScheduler({
    readState: () => view.state,
    runLayout: (state, options) => {
      runs.push({ state, options });
    },
    debounceMs: 32,
    leadingFrame,
    maxDelayMs: 96,
    clock: fake.clock,
  });
  const edit = (id: number): void => {
    view.state = { id };
    scheduler.schedule();
  };
  return { fake, view, runs, scheduler, edit, ids: () => runs.map((run) => run.state.id) };
};

describe("layoutScheduler", () => {
  test("coalesces a burst into one layout run of the state current at the frame", () => {
    const rig = makeRig();

    rig.edit(1);
    rig.edit(2);
    rig.edit(3);

    // One armed debounce timer, nothing run yet.
    expect(rig.fake.pendingTimers()).toBe(1);
    expect(rig.runs).toHaveLength(0);

    rig.fake.fireTimers(); // debounce elapses -> schedules a frame
    rig.fake.fireFrames(); // frame runs layout once

    expect(rig.ids()).toEqual([3]);
  });

  test("runs an incremental transaction pass", () => {
    const rig = makeRig();

    rig.edit(1);
    rig.fake.fireTimers();
    rig.fake.fireFrames();

    expect(rig.runs.map((run) => run.options)).toEqual([
      { reason: "transaction", measure: LAYOUT_MEASURE.incremental },
    ]);
  });

  // #1142: an edit schedules a pass, the host replaces the state wholesale
  // (ref.loadDocument), and the frame only comes afterwards (a hidden page).
  test("lays out the state that replaced the edited one, never the edited one", () => {
    const rig = makeRig({ leadingFrame: true });

    rig.edit(1);
    rig.view.state = { id: 2 };
    rig.fake.advance(16_000);
    rig.fake.fireTimers();
    rig.fake.fireFrames();

    expect(rig.ids()).toEqual([2]);
  });

  test("skips the pass when there is no editor state", () => {
    const rig = makeRig({ leadingFrame: true });

    rig.edit(1);
    rig.view.state = null;
    rig.fake.fireFrames();

    expect(rig.runs).toHaveLength(0);
  });

  test("leading-frame mode paints the first edit immediately and debounces the rest", () => {
    const rig = makeRig({ leadingFrame: true });

    rig.edit(1);
    rig.edit(2);

    expect(rig.fake.pendingTimers()).toBe(0);
    expect(rig.fake.pendingFrames()).toBe(1);
    expect(rig.runs).toHaveLength(0);

    rig.fake.fireFrames();

    expect(rig.ids()).toEqual([2]);

    rig.edit(3);
    rig.edit(4);

    expect(rig.fake.pendingTimers()).toBe(1);
    expect(rig.fake.pendingFrames()).toBe(0);

    rig.fake.fireTimers();
    rig.fake.fireFrames();

    expect(rig.ids()).toEqual([2, 4]);
  });

  test("exceeding maxDelay flushes with zero debounce", () => {
    const rig = makeRig();

    rig.edit(1);
    rig.fake.advance(100); // past maxDelay since the first edit
    rig.edit(2); // re-arms with delay 0
    rig.fake.fireTimers();
    rig.fake.fireFrames();

    expect(rig.ids()).toEqual([2]);
  });

  test("dispose cancels a pending run", () => {
    const rig = makeRig();

    rig.edit(1);
    rig.scheduler.dispose();
    rig.fake.fireTimers();
    rig.fake.fireFrames();

    expect(rig.runs).toHaveLength(0);
    expect(rig.fake.pendingTimers()).toBe(0);
    expect(rig.fake.pendingFrames()).toBe(0);
  });

  test("dispose cancels a pending leading frame", () => {
    const rig = makeRig({ leadingFrame: true });

    rig.edit(1);
    rig.scheduler.dispose();
    rig.fake.fireFrames();

    expect(rig.runs).toHaveLength(0);
    expect(rig.fake.pendingFrames()).toBe(0);
  });
});
