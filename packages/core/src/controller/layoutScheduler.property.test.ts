/**
 * Property-based tests for the layout scheduler's invariants: any burst of
 * schedule() calls that lands inside the coalescing window collapses to exactly
 * one layout run (a typing storm paints once), and every run lays out the
 * editor's state as it is when the run happens, whatever replaced it after the
 * request (#1142).
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { createLayoutScheduler, type SchedulerClock } from "./layoutScheduler";

setDefaultTimeout(propertyTestTimeout(30_000));

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
  };
};

const SCHEDULER_OP = {
  edit: "edit",
  replace: "replace",
  advance: "advance",
  fireTimers: "fire-timers",
  fireFrames: "fire-frames",
} as const;

type SchedulerOp =
  | { type: typeof SCHEDULER_OP.edit; id: number }
  | { type: typeof SCHEDULER_OP.replace; id: number }
  | { type: typeof SCHEDULER_OP.advance; ms: number }
  | { type: typeof SCHEDULER_OP.fireTimers }
  | { type: typeof SCHEDULER_OP.fireFrames };

const schedulerOp: fc.Arbitrary<SchedulerOp> = fc.oneof(
  fc.record({ type: fc.constant(SCHEDULER_OP.edit), id: fc.integer() }),
  fc.record({ type: fc.constant(SCHEDULER_OP.replace), id: fc.integer() }),
  fc.record({ type: fc.constant(SCHEDULER_OP.advance), ms: fc.integer({ min: 0, max: 200 }) }),
  fc.record({ type: fc.constant(SCHEDULER_OP.fireTimers) }),
  fc.record({ type: fc.constant(SCHEDULER_OP.fireFrames) }),
);

describe("layoutScheduler (properties)", () => {
  test("a coalesced burst runs layout once, with the state current at the frame", () => {
    fc.assert(
      fc.property(fc.array(fc.integer(), { minLength: 1, maxLength: 40 }), (burst) => {
        const fake = makeFakeClock();
        const view: { state: TestState } = { state: { id: Number.NaN } };
        const runs: TestState[] = [];
        const scheduler = createLayoutScheduler({
          readState: () => view.state,
          runLayout: (state) => {
            runs.push(state);
          },
          debounceMs: 32,
          maxDelayMs: 96,
          clock: fake.clock,
        });

        // Whole burst scheduled at t=0, so nothing exceeds maxDelay — it must
        // coalesce into a single pass.
        for (const id of burst) {
          view.state = { id };
          scheduler.schedule();
        }
        fake.fireTimers(); // debounce elapses -> arms one frame
        fake.fireFrames(); // frame runs the single layout

        expect(runs).toHaveLength(1);
        expect(runs[0]).toBe(view.state);
      }),
      propertyConfig(),
    );
  });

  test("every run lays out the state the editor holds when it runs", () => {
    fc.assert(
      fc.property(fc.array(schedulerOp, { maxLength: 60 }), fc.boolean(), (ops, leadingFrame) => {
        const fake = makeFakeClock();
        const view: { state: TestState } = { state: { id: 0 } };
        const stale: TestState[] = [];
        const scheduler = createLayoutScheduler({
          readState: () => view.state,
          runLayout: (state) => {
            if (state !== view.state) {
              stale.push(state);
            }
          },
          debounceMs: 32,
          leadingFrame,
          maxDelayMs: 96,
          clock: fake.clock,
        });

        for (const op of ops) {
          switch (op.type) {
            case SCHEDULER_OP.edit:
              view.state = { id: op.id };
              scheduler.schedule();
              break;
            case SCHEDULER_OP.replace:
              // A wholesale replacement (document load) schedules nothing.
              view.state = { id: op.id };
              break;
            case SCHEDULER_OP.advance:
              fake.advance(op.ms);
              break;
            case SCHEDULER_OP.fireTimers:
              fake.fireTimers();
              break;
            case SCHEDULER_OP.fireFrames:
              fake.fireFrames();
              break;
            default: {
              const unreachable: never = op;
              throw new Error(`unhandled op ${JSON.stringify(unreachable)}`);
            }
          }
        }

        expect(stale).toEqual([]);
      }),
      propertyConfig(),
    );
  });

  test("dispose before the frame fires cancels the run entirely", () => {
    fc.assert(
      fc.property(fc.integer({ min: 1, max: 40 }), (burstLength) => {
        const fake = makeFakeClock();
        let runCount = 0;
        const scheduler = createLayoutScheduler({
          readState: () => ({ id: 0 }),
          runLayout: () => {
            runCount += 1;
          },
          debounceMs: 32,
          maxDelayMs: 96,
          clock: fake.clock,
        });

        for (let index = 0; index < burstLength; index += 1) {
          scheduler.schedule();
        }
        scheduler.dispose();
        fake.fireTimers();
        fake.fireFrames();

        expect(runCount).toBe(0);
      }),
      propertyConfig(),
    );
  });
});
