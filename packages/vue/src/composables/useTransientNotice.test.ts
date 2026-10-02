import { expect, test } from "bun:test";
import { panic } from "better-result";
import { effectScope } from "vue";
import { useTransientNotice } from "./useTransientNotice";

const createClock = () => {
  let now = 0;
  let nextId = 0;
  const timers = new Map<number, { due: number; callback: () => void }>();
  return {
    setTimer: (callback: () => void, delay: number) => {
      const id = nextId++;
      timers.set(id, { due: now + delay, callback });
      return id;
    },
    clearTimer: (id: number) => {
      timers.delete(id);
    },
    advance: (delay: number) => {
      now += delay;
      for (const [id, timer] of timers) {
        if (timer.due > now) continue;
        timers.delete(id);
        timer.callback();
      }
    },
    pending: () => timers.size,
  };
};

test("notices expire, replacement gets its full lifetime, and disposal cancels pending work", () => {
  const clock = createClock();
  const scope = effectScope();
  const notice = scope.run(() => useTransientNotice(clock)) ?? panic("Expected notice scope");
  try {
    notice.show("First refusal");
    clock.advance(2_000);
    notice.show("Replacement refusal");
    clock.advance(1_000);
    expect(notice.message.value).toBe("Replacement refusal");
    expect(clock.pending()).toBe(1);
    clock.advance(2_000);
    expect(notice.message.value).toBeNull();
    expect(clock.pending()).toBe(0);
    notice.show("Last refusal");
    scope.stop();
    expect(notice.message.value).toBeNull();
    expect(clock.pending()).toBe(0);
    clock.advance(3_000);
    expect(notice.message.value).toBeNull();
  } finally {
    scope.stop();
  }
});
