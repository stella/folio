import { expect, spyOn, test } from "bun:test";
import { CANONICAL_INPUT_TIMER_OWNER, setCanonicalInputTimer } from "./canonicalInputTimer";

test("only the scheduled wrapper carries canonical input ownership", () => {
  const schedule = spyOn(globalThis, "setTimeout");
  let calls = 0;
  const callback = () => {
    calls++;
  };
  try {
    const timer = setCanonicalInputTimer(callback, 60_000);
    clearTimeout(timer);
    expect(schedule.mock.calls.at(0)?.at(1)).toBe(60_000);
    const wrapper = schedule.mock.calls.at(0)?.at(0);
    if (typeof wrapper !== "function") throw new TypeError("Owned timer callback unavailable");
    expect(Reflect.get(wrapper, Symbol.for(CANONICAL_INPUT_TIMER_OWNER))).toBe(true);
    expect(Reflect.has(callback, Symbol.for(CANONICAL_INPUT_TIMER_OWNER))).toBe(false);
    wrapper();
    expect(calls).toBe(1);
    const unrelated = setTimeout(callback, 60_000);
    clearTimeout(unrelated);
    const unowned = schedule.mock.calls.at(1)?.at(0);
    if (typeof unowned !== "function") throw new TypeError("Unowned timer callback unavailable");
    expect(Reflect.has(unowned, Symbol.for(CANONICAL_INPUT_TIMER_OWNER))).toBe(false);
  } finally {
    schedule.mockRestore();
  }
});
