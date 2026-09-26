import { describe, expect, test } from "bun:test";

import { debounce } from "./debounce";

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

describe("debounce", () => {
  test("runs once, after the last call of a burst", async () => {
    let runs = 0;
    const debounced = debounce(() => {
      runs += 1;
    }, 20);

    debounced.schedule();
    debounced.schedule();
    await wait(5);
    debounced.schedule();

    expect(runs).toBe(0);
    await wait(60);
    expect(runs).toBe(1);
  });

  test("does not run once disposed", async () => {
    let runs = 0;
    const debounced = debounce(() => {
      runs += 1;
    }, 10);

    debounced.schedule();
    debounced.dispose();
    await wait(40);

    expect(runs).toBe(0);
  });
});
