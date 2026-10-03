import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { resolvePainterTarget } from "../tests/visual/painterTargetCommit";

test(
  "paint commits re-resolve changing targets and return one consistent snapshot",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.integer({ min: 1, max: 100 }), { minLength: 1, maxLength: 20 }),
        async (edits) => {
          let current = 0;
          let painted = -1;
          let listener: (() => void) | undefined;
          let subscriptions = 0;
          const pending = resolvePainterTarget({
            subscribe: (notify) => {
              listener = notify;
              subscriptions++;
              return () => {
                listener = undefined;
                subscriptions--;
              };
            },
            read: () =>
              painted === current ? { position: current, coordinate: painted * 10 } : null,
          });
          for (const edit of edits) {
            current += edit;
            listener?.();
            expect(subscriptions).toBe(1);
          }
          painted = current;
          listener?.();
          const result = await pending;
          expect(result).toEqual({ position: current, coordinate: current * 10 });
          expect(subscriptions).toBe(0);
          current++;
          expect(result.position * 10).toBe(result.coordinate);
        },
      ),
      { numRuns: 50 },
    );
  },
  propertyTestTimeout(10_000),
);

test("ready and failed reads release their painter subscriptions", async () => {
  for (const fails of [false, true]) {
    let subscribed = false;
    const failure = new TypeError("target read failed");
    const pending = resolvePainterTarget({
      subscribe: () => {
        subscribed = true;
        return () => {
          subscribed = false;
        };
      },
      read: () => {
        if (fails) throw failure;
        return { position: 7 };
      },
    });
    if (fails) await expect(pending).rejects.toBe(failure);
    else expect(await pending).toEqual({ position: 7 });
    expect(subscribed).toBe(false);
  }
});
