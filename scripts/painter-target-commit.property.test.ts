import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { resolvePainterTarget } from "../tests/visual/painterTargetCommit";

test(
  "paint commits re-resolve changing targets and return one consistent snapshot",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.integer({ min: 1, max: 100 }), {
          minLength: 1,
          maxLength: 20,
        }),
        async (edits) => {
          let current = 0;
          let painted = -1;
          let listener: (() => void) | undefined;
          let subscriptions = 0;
          const pending = resolvePainterTarget({
            commit: () => {},
            subscribe: (notify) => {
              listener = notify;
              subscriptions++;
              return () => {
                listener = undefined;
                subscriptions--;
              };
            },
            read: () =>
              painted === current
                ? { position: current, coordinate: painted * 10 }
                : null,
          });
          for (const edit of edits) {
            current += edit;
            listener?.();
            expect(subscriptions).toBe(1);
          }
          painted = current;
          listener?.();
          const result = await pending;
          expect(result).toEqual({
            position: current,
            coordinate: current * 10,
          });
          expect(subscriptions).toBe(0);
          current++;
          expect(result.position * 10).toBe(result.coordinate);
        },
      ),
      {
        numRuns: 50,
        id: "paint commits re-resolve changing targets and return one consistent snapshot",
      },
    );
  },
  propertyTestTimeout(10_000),
);

test("ready and failed reads release their painter subscriptions", async () => {
  for (const fails of [false, true]) {
    let subscribed = false;
    const failure = new TypeError("target read failed");
    const pending = resolvePainterTarget({
      commit: () => {},
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

test(
  "current paint owns targets even when stale coordinates successfully roundtrip",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 1, max: 100_000 }),
        fc.integer({ min: 1, max: 100 }),
        fc.integer({ min: 0, max: 10 }),
        fc.constantFrom("immediate", "virtualized"),
        async (current, delta, earlyPaints, visibility) => {
          let painted = current - delta;
          let visible = visibility === "immediate";
          let listener: (() => void) | undefined;
          let subscriptions = 0;
          let commits = 0;
          const pending = resolvePainterTarget({
            subscribe: (notify) => {
              listener = notify;
              subscriptions++;
              return () => {
                listener = undefined;
                subscriptions--;
              };
            },
            commit: () => {
              commits++;
              // A painter emits while laying out; reading before commit returns
              // can see its previous document even though coordinates roundtrip.
              for (let i = 0; i < earlyPaints; i++) listener?.();
              painted = current;
              listener?.();
            },
            read: () =>
              visible ? { position: current, coordinate: painted * 10 } : null,
          });
          if (visibility === "virtualized") {
            expect(subscriptions).toBe(1);
            visible = true;
            listener?.();
          }
          const target = await pending;
          expect(target).toEqual({
            position: current,
            coordinate: current * 10,
          });
          expect(commits).toBe(1);
          expect(subscriptions).toBe(0);
        },
      ),
      {
        numRuns: 50,
        id: "current paint owns targets even when stale coordinates successfully roundtrip",
      },
    );
  },
  propertyTestTimeout(10_000),
);

test("failed paint commits reject and release the subscription before reading", async () => {
  let subscribed = false;
  let reads = 0;
  const failure = new TypeError("paint failed");
  const pending = resolvePainterTarget({
    subscribe: (notify) => {
      // Synchronous subscription notifications also cannot read before commit.
      notify();
      subscribed = true;
      return () => {
        subscribed = false;
      };
    },
    commit: () => {
      throw failure;
    },
    read: () => {
      reads++;
      return { position: 7 };
    },
  });
  await expect(pending).rejects.toBe(failure);
  expect(subscribed).toBe(false);
  expect(reads).toBe(0);
});

/**
 * A font set whose loads finish in a later task, after every microtask the
 * read that started them queued: the next layout's fonts arrive late.
 */
const scriptedFonts = () => {
  let ready = Promise.resolve();
  const fonts = {
    status: "loaded" as FontFaceSetLoadStatus,
    get ready() {
      return ready;
    },
    start() {
      if (fonts.status === "loading") return;
      fonts.status = "loading";
      ready = new Promise((resolve) => {
        setTimeout(() => {
          fonts.status = "loaded";
          resolve();
        }, 0);
      });
    },
  };
  return fonts;
};

test(
  "a painted target is read only from paint no font load can move",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.boolean(),
        fc.array(fc.boolean(), { minLength: 1, maxLength: 6 }),
        async (loadingBefore, readsStartLoads) => {
          const fonts = scriptedFonts();
          if (loadingBefore) fonts.start();
          let commits = 0;
          let reads = 0;
          const committedWhileLoading: number[] = [];
          const pending = resolvePainterTarget({
            fonts,
            subscribe: () => () => {},
            commit: () => {
              commits++;
              if (fonts.status === "loading")
                committedWhileLoading.push(commits);
            },
            read: () => {
              const read = reads++;
              // Painting the read's text can start a subset load.
              if (readsStartLoads[read] === true) fonts.start();
              return { read };
            },
          });
          const target = await pending;
          // A read stands only when no load started with it.
          const expected = readsStartLoads.findIndex((starts) => !starts);
          expect(target).toEqual({
            read: expected === -1 ? readsStartLoads.length : expected,
          });
          expect(committedWhileLoading).toEqual([]);
          expect(commits).toBe(reads);
          expect(fonts.status).toBe("loaded");
        },
      ),
      {
        numRuns: 100,
        id: "a painted target is read only from paint no font load can move",
      },
    );
  },
  propertyTestTimeout(10_000),
);
