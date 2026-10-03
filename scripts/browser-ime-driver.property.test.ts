import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { browserImeActionArbitrary } from "../tests/visual/browserInputTrace";
import { runBrowserImeLifecycle } from "../tests/visual/browserImeDriver";

setDefaultTimeout(propertyTestTimeout(5_000));

test("generated IME lifecycles deliver every update and exactly their declared completion", async () => {
  await assertProperty(
    fc.asyncProperty(browserImeActionArbitrary, async (action) => {
      const events: { kind: "update" | "commit" | "cancel"; text?: string }[] = [];
      await runBrowserImeLifecycle(
        {
          update: async (text) => {
            events.push({ kind: "update", text });
          },
          commit: async (text) => {
            events.push({ kind: "commit", text });
          },
          cancel: async () => {
            events.push({ kind: "cancel" });
          },
        },
        action,
      );
      expect(events.slice(0, -1)).toEqual(action.updates.map((text) => ({ kind: "update", text })));
      expect(events.at(-1)).toEqual(
        action.completion === "commit"
          ? { kind: "commit", text: action.updates.at(-1) }
          : { kind: "cancel" },
      );
    }),
    { numRuns: 100 },
  );
});

test("empty IME lifecycles reject before sending a browser command", async () => {
  for (const completion of ["commit", "cancel"] as const) {
    const events: string[] = [];
    await expect(
      runBrowserImeLifecycle(
        {
          update: async () => {
            events.push("update");
          },
          commit: async () => {
            events.push("commit");
          },
          cancel: async () => {
            events.push("cancel");
          },
        },
        { kind: "imeReplacement", updates: [], completion },
      ),
    ).rejects.toThrow("IME lifecycle has no updates");
    expect(events).toEqual([]);
  }
});

test("failed updates propagate without completing the composition", async () => {
  const failure = new TypeError("Browser command failed");
  const events: string[] = [];
  await expect(
    runBrowserImeLifecycle(
      {
        update: async (text) => {
          events.push(text);
          if (text === "東京") throw failure;
        },
        commit: async () => {
          events.push("commit");
        },
        cancel: async () => {
          events.push("cancel");
        },
      },
      { kind: "imeReplacement", updates: ["alpha", "東京", "café"], completion: "commit" },
    ),
  ).rejects.toBe(failure);
  expect(events).toEqual(["alpha", "東京"]);
});
