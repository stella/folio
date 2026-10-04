import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../test/property-testing";
import { waitForCanonicalLoadOwner } from "./canonicalLoadOwner";

test(
  "canonical loading waits for adoption even when repeated loads have identical content",
  async () => {
    let cases = 0;
    const initialStates = ["previous", "absent"] as const;
    const exercised = new Set<(typeof initialStates)[number]>();
    await assertProperty(
      fc.asyncProperty(fc.integer({ min: 1, max: 5 }), fc.string(), async (delay, text) => {
        for (const state of initialStates) {
          exercised.add(state);
          let previousOwner = { text };
          for (let load = 0; load < 2; load++) {
            const ownerBeforeLoad = previousOwner;
            const nextOwner = { text };
            let frames = 0;
            cases += 1;
            await waitForCanonicalLoadOwner({
              previousOwner: ownerBeforeLoad,
              getOwner: () => {
                if (frames >= delay) return nextOwner;
                return state === "previous" ? ownerBeforeLoad : null;
              },
              waitFrame: async () => {
                frames += 1;
              },
            });
            expect(frames).toBe(delay);
            expect(nextOwner).toEqual(previousOwner);
            expect(nextOwner).not.toBe(previousOwner);
            previousOwner = nextOwner;
          }
        }
      }),
      { numRuns: 30 },
    );
    expect(cases).toBeGreaterThan(0);
    expect([...exercised].sort()).toEqual([...initialStates].sort());
  },
  propertyTestTimeout(5_000),
);
