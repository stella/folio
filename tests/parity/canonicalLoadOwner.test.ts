import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../test/property-testing";
import { waitForCanonicalLoadOwner } from "./canonicalLoadOwner";

test(
  "canonical loading waits for adoption even when repeated loads have identical content",
  async () => {
    const initialStates = ["previous", "absent"] as const;
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 1, max: 5 }),
        fc.string(),
        async (delay, text) => {
          for (const state of initialStates) {
            let previousOwner = { text };
            for (let load = 0; load < 2; load++) {
              const ownerBeforeLoad = previousOwner;
              const nextOwner = { text };
              let frames = 0;
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
        },
      ),
      {
        numRuns: 30,
        id: "canonical loading waits for adoption even when repeated loads have identical content",
      },
    );
  },
  propertyTestTimeout(5_000),
);
