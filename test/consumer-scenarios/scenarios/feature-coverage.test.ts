import assert from "node:assert/strict";
import { test } from "node:test";

import { paragraphFeatures, type Feature } from "../support/targets.ts";

test("paragraph feature scan follows nested arrays and content controls", () => {
  const found = new Map<string, Set<Feature>>();
  paragraphFeatures(
    [
      {
        type: "sdt",
        content: [
          {
            type: "paragraph",
            paraId: "A1B2C3D4",
            content: [{ type: "run", content: [{ type: "simpleField" }] }],
          },
        ],
      },
    ],
    found,
    new Set(),
  );
  assert.deepEqual([...(found.get("A1B2C3D4") ?? [])].sort(), ["contentControl", "field"]);
});
