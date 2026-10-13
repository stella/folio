import assert from "node:assert/strict";
import { test } from "node:test";

import { comparePreservedLinks } from "../support/link-oracle.ts";

test("missing saved projection is a failure only when reader text is unchanged", () => {
  const id = "7A1B2C3D";
  const text = "Read the schedule.";
  const target = "https://example.com/schedule";
  const targets = Array.from({ length: text.length }, (_, offset) =>
    offset >= 9 && offset < 17 ? target : null,
  );
  const before = new Map([[id, { text, targets }]]);
  const missing = new Map(before);
  missing.delete(id);

  assert.deepEqual(comparePreservedLinks({ before, after: missing, afterRows: [{ id, text }] }), {
    problems: [`block ${id} ("${text}") has no matching link projection`],
    checked: 1,
  });
  assert.deepEqual(
    comparePreservedLinks({ before, after: missing, afterRows: [{ id, text: `${text} revised` }] }),
    {
      problems: [],
      checked: 0,
    },
  );
});

test("empty paragraphs have no character link targets to preserve", () => {
  const id = "21D02091";
  const before = new Map([[id, { text: "", targets: [] }]]);
  for (const after of [new Map(), new Map(before)]) {
    for (const afterRows of [[], [{ id, text: "" }], [{ id, text: "Changed" }]]) {
      assert.deepEqual(comparePreservedLinks({ before, after, afterRows }), {
        problems: [],
        checked: 0,
      });
    }
  }
});
