import assert from "node:assert/strict";
import { test } from "node:test";

import { comparePreservedLinks } from "./link-oracle.ts";

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
