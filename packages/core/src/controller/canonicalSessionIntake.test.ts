import { expect, test } from "bun:test";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { createEmptyDocument } from "../utils/createDocument";

// A package reaches a session through `ensureParaIds`; a new blank document is
// not read from a package, so it has to start with what a session requires.
test("a session opens a new blank document", () => {
  const document = createEmptyDocument({ initialText: "Blank" });
  expect(createCanonicalSession(document).isOk()).toBe(true);
  expect(createCanonicalSession(createEmptyDocument()).isOk()).toBe(true);
});
