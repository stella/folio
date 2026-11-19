import { describe, expect, test } from "bun:test";

import { trustViolation } from "./check-lockfile-trust";

describe("lockfile trusted dependencies", () => {
  test("a lock that records the manifest's list passes", () => {
    expect(
      trustViolation(
        ".",
        { trustedDependencies: [] },
        { trustedDependencies: [] },
      ),
    ).toBeUndefined();
    expect(
      trustViolation(
        ".",
        { trustedDependencies: ["b", "a"] },
        { trustedDependencies: ["a", "b"] },
      ),
    ).toBeUndefined();
  });

  test("a lock that omits an empty list names the fix", () => {
    expect(trustViolation(".", { trustedDependencies: [] }, {})).toContain(
      "Run `bun install`",
    );
  });
});
