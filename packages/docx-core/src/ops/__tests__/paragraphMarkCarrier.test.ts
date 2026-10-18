import { describe, expect, test } from "bun:test";

import { paragraphMarkCarrier } from "../blockDeletion";

describe("paragraphMarkCarrier", () => {
  test("the following paragraph owns a deleted break", () => {
    expect(paragraphMarkCarrier({ followingKind: "paragraph", endsContainer: false })).toBe(
      "following",
    );
  });

  test("a container-final paragraph keeps its required terminator", () => {
    expect(paragraphMarkCarrier({ followingKind: undefined, endsContainer: true })).toBe("current");
  });

  test("no paragraph carrier survives before a table", () => {
    expect(paragraphMarkCarrier({ followingKind: "table", endsContainer: false })).toBe("none");
  });
});
