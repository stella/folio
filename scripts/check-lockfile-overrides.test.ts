import { describe, expect, test } from "bun:test";

import { overrideViolations } from "./check-lockfile-overrides";

const lockWith = (overrides: Record<string, string>, viewRange: string) =>
  ({
    overrides,
    workspaces: { "": { name: "root", devDependencies: { "prosemirror-model": "^1.25.9" } } },
    packages: {
      "prosemirror-view": [
        "prosemirror-view@1.42.6",
        "",
        { dependencies: { "prosemirror-model": viewRange } },
        "sha512-",
      ],
      "@stll/folio-core": ["@stll/folio-core@workspace:packages/core"],
    },
  }) as const;

const intentional = {
  "prosemirror-model": {
    dependent: "prosemirror-view",
    range: "^1.25.12",
    reason: "test",
  },
};

describe("lockfile override ranges", () => {
  test("an override inside every declared range passes", () => {
    expect(
      overrideViolations(lockWith({ "prosemirror-model": "1.25.12" }, "^1.25.12"), {}),
    ).toEqual([]);
  });

  test("an override below a dependent's minimum fails, naming the dependent", () => {
    expect(
      overrideViolations(lockWith({ "prosemirror-model": "1.25.11" }, "^1.25.12"), {}),
    ).toEqual(["prosemirror-view declares prosemirror-model ^1.25.12; the override pins 1.25.11."]);
  });

  test("a listed break permits exactly that dependent and range", () => {
    expect(
      overrideViolations(lockWith({ "prosemirror-model": "1.25.11" }, "^1.25.12"), intentional),
    ).toEqual([]);
    expect(
      overrideViolations(lockWith({ "prosemirror-model": "1.25.11" }, "^1.25.13"), intentional),
    ).toEqual([
      "prosemirror-view declares prosemirror-model ^1.25.13; the override pins 1.25.11.",
      "prosemirror-model: prosemirror-view no longer declares ^1.25.12 outside the override; remove the intentional break.",
    ]);
  });

  test("a listed break goes stale once the override satisfies the range or is removed", () => {
    expect(
      overrideViolations(lockWith({ "prosemirror-model": "1.25.12" }, "^1.25.12"), intentional),
    ).toEqual([
      "prosemirror-model: prosemirror-view no longer declares ^1.25.12 outside the override; remove the intentional break.",
    ]);
    expect(overrideViolations(lockWith({}, "^1.25.12"), intentional)).toEqual([
      "prosemirror-model: intentional break listed without an override; remove it.",
    ]);
  });
});
