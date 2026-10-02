import { describe, expect, test } from "bun:test";

import { staleIgnoreReason } from "./audit";

const ignore = {
  advisory: "GHSA-test",
  packageName: "vulnerable",
  vulnerableThrough: "1.4.0",
  parent: "parent",
  expires: "2026-11-01",
};

describe("self-expiring audit ignores", () => {
  test("an ignore applies while no fix is published and before it expires", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: true },
        "2026-10-02",
      ),
    ).toBeUndefined();
  });

  test("a published fix ends the ignore", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.1", stillRequiredByParent: true },
        "2026-10-02",
      ),
    ).toContain("vulnerable@1.4.1 is published");
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "2.0.0", stillRequiredByParent: true },
        "2026-10-02",
      ),
    ).toContain("is published");
  });

  test("a parent release that drops the package ends the ignore", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: false },
        "2026-10-02",
      ),
    ).toContain("no longer requires vulnerable");
  });

  test("the ignore expires on its date", () => {
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: true },
        "2026-11-01",
      ),
    ).toBeUndefined();
    expect(
      staleIgnoreReason(
        ignore,
        { latestVersion: "1.4.0", stillRequiredByParent: true },
        "2026-11-02",
      ),
    ).toContain("expired on 2026-11-01");
  });
});
