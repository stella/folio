import { describe, expect, test } from "bun:test";

import { isStaleDeclaration, renderReportDiff } from "./lib/api-report-diff";

const diff = (committed: string, generated: string, maxLines = 300): string =>
  renderReportDiff({ committed, generated, maxLines });

describe("renderReportDiff", () => {
  test("says nothing about lines both sides share", () => {
    expect(diff("a\nb\nc", "a\nb\nc")).toBe("");
  });

  test("marks the committed line and the generated one that replaced it", () => {
    expect(diff("export type A = 1;", "export type A = 2;")).toBe(
      "-export type A = 1;\n+export type A = 2;",
    );
  });

  test("reports an addition without reprinting the surrounding surface", () => {
    expect(diff("a\nc", "a\nb\nc")).toBe("+b");
  });

  test("reports a removal", () => {
    expect(diff("a\nb\nc", "a\nc")).toBe("-b");
  });

  test("keeps changes in document order across separate edits", () => {
    expect(diff("a\nb\nc\nd", "a\nB\nc\nD")).toBe("-b\n+B\n-d\n+D");
  });

  test("truncates a large diff and says how much it dropped", () => {
    const committed = Array.from({ length: 40 }, (_, index) => `old ${index}`).join("\n");
    const generated = Array.from({ length: 40 }, (_, index) => `new ${index}`).join("\n");
    const rendered = diff(committed, generated, 10);

    expect(rendered.split("\n")).toHaveLength(11);
    expect(rendered.endsWith("… 70 more diff line(s) truncated")).toBe(true);
  });

  test("does not truncate a diff that fits", () => {
    expect(diff("a", "b", 2)).toBe("-a\n+b");
  });
});

describe("isStaleDeclaration", () => {
  test("calls a declaration older than its sources stale", () => {
    expect(isStaleDeclaration({ declarationModifiedMs: 10, newestSourceModifiedMs: 20 })).toBe(
      true,
    );
  });

  test("accepts a declaration built after its sources", () => {
    expect(isStaleDeclaration({ declarationModifiedMs: 20, newestSourceModifiedMs: 10 })).toBe(
      false,
    );
  });

  test("accepts a build that finished in the same millisecond as the last write", () => {
    expect(isStaleDeclaration({ declarationModifiedMs: 10, newestSourceModifiedMs: 10 })).toBe(
      false,
    );
  });
});
