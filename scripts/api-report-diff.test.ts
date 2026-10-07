import { describe, expect, test } from "bun:test";

import { isStaleDeclaration, renderReportDiff, reportNewlineKind } from "./lib/api-report-diff";

const diff = (baseline: string, candidate: string, maxLines = 300): string =>
  renderReportDiff({ baseline, candidate, maxLines });

describe("reportNewlineKind", () => {
  test("new reports use LF without depending on the host platform", () => {
    expect(reportNewlineKind(undefined)).toBe("lf");
  });

  for (const [kind, newline] of [
    ["lf", "\n"],
    ["crlf", "\r\n"],
  ] as const) {
    test(`existing ${kind} snapshots retain their line endings across signature edits`, () => {
      for (const count of [1, 2, 10]) {
        const snapshot = Array.from({ length: count }, () => "export type A = 1;").join(newline);
        expect(reportNewlineKind(`${snapshot}${newline}`)).toBe(kind);
        expect(reportNewlineKind(`${snapshot}${newline}export type B = 2;${newline}`)).toBe(kind);
      }
    });
  }
});

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
