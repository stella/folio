import { describe, expect, test } from "bun:test";

import {
  compareCounts,
  countDiagnostics,
  parseDiagnostics,
} from "./test-typecheck-baseline";

const ROOT = "/checkout";

describe("test typecheck baseline", () => {
  test("entries key on file and code only, never on lines or messages", () => {
    const output = [
      "packages/a/src/x.test.ts(3,5): error TS2322: Type 'string' is not assignable to type 'number'.",
      "packages/a/src/x.test.ts(40,1): error TS2322: Type 'Cut' is not assignable to type 'KeptCut'.",
      "packages/a/src/y.test.ts(1,1): error TS2769: No overload matches this call.",
      "  Overload 1 of 2, '(a: string): void', gave the following error.",
      "    Argument of type 'Z' is not assignable.",
      "/checkout/packages/a/src/y.test.ts(9,2): error TS2307: Cannot find module './z'.",
      "",
      "Found 4 errors.",
    ].join("\n");
    const diagnostics = parseDiagnostics(output, { cwd: ROOT, root: ROOT });
    expect(countDiagnostics(diagnostics)).toEqual({
      "packages/a/src/x.test.ts": { TS2322: 2 },
      "packages/a/src/y.test.ts": { TS2307: 1, TS2769: 1 },
    });
    expect(diagnostics.map(({ line, message }) => [line, message])).toEqual([
      [3, "Type 'string' is not assignable to type 'number'."],
      [40, "Type 'Cut' is not assignable to type 'KeptCut'."],
      [1, "No overload matches this call."],
      [9, "Cannot find module './z'."],
    ]);
  });

  test("a reworded message for the same code leaves the counts unchanged", () => {
    const before = parseDiagnostics(
      "p/a.test.ts(1,1): error TS2322: Type 'A' is not assignable to type 'B'.",
      { cwd: ROOT, root: ROOT },
    );
    const after = parseDiagnostics(
      "p/a.test.ts(7,3): error TS2322: Type 'A2' is not assignable to type 'B2'.",
      { cwd: ROOT, root: ROOT },
    );
    expect(
      compareCounts(countDiagnostics(before), countDiagnostics(after)),
    ).toEqual({ added: [], fixed: [] });
  });

  test("a new error or a higher count is added; a lower count is fixed", () => {
    const baseline = {
      "a.test.ts": { TS1: 2, TS2: 1 },
      "gone.test.ts": { TS3: 1 },
    };
    const current = {
      "a.test.ts": { TS1: 3, TS2: 1 },
      "new.test.ts": { TS4: 1 },
    };
    expect(compareCounts(baseline, current)).toEqual({
      added: [
        { file: "a.test.ts", code: "TS1", from: 2, to: 3 },
        { file: "new.test.ts", code: "TS4", from: 0, to: 1 },
      ],
      fixed: [{ file: "gone.test.ts", code: "TS3", from: 1, to: 0 }],
    });
  });
});
