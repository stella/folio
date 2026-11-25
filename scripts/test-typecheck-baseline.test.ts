import { describe, expect, test } from "bun:test";

import {
  compareCounts,
  countDiagnostics,
  parseDiagnostics,
} from "./test-typecheck-baseline";

const ROOT = "/checkout";

describe("test typecheck baseline", () => {
  test("fingerprints ignore positions and chained detail", () => {
    const output = [
      "packages/a/src/x.test.ts(3,5): error TS2322: Type 'string' is not assignable to type 'number'.",
      "packages/a/src/x.test.ts(40,1): error TS2322: Type 'string' is not assignable to type 'number'.",
      "packages/a/src/y.test.ts(1,1): error TS2769: No overload matches this call.",
      "  Overload 1 of 2, '(a: string): void', gave the following error.",
      "    Argument of type 'Z' is not assignable.",
      "packages/a/src/y.test.ts(9,2): error TS2307: Cannot find module '/checkout/packages/a/src/z'.",
      "",
      "Found 4 errors.",
    ].join("\n");
    const counts = countDiagnostics(
      parseDiagnostics(output, { cwd: ROOT, root: ROOT }),
    );
    expect(counts).toEqual({
      "packages/a/src/x.test.ts": {
        "TS2322: Type 'string' is not assignable to type 'number'.": 2,
      },
      "packages/a/src/y.test.ts": {
        "TS2307: Cannot find module 'packages/a/src/z'.": 1,
        "TS2769: No overload matches this call.": 1,
      },
    });
  });

  test("a new error or a higher count is added; a lower count is fixed", () => {
    const baseline = {
      "a.test.ts": { "TS1: one": 2, "TS2: two": 1 },
      "gone.test.ts": { "TS3: three": 1 },
    };
    const current = {
      "a.test.ts": { "TS1: one": 3, "TS2: two": 1 },
      "new.test.ts": { "TS4: four": 1 },
    };
    expect(compareCounts(baseline, current)).toEqual({
      added: [
        { file: "a.test.ts", fingerprint: "TS1: one", from: 2, to: 3 },
        { file: "new.test.ts", fingerprint: "TS4: four", from: 0, to: 1 },
      ],
      fixed: [
        { file: "gone.test.ts", fingerprint: "TS3: three", from: 1, to: 0 },
      ],
    });
  });

  test("an unchanged set compares clean", () => {
    const counts = { "a.test.ts": { "TS1: one": 1 } };
    expect(compareCounts(counts, counts)).toEqual({ added: [], fixed: [] });
  });
});
