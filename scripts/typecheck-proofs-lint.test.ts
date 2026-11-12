/** Wiring test for the compile-time-proof placement lint rule. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const RULE_MARKER = "folio-typecheck-proofs(no-type-suppression-in-test)";
const GATE_MARKER = "folio-typecheck-proofs(no-unchecked-field-gate)";

setDefaultTimeout(subprocessTestTimeout());

const lintFixture = (fixture: string) => {
  const result = expectExited(
    runChild(
      [
        "bun",
        "--bun",
        "oxlint",
        "-c",
        "oxlint.fixtures.config.ts",
        "--no-ignore",
        path.join("test", "__fixtures__", fixture),
      ],
      { cwd: REPO_ROOT },
    ),
  );
  return `${result.stdout}${result.stderr}`;
};

const count = (output: string, marker: string): number =>
  output.split(marker).length - 1;

describe("compile-time proof placement", () => {
  test("rejects every suppression directive, in line and block comments", () => {
    expect(count(lintFixture("typecheck-proofs.invalid.ts"), RULE_MARKER)).toBe(
      4,
    );
  });

  test("leaves prose that names a directive alone", () => {
    expect(count(lintFixture("typecheck-proofs.valid.ts"), RULE_MARKER)).toBe(
      0,
    );
  });

  test("rejects a field-totality gate in a test the compiler does not read", () => {
    const output = lintFixture("typecheck-proofs.field-gate.invalid.ts");
    expect(count(output, GATE_MARKER)).toBe(4);
  });

  test("leaves a table over a value's own keys and a shape check alone", () => {
    const output = lintFixture("typecheck-proofs.field-gate.valid.ts");
    expect(count(output, GATE_MARKER)).toBe(0);
  });
});
