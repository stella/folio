import { expect, test } from "bun:test";
import { canonicalTimerOwner, isCanonicalInputTimer } from "../tests/parity/canonicalTimerOwner";

const captured = (owner: string, caller: string) =>
  `Error\n    at window.setTimeout (<anonymous>:10:39)\n    at ${owner}\n    at ${caller}`;
const root = "http://localhost:4200/@fs/work/folio/packages";
const input = `Object.handleKeyDown (${root}/core/src/controller/canonicalInput.ts:119:33)`;

test("a layout timer called through canonical history belongs to the layout", () => {
  const owner = `${root}/react/src/paged-editor/PagedEditor.tsx:1339:50`;
  const stack = captured(owner, input);
  expect(canonicalTimerOwner(stack)).toBe(`at ${owner}`);
  expect(isCanonicalInputTimer(stack)).toBe(false);
});

test("composition and input timers are recognized by their own frame", () => {
  for (const module of ["canonicalComposition", "canonicalInput"]) {
    expect(
      isCanonicalInputTimer(
        captured(`schedule (${root}/core/src/controller/${module}.ts:98:13)`, input),
      ),
    ).toBe(true);
  }
});

test("an unrelated same-named module does not own canonical input", () => {
  expect(isCanonicalInputTimer(captured(`${root}/react/src/canonicalInput.ts:1:1`, input))).toBe(
    false,
  );
});

test("missing owner capture fails instead of declaring a clean state", () => {
  for (const stack of ["unavailable", "Error\n    at window.setTimeout (<anonymous>:10:39)"])
    expect(() => isCanonicalInputTimer(stack)).toThrow("Timer capture has no owner frame");
});
