import { expect, test } from "bun:test";
import path from "node:path";

import baseline from "../.oxlint-plugins/folio-test-dom.baseline.json";
import {
  checkTestDomBaseline,
  compareTestDomBaseline,
  parseTestDomBaseline,
} from "./lib/test-dom-baseline";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const FILE = "packages/core/src/docx/tableCellStatedWidth.property.test.ts";

test("the current allowance is checked against the actual CI base or local merge base", () => {
  const result = checkTestDomBaseline({ root: REPO_ROOT, current: baseline });
  expect(result.baseSha).toMatch(/^[a-f\d]{40}$/u);
  expect(result.problems).toEqual([]);
  if (result.type === "bootstrap")
    expect(result.verifiedSources.sort()).toEqual(Object.keys(baseline).sort());
});

test("committed baselines permit only per-file shrinkage", () => {
  const base = { type: "committed", baseline: { [FILE]: 2 } } as const;
  expect(compareTestDomBaseline({ current: { [FILE]: 1 }, base })).toEqual([]);
  expect(compareTestDomBaseline({ current: {}, base })).toEqual([]);
  expect(compareTestDomBaseline({ current: { [FILE]: 3 }, base })).toHaveLength(
    1,
  );
  expect(
    compareTestDomBaseline({
      current: { "packages/core/src/new.test.ts": 1 },
      base,
    }),
  ).toHaveLength(1);
});

test.each([{ [FILE]: 1.5 }, { [FILE]: 0 }, { "../new.test.ts": 1 }])(
  "malformed current allowance fails closed: %o",
  (current) => {
    const base = { type: "committed", baseline: { [FILE]: 2 } } as const;
    expect(() => compareTestDomBaseline({ current, base })).toThrow();
  },
);

test("invalid current paths fail before Git or filesystem inspection", () => {
  expect(() =>
    checkTestDomBaseline({
      root: REPO_ROOT,
      current: { "../new.test.ts": 1 },
      requestedBase: "not-a-commit",
    }),
  ).toThrow("Invalid DOM baseline entry");
});

const BOOTSTRAP_CASES = {
  unchanged: { before: "existing source", now: "existing source", failures: 0 },
  changed: {
    before: "existing source",
    now: "new raw provider reference",
    failures: 1,
  },
  added: { before: undefined, now: "new source", failures: 1 },
  absent: { before: "existing source", now: undefined, failures: 1 },
} as const;

test.each(Object.entries(BOOTSTRAP_CASES))(
  "bootstrap binds every allowance to its exact base bytes: %s",
  (_name, { before, now, failures }) => {
    const sourceAtBase = new Map<string, Buffer>();
    const sourceNow = new Map<string, Buffer>();
    if (before !== undefined) sourceAtBase.set(FILE, Buffer.from(before));
    if (now !== undefined) sourceNow.set(FILE, Buffer.from(now));
    expect(
      compareTestDomBaseline({
        current: { [FILE]: 2 },
        base: { type: "bootstrap", sourceAtBase, sourceNow },
      }),
    ).toHaveLength(failures);
  },
);

test.each(["not-a-commit", "f".repeat(40)])(
  "an invalid actual base fails closed: %s",
  (requestedBase) => {
    expect(() =>
      checkTestDomBaseline({
        root: REPO_ROOT,
        current: baseline,
        requestedBase,
      }),
    ).toThrow();
  },
);

test("the documented manual sentinel resolves a real merge base", () => {
  const result = checkTestDomBaseline({
    root: REPO_ROOT,
    current: baseline,
    requestedBase: "unavailable-toolchain-base",
  });
  expect(result.baseSha).toMatch(/^[a-f\d]{40}$/u);
  expect(result.problems).toEqual([]);
});

test.each([
  "[]",
  '{"../new.test.ts":1}',
  '{"packages/core/src/new.test.ts":0}',
  '{"packages/core/src/new.test.ts":1.5}',
])("malformed baseline input fails closed: %s", (source) => {
  expect(() => parseTestDomBaseline(source)).toThrow();
});
