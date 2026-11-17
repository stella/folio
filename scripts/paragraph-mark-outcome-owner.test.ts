import { expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import { parseSourceModule, walkAst } from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(10_000));

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OWNER_NAMES = [
  "paragraphMarkOutcome",
  "paragraphMarkContainerEnd",
] as const;
const EXPECTED_CALLS = {
  "packages/docx-core/src/ops/resolve.ts": {
    paragraphMarkOutcome: 2,
    paragraphMarkContainerEnd: 0,
  },
  "packages/docx-core/src/ops/blockDeletion.ts": {
    paragraphMarkOutcome: 1,
    paragraphMarkContainerEnd: 0,
  },
  "packages/core/src/internal/wholeStoryRevisionResolution.ts": {
    paragraphMarkOutcome: 2,
    paragraphMarkContainerEnd: 1,
  },
  "packages/core/src/prosemirror/tableRunIn.ts": {
    paragraphMarkOutcome: 1,
    paragraphMarkContainerEnd: 0,
  },
  "packages/core/src/prosemirror/commands/comments.ts": {
    paragraphMarkOutcome: 1,
    paragraphMarkContainerEnd: 1,
  },
} as const satisfies Record<
  string,
  Record<(typeof OWNER_NAMES)[number], number>
>;

const readSource = (file: string): string =>
  existsSync(file) ? readFileSync(file, "utf8") : "";

type OwnerCallsOptions = {
  file: string;
  sourceText: string;
  owner: (typeof OWNER_NAMES)[number];
};
const ownerCalls = ({ file, sourceText, owner }: OwnerCallsOptions): number => {
  const { program } = parseSourceModule(file, sourceText);
  let calls = 0;
  walkAst(program, (node) => {
    if (
      node.type === "CallExpression" &&
      node.callee.type === "Identifier" &&
      node.callee.name === owner
    ) {
      calls++;
    }
  });
  return calls;
};

const violations = (sources: ReadonlyMap<string, string>): string[] =>
  Object.entries(EXPECTED_CALLS)
    .filter(([file, expected]) =>
      OWNER_NAMES.some(
        (owner) =>
          ownerCalls({ file, sourceText: sources.get(file) ?? "", owner }) !==
          expected[owner],
      ),
    )
    .map(([file]) => file);

test("every paragraph-mark resolver delegates its structural decision to the owner", () => {
  const sources = new Map(
    Object.keys(EXPECTED_CALLS).map((file) => [
      file,
      readSource(path.join(REPO_ROOT, file)),
    ]),
  );
  expect(violations(sources)).toEqual([]);
});

test("a caller that replaces the shared decision is rejected", () => {
  const sources = new Map(
    Object.entries(EXPECTED_CALLS).map(([file, counts]) => [
      file,
      OWNER_NAMES.flatMap((owner) =>
        Array.from({ length: counts[owner] }, () => `${owner}({});`),
      ).join("\n"),
    ]),
  );
  sources.set(
    "packages/core/src/prosemirror/commands/comments.ts",
    "const remove = true;",
  );
  expect(violations(sources)).toEqual([
    "packages/core/src/prosemirror/commands/comments.ts",
  ]);
});

test("a caller that replaces the shared logical container classification is rejected", () => {
  const sources = new Map(
    Object.entries(EXPECTED_CALLS).map(([file, counts]) => [
      file,
      OWNER_NAMES.flatMap((owner) =>
        Array.from({ length: counts[owner] }, () => `${owner}({});`),
      ).join("\n"),
    ]),
  );
  sources.set(
    "packages/core/src/prosemirror/commands/comments.ts",
    "paragraphMarkOutcome({}); const canRemove = precedingParagraph !== null;",
  );
  expect(violations(sources)).toEqual([
    "packages/core/src/prosemirror/commands/comments.ts",
  ]);
});
