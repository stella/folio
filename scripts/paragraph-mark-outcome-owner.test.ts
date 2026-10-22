import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";
import ts from "typescript";
import { scanTestTimeout } from "../test/scan-timeout";

setDefaultTimeout(scanTestTimeout(10_000));

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const EXPECTED_CALLS = new Map([
  ["packages/docx-core/src/ops/resolve.ts", 2],
  ["packages/docx-core/src/ops/blockDeletion.ts", 1],
  ["packages/core/src/internal/wholeStoryRevisionResolution.ts", 1],
  ["packages/core/src/prosemirror/tableRunIn.ts", 1],
  ["packages/core/src/prosemirror/commands/comments.ts", 1],
]);

const outcomeCalls = (file: string, sourceText: string): number => {
  const source = ts.createSourceFile(file, sourceText, ts.ScriptTarget.Latest, true);
  let calls = 0;
  const visit = (node: ts.Node): void => {
    if (
      ts.isCallExpression(node) &&
      ts.isIdentifier(node.expression) &&
      node.expression.text === "paragraphMarkOutcome"
    ) {
      calls++;
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return calls;
};

const violations = (sources: ReadonlyMap<string, string>): string[] =>
  [...EXPECTED_CALLS]
    .filter(([file, expected]) => outcomeCalls(file, sources.get(file) ?? "") !== expected)
    .map(([file]) => file);

test("every paragraph-mark resolver delegates its structural decision to the owner", () => {
  const sources = new Map(
    [...EXPECTED_CALLS.keys()].map((file) => [
      file,
      ts.sys.readFile(path.join(REPO_ROOT, file)) ?? "",
    ]),
  );
  expect(violations(sources)).toEqual([]);
});

test("a caller that replaces the shared decision is rejected", () => {
  const sources = new Map(
    [...EXPECTED_CALLS].map(([file, count]) => [
      file,
      Array.from({ length: count }, () => "paragraphMarkOutcome({});").join("\n"),
    ]),
  );
  sources.set("packages/core/src/prosemirror/commands/comments.ts", "const remove = true;");
  expect(violations(sources)).toEqual(["packages/core/src/prosemirror/commands/comments.ts"]);
});
