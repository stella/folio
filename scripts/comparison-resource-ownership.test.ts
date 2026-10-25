import { expect, test } from "bun:test";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  declaresBindings,
  type Node,
  nodeText,
  parseSourceModule,
  staticMemberName,
  walkAst,
} from "./lib/oxc-ast";

const ROOT = path.resolve(import.meta.dir, "..");
const OWNERS = new Set([
  "packages/core/src/ai-edits/headless.ts",
  "packages/core/src/compare/import-style-closure.ts",
]);
const IMPORT_STEPS = new Set([
  "planTargetNumberingReferences",
  "stageTargetNumbering",
  "stageTargetStyles",
]);

const violations = (source: string): string[] => {
  if (![...IMPORT_STEPS].some((name) => source.includes(name))) return [];
  const { program } = parseSourceModule("fixture.ts", source);
  const found: string[] = [];
  const memberName = (node: Node, ancestors: readonly Node[]) => {
    const accessed = staticMemberName(node);
    if (accessed !== undefined) return accessed;
    if (
      ancestors.at(-1)?.type !== "ObjectPattern" ||
      !declaresBindings(ancestors)
    )
      return undefined;
    if (node.type === "Property")
      return node.computed ? undefined : nodeText(source, node.key);
    if (node.type === "RestElement") return nodeText(source, node.argument);
    return undefined;
  };
  walkAst(program, (node, ancestors) => {
    const member = memberName(node, ancestors);
    if (member !== undefined && IMPORT_STEPS.has(member))
      found.push(nodeText(source, node));
  });
  return found;
};

test("resource import steps cannot escape the shared sequence through aliases", () => {
  expect(violations("access.stageTargetStyles({})")).toHaveLength(1);
  expect(
    violations("const plan = access['planTargetNumberingReferences']; plan()"),
  ).toHaveLength(1);
  expect(
    violations("const { stageTargetNumbering: stage } = access; stage()"),
  ).toHaveLength(1);
  expect(violations("importStyleClosureWithNumbering({})")).toEqual([]);
});

test("production resource imports use the shared numbering and style owner", () => {
  const files = execFileSync("git", ["ls-files", "packages/core/src"], {
    cwd: ROOT,
  })
    .toString()
    .trim()
    .split("\n")
    .filter(
      (file) =>
        file.endsWith(".ts") && !file.endsWith(".test.ts") && !OWNERS.has(file),
    );
  const found = files.flatMap((file) =>
    violations(readFileSync(path.join(ROOT, file), "utf8")).map(
      (entry) => `${file}: ${entry}`,
    ),
  );
  expect(found).toEqual([]);
});
