import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const root = path.resolve(import.meta.dir, "..");

const missingCompilerExports = (file: string, source: string): string[] => {
  const parsed = ts.createSourceFile(file, source, ts.ScriptTarget.Latest);
  const missing = new Set<string>();
  const visit = (node: ts.Node): void => {
    if (
      (ts.isPropertyAccessExpression(node) ||
        ts.isElementAccessExpression(node)) &&
      ts.isIdentifier(node.expression) &&
      node.expression.text === "ts"
    ) {
      let member: string | undefined;
      if (ts.isPropertyAccessExpression(node)) member = node.name.text;
      else if (ts.isStringLiteral(node.argumentExpression))
        member = node.argumentExpression.text;
      if (member !== undefined && Reflect.get(ts, member) === undefined)
        missing.add(member);
    }
    ts.forEachChild(node, visit);
  };
  visit(parsed);
  return [...missing].toSorted();
};

test("compiler export guard detects unavailable APIs without treating types or text as values", () => {
  expect(
    missingCompilerExports(
      "fixture.ts",
      `
        import ts from "typescript";
        ts.createSourceFile("fixture.ts", "", ts.ScriptTarget.Latest);
        ts.unavailableCompilerApi();
        ts["unavailableCompilerApi"]();
        ts["anotherUnavailableCompilerApi"]();
        let node: ts.Node;
        type Missing = ts.TypeOnlyName;
        // ts.commentOnlyName();
        const text = "ts.stringOnlyName()";
      `,
    ),
  ).toEqual(["anotherUnavailableCompilerApi", "unavailableCompilerApi"]);
});

test("every script uses compiler APIs exported by the installed TypeScript version", () => {
  const files = [
    ...new Bun.Glob("scripts/**/*.{ts,tsx,mts,cts,js,jsx,mjs,cjs}").scanSync(
      root,
    ),
  ];
  expect(files.length).toBeGreaterThan(0);
  const missing = files.flatMap((file) =>
    missingCompilerExports(
      file,
      readFileSync(path.join(root, file), "utf8"),
    ).map(
      (member) =>
        `${file}: ts.${member} is unavailable in TypeScript ${ts.version}`,
    ),
  );
  expect(missing).toEqual([]);
});
