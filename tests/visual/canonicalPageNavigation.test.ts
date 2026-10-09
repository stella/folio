import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import ts from "typescript";

test("canonical history page evaluations use the navigation owner", () => {
  for (const name of [
    "canonicalBrowserHistoryOracle.ts",
    "canonicalBrowserInputDriver.ts",
    "canonicalTimerProbe.ts",
  ]) {
    const source = ts.createSourceFile(
      name,
      readFileSync(new URL(name, import.meta.url), "utf8"),
      ts.ScriptTarget.Latest,
      true,
    );
    const evaluationOwner =
      name === "canonicalTimerProbe.ts" ? "evaluateCanonicalDocument" : "evaluateCanonicalPage";
    let evaluations = 0;
    const visit = (node: ts.Node) => {
      if (
        ts.isCallExpression(node) &&
        ts.isPropertyAccessExpression(node.expression) &&
        node.expression.expression.getText(source) === "page" &&
        node.expression.name.text === "evaluate"
      ) {
        evaluations++;
        let parent: ts.Node | undefined = node.parent;
        while (
          parent &&
          !(
            ts.isCallExpression(parent) &&
            ts.isIdentifier(parent.expression) &&
            parent.expression.text === evaluationOwner
          )
        ) {
          parent = parent.parent;
        }
        expect(parent, `${name}: evaluation bypasses the navigation owner`).toBeDefined();
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
    expect(evaluations, `${name}: guard must exercise page evaluations`).toBeGreaterThan(0);
  }
});
