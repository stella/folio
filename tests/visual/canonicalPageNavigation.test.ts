import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";

import {
  nodeText,
  parseSourceModule,
  walkAst,
} from "../../scripts/lib/oxc-ast";

test("canonical history page evaluations use the navigation owner", () => {
  for (const name of [
    "canonicalBrowserHistoryOracle.ts",
    "canonicalBrowserInputDriver.ts",
    "canonicalTimerProbe.ts",
  ]) {
    const { program, text } = parseSourceModule(
      name,
      readFileSync(new URL(name, import.meta.url), "utf8"),
    );
    const evaluationOwner =
      name === "canonicalTimerProbe.ts"
        ? "evaluateCanonicalDocument"
        : "evaluateCanonicalPage";
    let evaluations = 0;
    walkAst(program, (node, ancestors) => {
      if (
        node.type === "CallExpression" &&
        node.callee.type === "MemberExpression" &&
        !node.callee.computed &&
        nodeText(text, node.callee.object) === "page" &&
        node.callee.property.type === "Identifier" &&
        node.callee.property.name === "evaluate"
      ) {
        evaluations++;
        expect(
          ancestors.findLast(
            (parent) =>
              parent.type === "CallExpression" &&
              parent.callee.type === "Identifier" &&
              parent.callee.name === evaluationOwner,
          ),
          `${name}: evaluation bypasses the navigation owner`,
        ).toBeDefined();
      }
    });
    expect(
      evaluations,
      `${name}: guard must exercise page evaluations`,
    ).toBeGreaterThan(0);
  }
});
