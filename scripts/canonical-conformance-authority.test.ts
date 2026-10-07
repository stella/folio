import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import ts from "typescript";

const parse = (source: string) =>
  ts.createSourceFile("authority.ts", source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

const fixedAuthority = (call: ts.CallExpression, authority: string) => {
  const options = call.arguments.at(0);
  if (!options || !ts.isObjectLiteralExpression(options)) return false;
  const property = options.properties.at(-1);
  return (
    property !== undefined &&
    ts.isPropertyAssignment(property) &&
    ts.isIdentifier(property.name) &&
    property.name.text === "authority" &&
    ts.isStringLiteral(property.initializer) &&
    property.initializer.text === authority
  );
};

const standingFactoryIsCanonical = (source: string) => {
  let found = false;
  const visit = (node: ts.Node) => {
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.name.text === "runConformanceCase"
    ) {
      const factory = node.initializer;
      found = Boolean(
        factory &&
        ts.isArrowFunction(factory) &&
        factory.parameters.length === 1 &&
        ts.isCallExpression(factory.body) &&
        ts.isIdentifier(factory.body.expression) &&
        factory.body.expression.text === "runCaseWithAuthority" &&
        fixedAuthority(factory.body, "canonical"),
      );
    }
    ts.forEachChild(node, visit);
  };
  visit(parse(source));
  return found;
};

const randomBrowserAuthorityProblems = (source: string) => {
  const problems: string[] = [];
  let calls = 0;
  const visit = (node: ts.Node, randomLane = false) => {
    const inLane =
      randomLane || (ts.isForOfStatement(node) && node.expression.getText() === "config.seeds");
    if (inLane && ts.isCallExpression(node)) {
      if (ts.isIdentifier(node.expression) && node.expression.text === "runMode") {
        calls += 1;
        if (!fixedAuthority(node, "canonical")) problems.push("random browser authority");
      }
      if (
        ts.isPropertyAccessExpression(node.expression) &&
        node.expression.getText() === "test.fail"
      )
        problems.push("canonical random lane accepts a legacy failure");
    }
    ts.forEachChild(node, (child) => visit(child, inLane));
  };
  visit(parse(source));
  if (calls === 0) problems.push("missing standing browser calls");
  return problems;
};

test("standing conformance cannot select legacy authority through its public factory", () => {
  const source = readFileSync(
    new URL("../packages/core/src/__tests__/editorCommandConformance.ts", import.meta.url),
    "utf8",
  );
  expect(standingFactoryIsCanonical(source)).toBe(true);
  expect(
    standingFactoryIsCanonical(
      source.replaceAll('authority: "canonical"', 'authority: "prosemirror"'),
    ),
  ).toBe(false);
  expect(standingFactoryIsCanonical("const runConformanceCase = () => legacy();")).toBe(false);
  expect(
    standingFactoryIsCanonical(
      'const runConformanceCase = (options) => runCaseWithAuthority({ authority: "canonical", ...options });',
    ),
  ).toBe(false);
});

test("random browser conformance is strict canonical; legacy acceptance stays in explicit replay", () => {
  const source = readFileSync(
    new URL("../tests/visual/browser-input-fuzz.interactions.spec.ts", import.meta.url),
    "utf8",
  );
  expect(randomBrowserAuthorityProblems(source)).toEqual([]);
  const lane = 'for (const seed of config.seeds) { runMode({ authority: "canonical" }); }';
  expect(randomBrowserAuthorityProblems(lane)).toEqual([]);
  expect(randomBrowserAuthorityProblems(lane.replace('"canonical"', '"prosemirror"'))).toHaveLength(
    1,
  );
  expect(randomBrowserAuthorityProblems(lane.replace("});", "}); test.fail(true);"))).toHaveLength(
    1,
  );
  expect(randomBrowserAuthorityProblems("for (const seed of config.seeds) {}")).toHaveLength(1);
});

test("refusal expectation owners cannot read command descriptor or planner verdicts", () => {
  const forbiddenImports = (source: string) =>
    parse(source).statements.flatMap((node) => {
      if (!ts.isImportDeclaration(node) || !ts.isStringLiteral(node.moduleSpecifier)) return [];
      const module = node.moduleSpecifier.text;
      return /(?:^|\/)(?:canonicalCommands|canonicalStructure)(?:\.[cm]?tsx?)?$/u.test(module)
        ? [module]
        : [];
    });
  for (const owner of ["canonicalEditorHarness", "canonical-conformance-refusals"]) {
    const source = readFileSync(new URL(`../test/${owner}.ts`, import.meta.url), "utf8");
    expect(forbiddenImports(source)).toEqual([]);
    for (const module of ["canonicalCommands", "canonicalStructure"]) {
      expect(
        forbiddenImports(`${source}\nimport { verdict } from "../owner/${module}";`),
      ).toHaveLength(1);
    }
  }
});
