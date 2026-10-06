import { expect, test } from "bun:test";
import path from "node:path";
import ts from "typescript";
import type { CanonicalCommandIntent } from "../packages/core/src/prosemirror/canonicalCommands";

const root = path.resolve(import.meta.dir, "..");
const producers = {
  formatRun: [
    "canonicalCommands.ts#canonicalRunFormatting",
    "extensions/marks/markUtils.ts#canonicalSetMarkFormatting",
  ],
  setHyperlink: ["extensions/marks/HyperlinkExtension.ts#setHyperlink"],
  removeHyperlink: ["extensions/marks/HyperlinkExtension.ts#removeHyperlink"],
  insertHyperlink: ["extensions/marks/HyperlinkExtension.ts#insertHyperlink"],
  insertBreak: ["commands/pageBreak.ts#insertPageBreak"],
} as const satisfies Record<
  Extract<CanonicalCommandIntent, { from: number; to: number }>["type"],
  readonly string[]
>;

const typeOf = (node: ts.ObjectLiteralExpression) => {
  for (const property of node.properties) {
    if (!ts.isPropertyAssignment(property) || property.name.getText() !== "type") continue;
    const value = ts.isAsExpression(property.initializer)
      ? property.initializer.expression
      : property.initializer;
    if (ts.isStringLiteral(value)) return value.text;
  }
  return undefined;
};

const producerDeclaration = (
  node: ts.Node,
): ts.VariableDeclaration | ts.FunctionDeclaration | undefined => {
  if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name)) return node;
  if (ts.isFunctionDeclaration(node)) return node;
  if (!node.parent) return undefined;
  return producerDeclaration(node.parent);
};

const rangeInputRoot = (node: ts.Node): ts.Node => {
  let current = node;
  while (current.parent) {
    const parent = current.parent;
    if (
      ts.isCallExpression(parent) &&
      parent.expression.getText() === "withCanonicalCommand" &&
      parent.arguments.at(1) === current
    )
      return current;
    current = parent;
  }
  const declaration = producerDeclaration(node);
  if (declaration && ts.isVariableDeclaration(declaration))
    return declaration.initializer ?? declaration;
  return declaration ?? node;
};

const memberName = (node: ts.Node) => {
  if (ts.isPropertyAccessExpression(node)) return node.name.text;
  if (ts.isElementAccessExpression(node) && ts.isStringLiteral(node.argumentExpression))
    return node.argumentExpression.text;
  return undefined;
};

/** Follow descriptor helpers, regardless of their names or selection binding syntax. */
const rawRangeInputs = (source: ts.SourceFile, roots: readonly ts.Node[]) => {
  const helpers = new Map<string, ts.Node[]>();
  const declarations: ts.VariableDeclaration[] = [];
  const collect = (node: ts.Node): void => {
    if (ts.isVariableDeclaration(node)) {
      declarations.push(node);
      if (
        ts.isIdentifier(node.name) &&
        node.initializer &&
        (ts.isArrowFunction(node.initializer) || ts.isFunctionExpression(node.initializer))
      ) {
        const entries = helpers.get(node.name.text) ?? [];
        entries.push(node.initializer);
        helpers.set(node.name.text, entries);
      }
    }
    if (ts.isFunctionDeclaration(node) && node.name) {
      const entries = helpers.get(node.name.text) ?? [];
      entries.push(node);
      helpers.set(node.name.text, entries);
    }
    ts.forEachChild(node, collect);
  };
  collect(source);
  const aliases = new Set<string>();
  const isSelection = (node: ts.Node) =>
    memberName(node) === "selection" || (ts.isIdentifier(node) && aliases.has(node.text));
  let previousSize = -1;
  while (aliases.size !== previousSize) {
    previousSize = aliases.size;
    for (const declaration of declarations) {
      if (!declaration.initializer) continue;
      if (ts.isIdentifier(declaration.name) && isSelection(declaration.initializer))
        aliases.add(declaration.name.text);
      if (ts.isObjectBindingPattern(declaration.name)) {
        for (const binding of declaration.name.elements) {
          if (
            (binding.propertyName ?? binding.name).getText() === "selection" &&
            ts.isIdentifier(binding.name)
          )
            aliases.add(binding.name.text);
        }
      }
    }
  }
  const failures: string[] = [];
  const visited = new Set<ts.Node>();
  const visit = (node: ts.Node): void => {
    if (visited.has(node)) return;
    visited.add(node);
    if (
      (ts.isPropertyAccessExpression(node) || ts.isElementAccessExpression(node)) &&
      ["from", "to"].includes(memberName(node) ?? "") &&
      isSelection(node.expression)
    )
      failures.push(node.getText());
    if (
      ts.isVariableDeclaration(node) &&
      node.initializer &&
      isSelection(node.initializer) &&
      ts.isObjectBindingPattern(node.name)
    ) {
      for (const binding of node.name.elements) {
        if (["from", "to"].includes((binding.propertyName ?? binding.name).getText()))
          failures.push(binding.getText());
      }
    }
    if (ts.isCallExpression(node) && ts.isIdentifier(node.expression)) {
      for (const helper of helpers.get(node.expression.text) ?? []) visit(helper);
    }
    ts.forEachChild(node, visit);
  };
  for (const rootNode of roots) visit(rootNode);
  return failures;
};

test("range guard follows the called helper across raw selection binding shapes", () => {
  const bindings = [
    "return { from: state.selection.from, to: state.selection.to };",
    "const { from, to, empty } = state.selection; return { from, to };",
    "const { from: left, to: right } = state.selection; return { from: left, to: right };",
    "const { selection: selected } = state; const { from, to } = selected; return { from, to };",
    "const selected = state['selection']; return { from: selected['from'], to: selected['to'] };",
    "return canonicalSelectionRange(state);",
  ];
  for (const [index, binding] of bindings.entries()) {
    const parsed = ts.createSourceFile(
      "fixture.ts",
      `
      import { canonicalSelectionRange } from './canonicalSelectionRange';
      const canonicalRemovalRange = (state) => { ${binding} };
      const descriptor = (state) => {
        const range = canonicalRemovalRange(state);
        return [{ type: 'removeHyperlink', ...range }];
      };
      withCanonicalCommand(raw, descriptor);
    `,
      ts.ScriptTarget.Latest,
      true,
    );
    const descriptor = parsed.statements.find(
      (statement) =>
        ts.isVariableStatement(statement) &&
        statement.declarationList.declarations.some(
          (declaration) => declaration.name.getText() === "descriptor",
        ),
    );
    if (!descriptor) throw new TypeError("Range guard fixture lost its descriptor");
    expect(rawRangeInputs(parsed, [descriptor]).length > 0).toBe(index < bindings.length - 1);
  }
});

test("every canonical range producer normalizes selection endpoints", async () => {
  const packageJson = await Bun.file(path.join(root, "package.json")).json();
  expect(packageJson.scripts["test:source-guards"].split(" ")).toContain(
    "scripts/canonical-range-producers.test.ts",
  );
  const found = new Map<string, string[]>();
  const glob = new Bun.Glob("**/*.ts");
  const directory = path.join(root, "packages/core/src/prosemirror");
  for await (const file of glob.scan({ cwd: directory })) {
    if (file.endsWith(".test.ts")) continue;
    const source = await Bun.file(path.join(directory, file)).text();
    const parsed = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true);
    const roots: ts.Node[] = [];
    const visit = (node: ts.Node): void => {
      if (ts.isObjectLiteralExpression(node)) {
        const type = typeOf(node);
        if (type !== undefined && Object.hasOwn(producers, type)) {
          roots.push(rangeInputRoot(node));
          const declaration = producerDeclaration(node);
          const owner = `${file}#${declaration?.name?.getText()}`;
          const entries = found.get(type) ?? [];
          entries.push(owner);
          found.set(type, entries);
          // Direct endpoints must be computed carrier positions; selection ranges use the helper.
          for (const property of node.properties) {
            if (
              ts.isPropertyAssignment(property) &&
              ["from", "to"].includes(property.name.getText())
            ) {
              expect(owner).toBe("extensions/marks/markUtils.ts#canonicalSetMarkFormatting");
              expect(property.initializer.getText()).toBe(
                `representation.${property.name.getText()}`,
              );
            }
            if (ts.isSpreadAssignment(property)) {
              const range = property.expression.getText();
              if (range === "range") {
                expect(owner).toBe("extensions/marks/HyperlinkExtension.ts#removeHyperlink");
              } else if (range.includes("Selection") || range.includes("selection")) {
                expect(range).toBe("canonicalSelectionRange(state)");
              }
            }
          }
          if (owner.endsWith("#canonicalSetMarkFormatting")) {
            expect(declaration?.getText()).toContain("canonicalSelectionRange(state)");
            expect(declaration?.getText()).not.toMatch(/state\.selection\.(from|to)/u);
          } else {
            expect(
              node.properties.some(
                (property) =>
                  ts.isSpreadAssignment(property) &&
                  (property.expression.getText() === "canonicalSelectionRange(state)" ||
                    (type === "removeHyperlink" && property.expression.getText() === "range")),
              ),
            ).toBe(true);
          }
        }
      }
      // Range intermediates cannot bypass normalization through destructured raw endpoints.
      if (
        ts.isVariableDeclaration(node) &&
        node.initializer?.getText() === "state.selection" &&
        ts.isObjectBindingPattern(node.name)
      ) {
        const names = node.name.elements.map((element) =>
          (element.propertyName ?? element.name).getText(),
        );
        const producer = producerDeclaration(node.parent)?.name?.getText();
        if (producer === "canonicalSetMarkFormatting" || producer === "removalRange") {
          expect(names).not.toContain("from");
          expect(names).not.toContain("to");
        }
      }
      ts.forEachChild(node, visit);
    };
    visit(parsed);
    if (roots.length > 0) expect(rawRangeInputs(parsed, roots)).toEqual([]);
    if (file === "extensions/marks/HyperlinkExtension.ts") {
      const removal = source.slice(
        source.indexOf("const removalRange"),
        source.indexOf("const removeHyperlink"),
      );
      expect(removal).toContain("canonicalSelectionRange(state)");
      expect(removal).not.toMatch(/state\.selection\.(from|to)/u);
    }
  }
  expect([...found.keys()].sort()).toEqual(Object.keys(producers).sort());
  for (const [type, expected] of Object.entries(producers)) {
    expect(found.get(type)?.sort()).toEqual([...expected].sort());
  }
});
