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

test("every canonical range producer normalizes selection endpoints", async () => {
  const found = new Map<string, string[]>();
  const glob = new Bun.Glob("**/*.ts");
  const directory = path.join(root, "packages/core/src/prosemirror");
  for await (const file of glob.scan({ cwd: directory })) {
    if (file.endsWith(".test.ts")) continue;
    const source = await Bun.file(path.join(directory, file)).text();
    const parsed = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true);
    const visit = (node: ts.Node): void => {
      if (ts.isObjectLiteralExpression(node)) {
        const type = typeOf(node);
        if (type !== undefined && Object.hasOwn(producers, type)) {
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
