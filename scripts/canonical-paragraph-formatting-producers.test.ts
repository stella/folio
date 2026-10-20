import { expect, setDefaultTimeout, test } from "bun:test";
import { readdir } from "node:fs/promises";
import path from "node:path";
import ts from "typescript";
import { scanTestTimeout } from "../test/scan-timeout";

setDefaultTimeout(scanTestTimeout(10_000));

const SOURCE_ROOTS = ["packages/core/src", "packages/docx-core/src"] as const;

const filesBelow = async (directory: string): Promise<string[]> => {
  const entries = await readdir(directory, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const child = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...(await filesBelow(child)));
    else if (
      /\.ts$/u.test(entry.name) &&
      !child.includes(`${path.sep}__tests__${path.sep}`) &&
      !/\.(?:test|property|typecheck)\.ts$/u.test(entry.name)
    ) {
      files.push(child);
    }
  }
  return files;
};

const ALLOWED_DIRECT_PRODUCERS = new Set([
  // Filled with non-empty or owner-finalized paragraph producers. Keep entries
  // explicit so adding another construction requires a deliberate review.
  "packages/core/src/layout-bridge/convert/endnoteLayout.ts#SEPARATOR_STORY#1",
  "packages/core/src/utils/createDocument.ts#createEmptyDocument#1",
  "packages/docx-core/src/legal-source/compile.ts#paragraph#1",
  "packages/docx-core/src/legal-source/compile.ts#plainParagraph#1",
  "packages/docx-core/src/markdown/content.ts#listPara#1",
  "packages/docx-core/src/ops/tocIntent.ts#tocParagraphs#1",
  "packages/docx-core/src/ops/tocIntent.ts#tocParagraphs#2",
  "packages/docx-core/src/serialize/docx.ts#serializeStyle#1",
  "packages/docx-core/src/serialize/docx.ts#serializeNumberingLevel#1",
]);

const enclosingDeclaration = (
  node: ts.Node,
  source: ts.SourceFile,
): { name: string; node: ts.Node } => {
  let nearestVariable: ts.VariableDeclaration | undefined;
  for (let ancestor = node.parent; ancestor; ancestor = ancestor.parent) {
    if (ts.isFunctionDeclaration(ancestor) && ancestor.name) {
      return { name: ancestor.name.text, node: ancestor };
    }
    if (ts.isMethodDeclaration(ancestor)) {
      return { name: ancestor.name.getText(source), node: ancestor };
    }
    if (
      (ts.isArrowFunction(ancestor) || ts.isFunctionExpression(ancestor)) &&
      ts.isVariableDeclaration(ancestor.parent)
    ) {
      return { name: ancestor.parent.name.getText(source), node: ancestor.parent };
    }
    if (ts.isFunctionExpression(ancestor) && ancestor.name) {
      return { name: ancestor.name.text, node: ancestor };
    }
    if (!nearestVariable && ts.isVariableDeclaration(ancestor)) nearestVariable = ancestor;
  }
  return nearestVariable
    ? { name: nearestVariable.name.getText(source), node: nearestVariable }
    : { name: "<module>", node: source };
};

const directProducerKeys = (file: string, sourceText: string): string[] => {
  const source = ts.createSourceFile(file, sourceText, ts.ScriptTarget.Latest, true);
  const declarationCounts = new Map<ts.Node, number>();
  const producers: string[] = [];
  const visit = (node: ts.Node): void => {
    if (ts.isObjectLiteralExpression(node)) {
      const typeProperty = node.properties.find(
        (property): property is ts.PropertyAssignment =>
          ts.isPropertyAssignment(property) &&
          property.name.getText(source) === "type" &&
          ts.isStringLiteral(property.initializer) &&
          property.initializer.text === "paragraph",
      );
      const formattingProperty = node.properties.find(
        (property): property is ts.PropertyAssignment | ts.ShorthandPropertyAssignment =>
          (ts.isPropertyAssignment(property) || ts.isShorthandPropertyAssignment(property)) &&
          property.name.getText(source) === "formatting",
      );
      if (typeProperty && formattingProperty) {
        const declaration = enclosingDeclaration(node, source);
        const ordinal = (declarationCounts.get(declaration.node) ?? 0) + 1;
        declarationCounts.set(declaration.node, ordinal);
        producers.push(`${file}#${declaration.name}#${String(ordinal)}`);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return producers;
};

test("paragraph formatting producers use the canonical finalizer", async () => {
  const producers: string[] = [];
  for (const root of SOURCE_ROOTS) {
    for (const file of await filesBelow(root)) {
      const sourceText = await Bun.file(file).text();
      producers.push(...directProducerKeys(file, sourceText));
    }
  }

  expect(producers.toSorted()).toEqual([...ALLOWED_DIRECT_PRODUCERS].toSorted());
});

test("producer keys are stable outside their declaration and ordinal within it", () => {
  const sourceText = `
const first = () => ({ type: "paragraph", formatting: {}, content: [] });
const second = () => ({ type: "paragraph", formatting: {}, content: [] });
`;
  const expected = ["fixture.ts#first#1", "fixture.ts#second#1"];

  expect(directProducerKeys("fixture.ts", sourceText)).toEqual(expected);
  expect(directProducerKeys("fixture.ts", `const unrelated = true;\n${sourceText}`)).toEqual(
    expected,
  );
  expect(
    directProducerKeys(
      "fixture.ts",
      sourceText.replace(
        'const second = () => ({ type: "paragraph", formatting: {}, content: [] });',
        `const second = () => [
  { type: "paragraph", formatting: {}, content: [] },
  { type: "paragraph", formatting: {}, content: [] },
];`,
      ),
    ),
  ).toEqual(["fixture.ts#first#1", "fixture.ts#second#1", "fixture.ts#second#2"]);
});
