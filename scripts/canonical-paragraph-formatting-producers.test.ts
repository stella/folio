import { expect, setDefaultTimeout, test } from "bun:test";
import { readdir } from "node:fs/promises";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import { type Node, nodeText, parseSourceModule, walkAst } from "./lib/oxc-ast";

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

/** A binding as written, without the type annotation oxc includes in its span. */
const bindingText = (text: string, binding: Node) => {
  const annotation =
    binding.type === "Identifier" ||
    binding.type === "ObjectPattern" ||
    binding.type === "ArrayPattern"
      ? binding.typeAnnotation
      : undefined;
  return annotation
    ? text.slice(binding.start, annotation.start).trimEnd()
    : nodeText(text, binding);
};

/** The written name of a key; a computed key keeps its brackets. */
const keyText = (text: string, node: Node): string | undefined => {
  if (
    node.type !== "Property" &&
    node.type !== "MethodDefinition" &&
    node.type !== "TSAbstractMethodDefinition"
  )
    return undefined;
  const key = nodeText(text, node.key);
  return node.computed ? `[${key}]` : key;
};

const isMethod = (node: Node) =>
  ((node.type === "MethodDefinition" ||
    node.type === "TSAbstractMethodDefinition") &&
    node.kind === "method") ||
  (node.type === "Property" && node.method);

const enclosingDeclaration = (
  ancestors: readonly Node[],
  text: string,
): { name: string; node: Node } => {
  let nearestVariable: Node | undefined;
  for (let index = ancestors.length - 1; index >= 0; index -= 1) {
    const ancestor = ancestors[index];
    if (ancestor === undefined) continue;
    const parent = ancestors[index - 1];
    if (ancestor.type === "FunctionDeclaration" && ancestor.id) {
      return { name: ancestor.id.name, node: ancestor };
    }
    if (isMethod(ancestor)) {
      return { name: keyText(text, ancestor) ?? "", node: ancestor };
    }
    if (
      (ancestor.type === "ArrowFunctionExpression" ||
        ancestor.type === "FunctionExpression") &&
      parent?.type === "VariableDeclarator"
    ) {
      return { name: bindingText(text, parent.id), node: parent };
    }
    if (ancestor.type === "FunctionExpression" && ancestor.id) {
      return { name: ancestor.id.name, node: ancestor };
    }
    if (!nearestVariable && ancestor.type === "VariableDeclarator")
      nearestVariable = ancestor;
  }
  if (nearestVariable?.type === "VariableDeclarator")
    return {
      name: bindingText(text, nearestVariable.id),
      node: nearestVariable,
    };
  const [program] = ancestors;
  if (program === undefined) throw new TypeError("Producer has no module");
  return { name: "<module>", node: program };
};

const isPropertyAssignment = (node: Node) =>
  node.type === "Property" &&
  node.kind === "init" &&
  !node.method &&
  !node.shorthand;

const directProducerKeys = (file: string, sourceText: string): string[] => {
  const { program } = parseSourceModule(file, sourceText);
  const declarationCounts = new Map<Node, number>();
  const producers: string[] = [];
  walkAst(program, (node, ancestors) => {
    if (node.type !== "ObjectExpression") return;
    const typeProperty = node.properties.find(
      (property) =>
        isPropertyAssignment(property) &&
        property.type === "Property" &&
        keyText(sourceText, property) === "type" &&
        property.value.type === "Literal" &&
        property.value.value === "paragraph",
    );
    const formattingProperty = node.properties.find(
      (property) =>
        property.type === "Property" &&
        property.kind === "init" &&
        !property.method &&
        keyText(sourceText, property) === "formatting",
    );
    if (typeProperty && formattingProperty) {
      const declaration = enclosingDeclaration(ancestors, sourceText);
      const ordinal = (declarationCounts.get(declaration.node) ?? 0) + 1;
      declarationCounts.set(declaration.node, ordinal);
      producers.push(`${file}#${declaration.name}#${String(ordinal)}`);
    }
  });
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

  expect(producers.toSorted()).toEqual(
    [...ALLOWED_DIRECT_PRODUCERS].toSorted(),
  );
});

test("producer keys are stable outside their declaration and ordinal within it", () => {
  const sourceText = `
const first = () => ({ type: "paragraph", formatting: {}, content: [] });
const second = () => ({ type: "paragraph", formatting: {}, content: [] });
`;
  const expected = ["fixture.ts#first#1", "fixture.ts#second#1"];

  expect(directProducerKeys("fixture.ts", sourceText)).toEqual(expected);
  expect(
    directProducerKeys("fixture.ts", `const unrelated = true;\n${sourceText}`),
  ).toEqual(expected);
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
  ).toEqual([
    "fixture.ts#first#1",
    "fixture.ts#second#1",
    "fixture.ts#second#2",
  ]);
});
