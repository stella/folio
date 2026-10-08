import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import ts from "typescript";
import JSZip from "jszip";
import { buildPanelLayoutDocument } from "../tests/support/panelLayoutDocument";
import type { PanelLayoutReview, PanelLayoutSections } from "../tests/support/panelLayoutDocument";
import { buildScrollRootDocument } from "../tests/support/scrollRootDocument";
import { generateDocxFixture, validateDocxFixture } from "../tests/support/validatedDocxFixture";

const root = resolve(import.meta.dir, "..");
const sharedRoots = ["tests/support", "tests/visual/fixtures"];
const reviewCases = {
  none: "none",
  "changes-only": "changes-only",
  "comment-and-changes": "comment-and-changes",
} as const satisfies Record<PanelLayoutReview, PanelLayoutReview>;
const sectionCases = {
  portrait: "portrait",
  "landscape-then-portrait": "landscape-then-portrait",
} as const satisfies Record<PanelLayoutSections, PanelLayoutSections>;
const fixtureCases = [
  { file: "tests/support/scrollRootDocument.ts", build: buildScrollRootDocument },
  ...Object.values(reviewCases).flatMap((review) =>
    Object.values(sectionCases).map((sections) => ({
      file: "tests/support/panelLayoutDocument.ts",
      build: () => buildPanelLayoutDocument(review, sections),
    })),
  ),
];

const packageCalls = (text: string) => {
  const source = ts.createSourceFile("fixture.ts", text, ts.ScriptTarget.Latest, true);
  const serializers = new Set(["createDocx"]);
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement)) continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings || !ts.isNamedImports(bindings)) continue;
    for (const binding of bindings.elements) {
      if ((binding.propertyName ?? binding.name).text === "createDocx")
        serializers.add(binding.name.text);
    }
  }
  const calls: string[] = [];
  const memberName = (node: ts.Node) => {
    if (ts.isPropertyAccessExpression(node)) return node.name.text;
    if (ts.isElementAccessExpression(node) && ts.isStringLiteral(node.argumentExpression))
      return node.argumentExpression.text;
    return undefined;
  };
  const visit = (node: ts.Node): void => {
    const member = memberName(node);
    if (member === "generateAsync") calls.push("unchecked-zip");
    if (member === "createDocx") calls.push("unchecked-serializer");
    if (
      ts.isBindingElement(node) &&
      (node.propertyName ?? node.name).getText(source) === "generateAsync"
    ) {
      calls.push("unchecked-zip");
    }
    if (
      ts.isIdentifier(node) &&
      serializers.has(node.text) &&
      !ts.isImportSpecifier(node.parent) &&
      !(ts.isPropertyAccessExpression(node.parent) && node.parent.name === node)
    ) {
      calls.push("unchecked-serializer");
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return calls;
};

test("every shared fixture and generator uses validated package construction", () => {
  const unguarded: string[] = [];
  const supportProducers: string[] = [];
  for (const directory of sharedRoots) {
    for (const file of new Bun.Glob("**/*.ts").scanSync({
      cwd: resolve(root, directory),
      onlyFiles: true,
    })) {
      const relative = `${directory}/${file}`;
      if (relative === "tests/support/validatedDocxFixture.ts") continue;
      const source = readFileSync(resolve(root, relative), "utf8");
      if (packageCalls(source).length > 0) unguarded.push(relative);
      if (directory === "tests/support" && /\bgenerateDocxFixture\s*\(/u.test(source)) {
        supportProducers.push(relative);
      }
    }
  }
  expect(unguarded).toEqual([]);
  expect(supportProducers.sort()).toEqual(
    [...new Set(fixtureCases.map(({ file }) => file))].sort(),
  );
});

test("fixture guard detects direct package generators", () => {
  expect(packageCalls('return zip.generateAsync({type: "uint8array"})')).toEqual(["unchecked-zip"]);
  expect(packageCalls("return createDocx(document)")).toEqual(["unchecked-serializer"]);
  expect(
    packageCalls(`import { createDocx as serialize } from "core"; return serialize(document)`),
  ).toEqual(["unchecked-serializer"]);
  expect(packageCalls(`return core["createDocx"](document)`)).toEqual(["unchecked-serializer"]);
  expect(packageCalls(`const generate = zip["generateAsync"]; return generate(options)`)).toEqual([
    "unchecked-zip",
  ]);
  expect(packageCalls(`const { generateAsync: generate } = zip; return generate(options)`)).toEqual(
    ["unchecked-zip"],
  );
  expect(packageCalls('return generateDocxFixture(zip, "valid")')).toEqual([]);
});

test("every shared fixture variant passes schema validation at setup", async () => {
  for (const { file, build } of fixtureCases) {
    await validateDocxFixture(await build(), file);
  }
}, 60_000);

test("invalid fixture fails at construction with its name and schema reason", async () => {
  const zip = await JSZip.loadAsync(await buildScrollRootDocument());
  const part = zip.file("word/document.xml");
  if (!part) throw new Error("Missing fixture document part");
  zip.file("word/document.xml", (await part.async("string")).replace(' w:header="720"', ""));
  await expect(generateDocxFixture(zip, "deliberately-invalid")).rejects.toThrow(
    /Invalid DOCX fixture deliberately-invalid: invalid_schema_attribute/u,
  );
});
