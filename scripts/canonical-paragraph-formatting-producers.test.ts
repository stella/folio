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
  "packages/core/src/layout-bridge/convert/endnoteLayout.ts:93",
  "packages/core/src/utils/createDocument.ts:469",
  "packages/docx-core/src/legal-source/compile.ts:228",
  "packages/docx-core/src/legal-source/compile.ts:242",
  "packages/docx-core/src/markdown/content.ts:69",
  "packages/docx-core/src/ops/tocIntent.ts:157",
  "packages/docx-core/src/ops/tocIntent.ts:173",
  "packages/docx-core/src/serialize/docx.ts:628",
  "packages/docx-core/src/serialize/docx.ts:686",
]);

test("paragraph formatting producers use the canonical finalizer", async () => {
  const producers: string[] = [];
  for (const root of SOURCE_ROOTS) {
    for (const file of await filesBelow(root)) {
      const sourceText = await Bun.file(file).text();
      const source = ts.createSourceFile(file, sourceText, ts.ScriptTarget.Latest, true);
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
            const line =
              source.getLineAndCharacterOfPosition(formattingProperty.getStart(source)).line + 1;
            producers.push(`${file}:${String(line)}`);
          }
        }
        ts.forEachChild(node, visit);
      };
      visit(source);
    }
  }

  expect(producers.toSorted()).toEqual([...ALLOWED_DIRECT_PRODUCERS].toSorted());
});
