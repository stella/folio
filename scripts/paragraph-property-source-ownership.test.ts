/** Static ownership guard for hidden paragraph-property captures. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import path from "node:path";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OWNER = "packages/core/src/docx/paragraphPropertySource.ts";
const PM_PROVENANCE_FILES = new Set([
  "packages/core/src/ai-edits/headless.ts",
  "packages/core/src/prosemirror/conversion/fromProseDoc.ts",
  "packages/core/src/prosemirror/conversion/toProseDoc.ts",
  "packages/core/src/prosemirror/extensions/features/AutoBidiDetectionExtension.ts",
  "packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension.ts",
]);

const relativePath = (sourceFile: ts.SourceFile): string =>
  path.relative(REPO_ROOT, sourceFile.fileName).replaceAll("\\", "/");

const containsParagraph = (type: ts.Type, location: ts.Node, checker: ts.TypeChecker): boolean => {
  if (type.isUnionOrIntersection()) {
    return type.types.some((part) => containsParagraph(part, location, checker));
  }
  const discriminator = type.getProperty("type");
  if (!discriminator) {
    return false;
  }
  const discriminatorType = checker.getTypeOfSymbolAtLocation(discriminator, location);
  return discriminatorType.isStringLiteral() && discriminatorType.value === "paragraph";
};

const calledMemberName = (node: ts.CallExpression): string | null => {
  if (!ts.isPropertyAccessExpression(node.expression)) {
    return null;
  }
  return node.expression.name.text;
};

const isPmNodeTypeCreate = (node: ts.CallExpression): boolean => {
  if (!ts.isPropertyAccessExpression(node.expression) || node.expression.name.text !== "create") {
    return false;
  }
  const owner = node.expression.expression;
  return ts.isPropertyAccessExpression(owner) && owner.name.text === "type";
};

const containsDocument = (type: ts.Type, location: ts.Node, checker: ts.TypeChecker): boolean => {
  if (type.isUnionOrIntersection()) {
    return type.types.some((part) => containsDocument(part, location, checker));
  }
  const packageProperty = type.getProperty("package");
  if (!packageProperty) {
    return false;
  }
  const packageType = checker.getTypeOfSymbolAtLocation(packageProperty, location);
  return packageType.getProperty("document") !== undefined;
};

const ownershipViolations = (): string[] => {
  const configPath = path.join(REPO_ROOT, "packages/core/tsconfig.build.json");
  const config = ts.readConfigFile(configPath, ts.sys.readFile);
  if (config.error) {
    panic(ts.flattenDiagnosticMessageText(config.error.messageText, "\n"));
  }
  const parsed = ts.parseJsonConfigFileContent(config.config, ts.sys, path.dirname(configPath));
  const program = ts.createProgram({ options: parsed.options, rootNames: parsed.fileNames });
  const checker = program.getTypeChecker();
  const violations: string[] = [];

  for (const sourceFile of program.getSourceFiles()) {
    const file = relativePath(sourceFile);
    if (
      file === OWNER ||
      !file.startsWith("packages/core/src/") ||
      file.endsWith(".test.ts") ||
      file.includes("/generated/")
    ) {
      continue;
    }
    const visit = (node: ts.Node): void => {
      if (
        ts.isSpreadAssignment(node) &&
        containsParagraph(checker.getTypeAtLocation(node.expression), node.expression, checker)
      ) {
        const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
        violations.push(`${file}:${String(line + 1)} spreads a Paragraph`);
      }
      if (ts.isCallExpression(node)) {
        if (
          ts.isIdentifier(node.expression) &&
          node.expression.text === "structuredClone" &&
          node.arguments.some((argument) =>
            containsDocument(checker.getTypeAtLocation(argument), argument, checker),
          )
        ) {
          const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
          violations.push(`${file}:${String(line + 1)} calls structuredClone`);
        }
        const member = calledMemberName(node);
        if (
          PM_PROVENANCE_FILES.has(file) &&
          (member === "copy" || member === "setNodeMarkup" || isPmNodeTypeCreate(node))
        ) {
          const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
          violations.push(`${file}:${String(line + 1)} calls PM ${member ?? "clone"}`);
        }
      }
      ts.forEachChild(node, visit);
    };
    visit(sourceFile);
  }
  return violations;
};

setDefaultTimeout(30_000);

describe("paragraph property source ownership", () => {
  test("every typed and provenance-aware PM clone uses the owning helper", () => {
    expect(ownershipViolations()).toEqual([]);
  });
});
