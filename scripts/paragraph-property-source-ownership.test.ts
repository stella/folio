/** Static ownership guard for hidden paragraph-property captures. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import path from "node:path";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OWNER = "packages/core/src/docx/paragraphPropertySource.ts";
const TYPECHECK_CONFIGS = [
  "packages/core/tsconfig.build.json",
  "packages/react/tsconfig.build.json",
  "packages/vue/tsconfig.build.json",
] as const;
const PM_PARAGRAPH_PROVENANCE_FILES = new Set([
  "packages/core/src/ai-edits/headless.ts",
  "packages/core/src/prosemirror/conversion/fromProseDoc.ts",
  "packages/core/src/prosemirror/conversion/toProseDoc.ts",
  "packages/core/src/prosemirror/extensions/features/AutoBidiDetectionExtension.ts",
  "packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension.ts",
]);
const TOKEN_OWNER_FILES = new Set([
  OWNER,
  "packages/core/src/prosemirror/attrs/index.ts",
  "packages/core/src/prosemirror/conversion/fromProseDoc.ts",
  "packages/core/src/prosemirror/extensions/core/ParagraphExtension.ts",
  "packages/core/src/prosemirror/extensions/features/ParaIdAllocatorExtension.ts",
  "packages/core/src/prosemirror/schema/nodes.ts",
]);
const CONTRACT_OWNER_FILES = new Set([
  OWNER,
  "packages/core/src/prosemirror/conversion/fromProseDoc.ts",
  "packages/core/src/prosemirror/conversion/toProseDoc.ts",
  "packages/core/src/prosemirror/extensions/core/DocExtension.ts",
  "packages/core/src/prosemirror/yjsParagraphSourceContract.ts",
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

type ParagraphProvenanceRebuild = "copy" | "setNodeMarkup" | "type.create";

const paragraphProvenanceRebuild = (node: ts.CallExpression): ParagraphProvenanceRebuild | null => {
  if (!ts.isPropertyAccessExpression(node.expression)) {
    return null;
  }
  if (node.expression.name.text === "copy" || node.expression.name.text === "setNodeMarkup") {
    return node.expression.name.text;
  }
  if (node.expression.name.text !== "create") {
    return null;
  }
  const owner = node.expression.expression;
  return ts.isPropertyAccessExpression(owner) && owner.name.text === "type" ? "type.create" : null;
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

const isProductionSource = (file: string): boolean =>
  file.startsWith("packages/") &&
  file.includes("/src/") &&
  !file.includes("/node_modules/") &&
  !file.endsWith(".test.ts") &&
  !file.endsWith(".test.tsx") &&
  !file.endsWith(".spec.ts") &&
  !file.endsWith(".spec.tsx") &&
  !file.includes("/__tests__/") &&
  !file.includes("/generated/");

type ProseDocToBlocksBindings = {
  direct: ReadonlySet<string>;
  namespaces: ReadonlySet<string>;
};

const importsProseDocToBlocks = (sourceFile: ts.SourceFile): ProseDocToBlocksBindings => {
  const direct = new Set<string>();
  const namespaces = new Set<string>();
  if (relativePath(sourceFile).endsWith("/prosemirror/conversion/fromProseDoc.ts")) {
    direct.add("proseDocToBlocks");
  }
  for (const statement of sourceFile.statements) {
    if (
      !ts.isImportDeclaration(statement) ||
      !ts.isStringLiteral(statement.moduleSpecifier) ||
      (!statement.moduleSpecifier.text.endsWith("/fromProseDoc") &&
        !statement.moduleSpecifier.text.endsWith("/prosemirror/conversion"))
    ) {
      continue;
    }
    const namedBindings = statement.importClause?.namedBindings;
    if (!namedBindings) {
      continue;
    }
    if (ts.isNamespaceImport(namedBindings)) {
      namespaces.add(namedBindings.name.text);
      continue;
    }
    for (const element of namedBindings.elements) {
      if ((element.propertyName ?? element.name).text === "proseDocToBlocks") {
        direct.add(element.name.text);
      }
    }
  }
  return { direct, namespaces };
};

const proseConversionViolations = (): string[] => {
  const violations: string[] = [];
  const sourcePaths = ts.sys.readDirectory(path.join(REPO_ROOT, "packages"), [".ts", ".tsx"]);
  for (const sourcePath of sourcePaths) {
    const file = path.relative(REPO_ROOT, sourcePath).replaceAll("\\", "/");
    if (!isProductionSource(file)) {
      continue;
    }
    const sourceText = ts.sys.readFile(sourcePath);
    if (sourceText === undefined) {
      panic(`Cannot read ${file}.`);
    }
    if (!sourceText.includes("proseDocToBlocks")) {
      continue;
    }
    const sourceFile = ts.createSourceFile(
      sourcePath,
      sourceText,
      ts.ScriptTarget.Latest,
      true,
      sourcePath.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
    );
    const proseDocToBlocksBindings = importsProseDocToBlocks(sourceFile);
    if (
      proseDocToBlocksBindings.direct.size === 0 &&
      proseDocToBlocksBindings.namespaces.size === 0
    ) {
      continue;
    }
    const visit = (node: ts.Node): void => {
      const callsDirectBinding =
        ts.isCallExpression(node) &&
        ts.isIdentifier(node.expression) &&
        proseDocToBlocksBindings.direct.has(node.expression.text);
      const callsNamespaceBinding =
        ts.isCallExpression(node) &&
        ts.isPropertyAccessExpression(node.expression) &&
        node.expression.name.text === "proseDocToBlocks" &&
        ts.isIdentifier(node.expression.expression) &&
        proseDocToBlocksBindings.namespaces.has(node.expression.expression.text);
      if (
        ts.isCallExpression(node) &&
        (callsDirectBinding || callsNamespaceBinding) &&
        node.arguments.length < 2
      ) {
        const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
        violations.push(
          `${file}:${String(line + 1)} converts a story without its property-source base`,
        );
      }
      ts.forEachChild(node, visit);
    };
    visit(sourceFile);
  }
  return violations;
};

const cloneOwnershipViolations = (): string[] => {
  const violations = new Set<string>();
  const visited = new Set<string>();

  for (const relativeConfigPath of TYPECHECK_CONFIGS) {
    const configPath = path.join(REPO_ROOT, relativeConfigPath);
    const config = ts.readConfigFile(configPath, ts.sys.readFile);
    if (config.error) {
      panic(ts.flattenDiagnosticMessageText(config.error.messageText, "\n"));
    }
    const parsed = ts.parseJsonConfigFileContent(config.config, ts.sys, path.dirname(configPath));
    const program = ts.createProgram({ options: parsed.options, rootNames: parsed.fileNames });
    const checker = program.getTypeChecker();

    for (const sourceFile of program.getSourceFiles()) {
      const file = relativePath(sourceFile);
      if (visited.has(file) || file === OWNER || !isProductionSource(file)) {
        continue;
      }
      visited.add(file);
      const visit = (node: ts.Node): void => {
        if (
          file.startsWith("packages/core/src/") &&
          ts.isSpreadAssignment(node) &&
          containsParagraph(checker.getTypeAtLocation(node.expression), node.expression, checker)
        ) {
          const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
          violations.add(`${file}:${String(line + 1)} spreads a Paragraph`);
        }
        if (
          ts.isCallExpression(node) &&
          ts.isIdentifier(node.expression) &&
          node.expression.text === "structuredClone" &&
          node.arguments.some((argument) =>
            containsDocument(checker.getTypeAtLocation(argument), argument, checker),
          )
        ) {
          const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
          violations.add(`${file}:${String(line + 1)} calls structuredClone on a Document`);
        }
        const provenanceRebuild = ts.isCallExpression(node)
          ? paragraphProvenanceRebuild(node)
          : null;
        if (provenanceRebuild && PM_PARAGRAPH_PROVENANCE_FILES.has(file)) {
          const { line } = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
          violations.add(`${file}:${String(line + 1)} calls PM ${provenanceRebuild} directly`);
        }
        ts.forEachChild(node, visit);
      };
      visit(sourceFile);
    }
  }
  return [...violations].toSorted();
};

const tokenOwnershipViolation = (file: string, sourceText: string): string | null => {
  if (!isProductionSource(file) || TOKEN_OWNER_FILES.has(file)) {
    return null;
  }
  return sourceText.includes("_docxParagraphSourceToken") ||
    sourceText.includes("PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR")
    ? `${file} references the private paragraph-source token`
    : null;
};

const contractOwnershipViolation = (file: string, sourceText: string): string | null => {
  if (!isProductionSource(file) || CONTRACT_OWNER_FILES.has(file)) {
    return null;
  }
  return sourceText.includes("_docxParagraphSourceContract") ||
    sourceText.includes("PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR")
    ? `${file} references the private paragraph-source contract`
    : null;
};

const scanPackageSources = (): { contract: string[]; token: string[] } => {
  const contract: string[] = [];
  const token: string[] = [];
  const sourcePaths = ts.sys.readDirectory(path.join(REPO_ROOT, "packages"), [
    ".ts",
    ".tsx",
    ".vue",
  ]);
  for (const sourcePath of sourcePaths) {
    const file = path.relative(REPO_ROOT, sourcePath).replaceAll("\\", "/");
    const sourceText = ts.sys.readFile(sourcePath);
    if (sourceText === undefined) {
      panic(`Cannot read ${file}.`);
    }
    const tokenViolation = tokenOwnershipViolation(file, sourceText);
    if (tokenViolation) {
      token.push(tokenViolation);
    }
    const contractViolation = contractOwnershipViolation(file, sourceText);
    if (contractViolation) {
      contract.push(contractViolation);
    }
  }
  return { contract: contract.toSorted(), token: token.toSorted() };
};

setDefaultTimeout(30_000);

describe("paragraph property source ownership", () => {
  test("recognizes every direct ProseMirror paragraph reconstruction form", () => {
    const sourceFile = ts.createSourceFile(
      "probe.ts",
      [
        "node.copy(content);",
        "transaction.setNodeMarkup(position, undefined, attrs);",
        "node.type.create(attrs, content, marks);",
        "node.create(attrs);",
      ].join("\n"),
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const rebuilds: ParagraphProvenanceRebuild[] = [];
    const visit = (node: ts.Node): void => {
      if (ts.isCallExpression(node)) {
        const rebuild = paragraphProvenanceRebuild(node);
        if (rebuild) {
          rebuilds.push(rebuild);
        }
      }
      ts.forEachChild(node, visit);
    };
    visit(sourceFile);

    expect(rebuilds).toEqual(["copy", "setNodeMarkup", "type.create"]);
  });

  test("every production story save supplies its property-source base", () => {
    expect(proseConversionViolations()).toEqual([]);
  });

  test("every typed clone and paragraph markup rebuild keeps an explicit source owner", () => {
    expect(cloneOwnershipViolations()).toEqual([]);
  });

  test("only the provenance kernel and lifecycle boundary can reference source metadata", () => {
    const violations = scanPackageSources();
    expect(violations.token).toEqual([]);
    expect(violations.contract).toEqual([]);
    expect(
      tokenOwnershipViolation(
        "packages/react/src/unsafe.ts",
        'attrs["_docxParagraphSourceToken"] = copied;',
      ),
    ).toContain("private paragraph-source token");
    expect(
      tokenOwnershipViolation(
        "packages/core/src/docx/paragraphPropertySource.ts",
        'attrs["_docxParagraphSourceToken"] = seeded;',
      ),
    ).toBeNull();
    expect(
      contractOwnershipViolation(
        "packages/react/src/unsafe.ts",
        'attrs["_docxParagraphSourceContract"] = copied;',
      ),
    ).toContain("private paragraph-source contract");
    expect(
      contractOwnershipViolation(
        "packages/core/src/prosemirror/yjsParagraphSourceContract.ts",
        'attrs["_docxParagraphSourceContract"] = seeded;',
      ),
    ).toBeNull();
  });
});
