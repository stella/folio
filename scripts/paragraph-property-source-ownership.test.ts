/** Static ownership guard for hidden paragraph-property captures. */

import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { existsSync, readFileSync, realpathSync } from "node:fs";
import path from "node:path";
import {
  type CallExpression,
  getTokenPosOfNode,
  isCallExpression,
  isIdentifier,
  isPropertyAccessExpression,
  isSpreadAssignment,
  type Node as CompilerNode,
  type SourceFile,
} from "tsrs/unstable/ast";
import {
  API,
  type Checker,
  type Type,
  TypeFlags,
  type UnionOrIntersectionType,
} from "tsrs/unstable/async";

import {
  createLineIndex,
  type Node,
  parseSourceModule,
  walkAst,
} from "./lib/oxc-ast";

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
  "packages/core/src/prosemirror/yjsDocumentMetadata.ts",
]);

/**
 * Package sources outside `node_modules`, in a stable order. The workspace
 * symlinks under `packages/<name>/node_modules/@stll` point back at the
 * packages, so a real path is listed once.
 */
const packageSources = (extensions: readonly string[]): string[] => {
  const root = path.join(REPO_ROOT, "packages");
  const seen = new Set<string>();
  const sources: string[] = [];
  const pattern = `**/*{${extensions.join(",")}}`;
  for (const file of new Bun.Glob(pattern).scanSync({ cwd: root })) {
    if (file.split("/").includes("node_modules")) continue;
    const absolute = path.join(root, file);
    const real = realpathSync(absolute);
    if (seen.has(real)) continue;
    seen.add(real);
    sources.push(absolute);
  }
  return sources.toSorted();
};

const relativePath = (fileName: string): string =>
  path.relative(REPO_ROOT, fileName).replaceAll("\\", "/");

const isUnionOrIntersection = (type: Type): type is UnionOrIntersectionType =>
  (type.flags & (TypeFlags.Union | TypeFlags.Intersection)) !== 0;

const containsParagraph = async (
  type: Type,
  checker: Checker,
): Promise<boolean> => {
  if (isUnionOrIntersection(type)) {
    for (const part of await type.getTypes()) {
      if (await containsParagraph(part, checker)) return true;
    }
    return false;
  }
  const discriminator = await type.getProperty("type");
  if (!discriminator) {
    return false;
  }
  const discriminatorType = await checker.getTypeOfSymbol(discriminator);
  return (
    (discriminatorType.flags & TypeFlags.StringLiteral) !== 0 &&
    Reflect.get(discriminatorType, "value") === "paragraph"
  );
};

type ParagraphProvenanceRebuild = "copy" | "setNodeMarkup" | "type.create";

const paragraphProvenanceRebuild = (
  node: CallExpression,
): ParagraphProvenanceRebuild | null => {
  if (!isPropertyAccessExpression(node.expression)) {
    return null;
  }
  if (
    node.expression.name.text === "copy" ||
    node.expression.name.text === "setNodeMarkup"
  ) {
    return node.expression.name.text;
  }
  if (node.expression.name.text !== "create") {
    return null;
  }
  const owner = node.expression.expression;
  return isPropertyAccessExpression(owner) && owner.name.text === "type"
    ? "type.create"
    : null;
};

const containsDocument = async (
  type: Type,
  checker: Checker,
): Promise<boolean> => {
  if (isUnionOrIntersection(type)) {
    for (const part of await type.getTypes()) {
      if (await containsDocument(part, checker)) return true;
    }
    return false;
  }
  const packageProperty = await type.getProperty("package");
  if (!packageProperty) {
    return false;
  }
  const packageType = await checker.getTypeOfSymbol(packageProperty);
  return (await packageType.getProperty("document")) !== undefined;
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

const importsProseDocToBlocks = (
  file: string,
  statements: readonly Node[],
): ProseDocToBlocksBindings => {
  const direct = new Set<string>();
  const namespaces = new Set<string>();
  if (file.endsWith("/prosemirror/conversion/fromProseDoc.ts")) {
    direct.add("proseDocToBlocks");
  }
  for (const statement of statements) {
    if (
      statement.type !== "ImportDeclaration" ||
      (!statement.source.value.endsWith("/fromProseDoc") &&
        !statement.source.value.endsWith("/prosemirror/conversion"))
    ) {
      continue;
    }
    for (const specifier of statement.specifiers) {
      if (specifier.type === "ImportNamespaceSpecifier") {
        namespaces.add(specifier.local.name);
        continue;
      }
      if (specifier.type !== "ImportSpecifier") {
        continue;
      }
      const imported =
        specifier.imported.type === "Identifier"
          ? specifier.imported.name
          : specifier.imported.value;
      if (imported === "proseDocToBlocks") {
        direct.add(specifier.local.name);
      }
    }
  }
  return { direct, namespaces };
};

const proseConversionViolations = (): string[] => {
  const violations: string[] = [];
  const sourcePaths = packageSources([".ts", ".tsx"]);
  for (const sourcePath of sourcePaths) {
    const file = path.relative(REPO_ROOT, sourcePath).replaceAll("\\", "/");
    if (!isProductionSource(file)) {
      continue;
    }
    const sourceText = readFileSync(sourcePath, "utf8");
    if (!sourceText.includes("proseDocToBlocks")) {
      continue;
    }
    const { program } = parseSourceModule(sourcePath, sourceText);
    const proseDocToBlocksBindings = importsProseDocToBlocks(
      file,
      program.body,
    );
    if (
      proseDocToBlocksBindings.direct.size === 0 &&
      proseDocToBlocksBindings.namespaces.size === 0
    ) {
      continue;
    }
    const lineAt = createLineIndex(sourceText);
    walkAst(program, (node) => {
      if (node.type !== "CallExpression") return;
      const callee = node.callee;
      const callsDirectBinding =
        callee.type === "Identifier" &&
        proseDocToBlocksBindings.direct.has(callee.name);
      const callsNamespaceBinding =
        callee.type === "MemberExpression" &&
        !callee.computed &&
        callee.property.type === "Identifier" &&
        callee.property.name === "proseDocToBlocks" &&
        callee.object.type === "Identifier" &&
        proseDocToBlocksBindings.namespaces.has(callee.object.name);
      if (
        (callsDirectBinding || callsNamespaceBinding) &&
        node.arguments.length < 2
      ) {
        const { line } = lineAt(node.start);
        violations.push(
          `${file}:${String(line + 1)} converts a story without its property-source base`,
        );
      }
    });
  }
  return violations;
};

const forEachNode = async (
  node: CompilerNode,
  visit: (node: CompilerNode) => Promise<void>,
): Promise<void> => {
  await visit(node);
  const children: CompilerNode[] = [];
  node.forEachChild((child) => {
    children.push(child);
  });
  for (const child of children) await forEachNode(child, visit);
};

const cloneOwnershipViolations = async (): Promise<string[]> => {
  const violations = new Set<string>();
  const visited = new Set<string>();
  const api = new API({ cwd: REPO_ROOT });
  try {
    for (const relativeConfigPath of TYPECHECK_CONFIGS) {
      const configPath = path.join(REPO_ROOT, relativeConfigPath);
      const snapshot = await api.createSnapshot({
        openProjects: [configPath],
      });
      const project =
        snapshot.getConfiguredProject(configPath) ??
        panic(`Cannot open ${relativeConfigPath}.`);
      const { checker, program } = project;

      for (const fileName of await program.getSourceFileNames()) {
        const file = relativePath(fileName);
        if (visited.has(file) || file === OWNER || !isProductionSource(file)) {
          continue;
        }
        visited.add(file);
        const sourceFile: SourceFile =
          (await program.getSourceFile(fileName)) ??
          panic(`Cannot read ${file}.`);
        const lineOf = (node: CompilerNode) =>
          sourceFile.getLineAndCharacterOfPosition(
            getTokenPosOfNode(node, sourceFile),
          ).line;
        await forEachNode(sourceFile, async (node) => {
          if (
            file.startsWith("packages/core/src/") &&
            isSpreadAssignment(node) &&
            (await containsParagraph(
              await checker.getTypeAtLocation(node.expression),
              checker,
            ))
          ) {
            violations.add(
              `${file}:${String(lineOf(node) + 1)} spreads a Paragraph`,
            );
          }
          if (
            isCallExpression(node) &&
            isIdentifier(node.expression) &&
            node.expression.text === "structuredClone"
          ) {
            for (const argument of node.arguments) {
              if (
                await containsDocument(
                  await checker.getTypeAtLocation(argument),
                  checker,
                )
              ) {
                violations.add(
                  `${file}:${String(lineOf(node) + 1)} calls structuredClone on a Document`,
                );
                break;
              }
            }
          }
          const provenanceRebuild = isCallExpression(node)
            ? paragraphProvenanceRebuild(node)
            : null;
          if (provenanceRebuild && PM_PARAGRAPH_PROVENANCE_FILES.has(file)) {
            violations.add(
              `${file}:${String(lineOf(node) + 1)} calls PM ${provenanceRebuild} directly`,
            );
          }
        });
      }
    }
  } finally {
    await api.close();
  }
  return [...violations].toSorted();
};

const tokenOwnershipViolation = (
  file: string,
  sourceText: string,
): string | null => {
  if (!isProductionSource(file) || TOKEN_OWNER_FILES.has(file)) {
    return null;
  }
  return sourceText.includes("_docxParagraphSourceToken") ||
    sourceText.includes("PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR")
    ? `${file} references the private paragraph-source token`
    : null;
};

const contractOwnershipViolation = (
  file: string,
  sourceText: string,
): string | null => {
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
  const sourcePaths = packageSources([".ts", ".tsx", ".vue"]);
  for (const sourcePath of sourcePaths) {
    const file = path.relative(REPO_ROOT, sourcePath).replaceAll("\\", "/");
    const sourceText = readFileSync(sourcePath, "utf8");
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
  test("names owners that exist, over sources that exist", () => {
    // A renamed owner leaves a name nothing matches, and the guard then reads
    // as "no violations" for a rule it has stopped enforcing.
    const named = [
      ...PM_PARAGRAPH_PROVENANCE_FILES,
      ...TOKEN_OWNER_FILES,
      ...CONTRACT_OWNER_FILES,
    ];
    expect(
      named.filter((file) => !existsSync(path.join(REPO_ROOT, file))),
    ).toEqual([]);

    const scanned = packageSources([".ts", ".tsx", ".vue"]).map((absolute) =>
      path.relative(REPO_ROOT, absolute).replaceAll("\\", "/"),
    );
    expect(scanned.filter((file) => file.includes("/node_modules/"))).toEqual(
      [],
    );
    expect(
      scanned.filter((file) => isProductionSource(file)).length,
    ).toBeGreaterThan(100);
  });

  test("recognizes every direct ProseMirror paragraph reconstruction form", async () => {
    const api = new API({ cwd: REPO_ROOT });
    try {
      const probe = await api.createSourceFile(
        "probe.ts",
        [
          "node.copy(content);",
          "transaction.setNodeMarkup(position, undefined, attrs);",
          "node.type.create(attrs, content, marks);",
          "node.create(attrs);",
        ].join("\n"),
      );
      const rebuilds: ParagraphProvenanceRebuild[] = [];
      await forEachNode(probe.sourceFile, async (node) => {
        if (isCallExpression(node)) {
          const rebuild = paragraphProvenanceRebuild(node);
          if (rebuild) {
            rebuilds.push(rebuild);
          }
        }
      });
      await probe.dispose();

      expect(rebuilds).toEqual(["copy", "setNodeMarkup", "type.create"]);
    } finally {
      await api.close();
    }
  });

  test("every production story save supplies its property-source base", () => {
    expect(proseConversionViolations()).toEqual([]);
  });

  test("every typed clone and paragraph markup rebuild keeps an explicit source owner", async () => {
    expect(await cloneOwnershipViolations()).toEqual([]);
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
        "packages/core/src/prosemirror/yjsDocumentMetadata.ts",
        'attrs["_docxParagraphSourceContract"] = seeded;',
      ),
    ).toBeNull();
  });
});
