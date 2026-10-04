// Keep root Playwright and parity harnesses from importing undeclared workspace
// packages at runtime. Node resolves these imports from the root manifest;
// a dependency declared by another workspace does not make it a root dependency.

import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

type AstNode = Record<string, unknown> & { type: string };
type WorkspacePolicy = { workspaceNames: ReadonlySet<string>; declaredNames: ReadonlySet<string> };

type RuleContext = {
  filename?: string;
  getFilename?: () => string;
  report: (descriptor: {
    node: unknown;
    messageId: "undeclaredWorkspaceRuntimeImport";
    data: { specifier: string; packageName: string };
  }) => void;
};

const PLUGIN_DIRECTORY = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(PLUGIN_DIRECTORY, "..");
const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);
let cachedPolicy: WorkspacePolicy | undefined;

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" && value !== null && "type" in value && typeof value.type === "string";

const visitChildren = (node: AstNode, visit: (child: AstNode) => void): void => {
  for (const [key, value] of Object.entries(node)) {
    if (SKIP_KEYS.has(key)) continue;
    if (Array.isArray(value)) {
      for (const item of value) if (isAstNode(item)) visit(item);
      continue;
    }
    if (isAstNode(value)) visit(value);
  }
};

const readJson = (filename: string): unknown => JSON.parse(readFileSync(filename, "utf8"));

const packageNameFromManifest = (filename: string): string => {
  const manifest = readJson(filename);
  if (typeof manifest !== "object" || manifest === null || !("name" in manifest)) {
    throw new TypeError(`Workspace manifest has no package name: ${filename}`);
  }
  const { name } = manifest;
  if (typeof name !== "string" || name.length === 0) {
    throw new TypeError(`Workspace manifest has an invalid package name: ${filename}`);
  }
  return name;
};

const declaredPackageNames = (manifest: unknown): ReadonlySet<string> => {
  if (typeof manifest !== "object" || manifest === null) {
    throw new TypeError("Root package manifest must be an object.");
  }
  const names = new Set<string>();
  for (const field of [
    "dependencies",
    "devDependencies",
    "peerDependencies",
    "optionalDependencies",
  ]) {
    const dependencies = Reflect.get(manifest, field);
    if (dependencies === undefined) continue;
    if (typeof dependencies !== "object" || dependencies === null || Array.isArray(dependencies)) {
      throw new TypeError(`Root package manifest field ${field} must be an object.`);
    }
    for (const name of Object.keys(dependencies)) names.add(name);
  }
  return names;
};

const workspacePolicy = (): WorkspacePolicy => {
  if (cachedPolicy !== undefined) return cachedPolicy;
  const packagesDirectory = path.join(REPO_ROOT, "packages");
  const workspaceNames = new Set(
    readdirSync(packagesDirectory, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => path.join(packagesDirectory, entry.name, "package.json"))
      .filter(existsSync)
      .map(packageNameFromManifest),
  );
  const rootManifest = readJson(path.join(REPO_ROOT, "package.json"));
  cachedPolicy = { workspaceNames, declaredNames: declaredPackageNames(rootManifest) };
  return cachedPolicy;
};

const stringLiteral = (node: unknown): string | null => {
  if (!isAstNode(node) || node.type !== "Literal" || typeof node["value"] !== "string") return null;
  return node["value"];
};

const packageNameOf = (specifier: string): string | null => {
  if (specifier.startsWith(".") || specifier.startsWith("/") || specifier.startsWith("#")) {
    return null;
  }
  const segments = specifier.split("/");
  if (specifier.startsWith("@"))
    return segments.length >= 2 ? `${segments[0]}/${segments[1]}` : null;
  return segments[0] ?? null;
};

const hasRuntimeImportSpecifiers = (node: AstNode): boolean => {
  const specifiers = node["specifiers"];
  if (!Array.isArray(specifiers) || specifiers.length === 0) return true;
  return specifiers.some((specifier) => {
    if (!isAstNode(specifier)) return true;
    return specifier["importKind"] !== "type" && specifier["exportKind"] !== "type";
  });
};

const moduleSpecifierOf = (node: AstNode): string | null => {
  if (node.type === "ImportDeclaration") {
    if (node["importKind"] === "type" || !hasRuntimeImportSpecifiers(node)) return null;
    return stringLiteral(node["source"]);
  }
  if (node.type === "ExportNamedDeclaration" || node.type === "ExportAllDeclaration") {
    if (node["exportKind"] === "type" || !hasRuntimeImportSpecifiers(node)) return null;
    return stringLiteral(node["source"]);
  }
  if (node.type === "ImportExpression") return stringLiteral(node["source"]);
  if (node.type === "CallExpression") {
    const callee = node["callee"];
    const args = node["arguments"];
    if (isAstNode(callee) && callee.type === "Import" && Array.isArray(args)) {
      return stringLiteral(args.at(0));
    }
    if (
      isAstNode(callee) &&
      callee.type === "Identifier" &&
      callee["name"] === "require" &&
      Array.isArray(args)
    ) {
      return stringLiteral(args.at(0));
    }
  }
  return null;
};

const checkImport = (node: AstNode, context: RuleContext): void => {
  const specifier = moduleSpecifierOf(node);
  if (specifier === null) return;
  const packageName = packageNameOf(specifier);
  if (packageName === null) return;
  const policy = workspacePolicy();
  if (!policy.workspaceNames.has(packageName) || policy.declaredNames.has(packageName)) return;
  context.report({
    node,
    messageId: "undeclaredWorkspaceRuntimeImport",
    data: { specifier, packageName },
  });
};

const filenameOf = (context: RuleContext): string =>
  typeof context.getFilename === "function" ? context.getFilename() : (context.filename ?? "");

const isHarnessFile = (filename: string): boolean => {
  const normalized = filename.replaceAll("\\", "/");
  return (
    normalized.includes("tests/visual/") ||
    normalized.includes("tests/parity/") ||
    /(?:^|\/)test\/__fixtures__\/harness-workspace\.(?:invalid|valid)\.ts$/u.test(normalized)
  );
};

export default {
  meta: { name: "folio-harness-workspaces" },
  rules: {
    "no-undeclared-workspace-runtime-import": {
      meta: {
        type: "problem",
        messages: {
          undeclaredWorkspaceRuntimeImport:
            "`{{specifier}}` imports workspace package `{{packageName}}` at runtime, but the root manifest does not declare it. Use a relative source import or declare the dependency.",
        },
      },
      create(context: RuleContext) {
        return {
          Program: (program: unknown) => {
            if (!isAstNode(program) || !isHarnessFile(filenameOf(context))) return;
            const visit = (node: AstNode): void => {
              if (
                node.type === "ImportDeclaration" ||
                node.type === "ExportNamedDeclaration" ||
                node.type === "ExportAllDeclaration" ||
                node.type === "ImportExpression" ||
                node.type === "CallExpression"
              ) {
                checkImport(node, context);
              }
              visitChildren(node, visit);
            };
            visit(program);
          },
        };
      },
    },
  },
};
